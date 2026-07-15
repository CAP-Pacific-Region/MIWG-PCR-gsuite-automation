/**
 * Cadet → senior mail migration.
 *
 * Runs on the CADETS tenant (TRANSITION_CONFIG.ROLE === 'source'), reading the
 * cadet mailbox with the local SA and writing into the senior mailbox with the
 * peer SA. See CadetTransition.gs for detection, state, and why the cadets
 * tenant owns the whole lifecycle.
 *
 * Both mailboxes are reached by domain-wide delegation, and the source account
 * is SUSPENDED by the time this runs. That is fine and verified: impersonation
 * works against suspended users, so the mailbox stays readable and nothing here
 * ever unsuspends an account.
 *
 * SCOPES — these live on the SA's DWD grant in the Admin console, NOT in
 * appsscript.json (that file governs the script's own OAuth, which impersonated
 * calls do not use):
 *   cadets tenant, local SA:  gmail.readonly
 *   seniors tenant, peer SA:  gmail.insert
 *
 * RESUMABILITY. Apps Script kills an execution at 6 minutes, and a four-year
 * cadet mailbox holds thousands of messages, so a single pass cannot finish.
 * Work stops cleanly at a page boundary before the limit, writes the Gmail
 * pageToken to LastCursor, and schedules itself to continue.
 *
 * DUPLICATES. The cursor advances only after a full page imports, so an
 * expected time-limit stop cannot duplicate anything. A hard crash mid-page
 * (quota, network) re-imports that page on resume, since Gmail does not dedup on
 * import. Page size is kept small to bound that to a handful of messages. The
 * alternative — searching the destination for each source Message-ID before
 * importing — doubles the API cost of every message to defend against a rare
 * case, and duplicates are recoverable where lost mail is not.
 */

/** Messages per page. Small on purpose: caps duplicates from a mid-page crash. */
const MIGRATE_PAGE_SIZE_ = 25;

/**
 * Stop and reschedule at this much wall time. Apps Script's ceiling is 6
 * minutes; leaving ~90s of headroom means a slow final page still lands inside
 * the limit rather than being killed mid-import.
 */
const MIGRATE_SOFT_TIME_LIMIT_MS_ = 4.5 * 60 * 1000;

/** Label applied to migrated mail in the destination mailbox. */
const MIGRATE_LABEL_NAME_ = 'Migrated from cadets';

/** Continuation trigger handler name. */
const MIGRATE_CONTINUATION_FN_ = 'continueCadetTransitionMigration';

// ============================================================================
// ENTRY POINTS
// ============================================================================

/**
 * Migrates mail for every transition row with a resolved destination.
 *
 * Processes one member to completion before starting the next, so an interrupted
 * run leaves a clear picture: earlier rows COMPLETE, one IN_PROGRESS with a
 * cursor, later rows still PENDING.
 *
 * @returns {{migrated: number, incomplete: number, failed: number}}
 */
function migrateCadetTransitions() {
  if (TRANSITION_CONFIG.ROLE !== 'source') {
    Logger.info('Transition migration skipped — not the source tenant');
    return { migrated: 0, incomplete: 0, failed: 0 };
  }

  const started = new Date();
  const rows = readTransitions_();
  let migrated = 0;
  let incomplete = 0;
  let failed = 0;

  for (const capid in rows) {
    const row = rows[capid];

    if (!isMigratable_(row)) continue;

    if (new Date() - started > MIGRATE_SOFT_TIME_LIMIT_MS_) {
      Logger.info('Migration paused — time limit, remaining rows next run', { capid: capid });
      scheduleMigrationContinuation_();
      incomplete++;
      break;
    }

    try {
      const result = migrateOneTransition_(row, started);
      if (result.complete) migrated++; else incomplete++;
    } catch (e) {
      // FAILED is terminal until a human looks. It also blocks deletion — see
      // getHeldTransitionCapids() — because a failed migration followed by an
      // on-schedule deletion is the one outcome that loses mail for good.
      setTransitionField_(row._rowNumber, 'MigrationStatus', TRANSITION_CONFIG.STATUS.FAILED);
      setTransitionField_(row._rowNumber, 'Notes',
        `Failed ${new Date().toISOString()}: ${e && e.message ? e.message : String(e)}`);
      Logger.error('Transition migration failed', {
        capid: capid,
        cadetEmail: row.CadetEmail,
        seniorEmail: row.SeniorEmail,
        errorMessage: e && e.message ? e.message : String(e)
      });
      failed++;
    }
  }

  Logger.info('Transition migration pass completed', {
    duration: new Date() - started + 'ms',
    migrated: migrated,
    incomplete: incomplete,
    failed: failed
  });

  return { migrated: migrated, incomplete: incomplete, failed: failed };
}

/**
 * Continuation trigger target. Deletes the one-shot trigger that fired it, then
 * resumes.
 *
 * @param {Object} e - trigger event
 */
function continueCadetTransitionMigration(e) {
  if (e && e.triggerUid) {
    ScriptApp.getProjectTriggers()
      .filter(t => t.getUniqueId() === e.triggerUid)
      .forEach(t => ScriptApp.deleteTrigger(t));
  }
  migrateCadetTransitions();
}

/**
 * A row is migratable once it has a destination and is not already done.
 *
 * PENDING rows without a SeniorEmail are skipped rather than treated as an
 * error: a PATRON has no senior account to migrate to, and a recent SENIOR may
 * simply not have been provisioned yet. resolveTransitionDestinations() fills
 * that in when it appears.
 *
 * @param {Object} row
 * @returns {boolean}
 */
function isMigratable_(row) {
  if (!row.SeniorEmail) return false;
  return row.MigrationStatus === TRANSITION_CONFIG.STATUS.PENDING ||
         row.MigrationStatus === TRANSITION_CONFIG.STATUS.IN_PROGRESS;
}

// ============================================================================
// PER-MEMBER MIGRATION
// ============================================================================

/**
 * Migrates one member's mail, resuming from LastCursor if set.
 *
 * @param {Object} row - Transitions row
 * @param {Date} started - pass start, for the shared time budget
 * @returns {{complete: boolean, imported: number}}
 */
function migrateOneTransition_(row, started) {
  const cfg = getCrossTenantConfig_();

  const sourceToken = getImpersonatedToken_(
    row.CadetEmail, 'https://www.googleapis.com/auth/gmail.readonly'
  );
  const destToken = xtPeerToken_(
    'https://www.googleapis.com/auth/gmail.insert', cfg
  );

  const labelId = ensureDestinationLabel_(row.SeniorEmail, destToken);

  setTransitionField_(row._rowNumber, 'MigrationStatus', TRANSITION_CONFIG.STATUS.IN_PROGRESS);

  let pageToken = row.LastCursor || '';
  let imported = Number(row.MessagesMigrated) || 0;

  Logger.info('Migrating mailbox', {
    capid: row.CAPID,
    from: row.CadetEmail,
    to: row.SeniorEmail,
    resumingAt: pageToken ? pageToken.substring(0, 12) + '…' : '(start)',
    alreadyImported: imported
  });

  do {
    if (new Date() - started > MIGRATE_SOFT_TIME_LIMIT_MS_) {
      setTransitionField_(row._rowNumber, 'LastCursor', pageToken);
      setTransitionField_(row._rowNumber, 'MessagesMigrated', imported);
      scheduleMigrationContinuation_();
      Logger.info('Mailbox migration paused — will resume', {
        capid: row.CAPID, imported: imported
      });
      return { complete: false, imported: imported };
    }

    const page = listSourceMessages_(row.CadetEmail, sourceToken, pageToken);

    for (const msg of page.messages) {
      const raw = getSourceMessageRaw_(row.CadetEmail, sourceToken, msg.id);
      importToDestination_(row.SeniorEmail, destToken, raw, labelId);
      imported++;
    }

    // Only now is the page fully imported, so the cursor may safely advance.
    pageToken = page.nextPageToken || '';
    setTransitionField_(row._rowNumber, 'LastCursor', pageToken);
    setTransitionField_(row._rowNumber, 'MessagesMigrated', imported);

  } while (pageToken);

  setTransitionField_(row._rowNumber, 'MigrationStatus', TRANSITION_CONFIG.STATUS.COMPLETE);
  setTransitionField_(row._rowNumber, 'MigratedDate', new Date().toISOString());
  setTransitionField_(row._rowNumber, 'LastCursor', '');

  Logger.info('Mailbox migration complete', {
    capid: row.CAPID,
    from: row.CadetEmail,
    to: row.SeniorEmail,
    imported: imported
  });

  return { complete: true, imported: imported };
}

// ============================================================================
// GMAIL
// ============================================================================

/**
 * Lists one page of source message ids.
 *
 * includeSpamTrash is false: spam and trash are not worth carrying across, and
 * importing known spam into the destination risks its reputation.
 *
 * @param {string} user
 * @param {string} token
 * @param {string} pageToken
 * @returns {{messages: Array<{id: string}>, nextPageToken: string}}
 */
function listSourceMessages_(user, token, pageToken) {
  const url = 'https://gmail.googleapis.com/gmail/v1/users/me/messages' +
    '?maxResults=' + MIGRATE_PAGE_SIZE_ +
    '&includeSpamTrash=false' +
    (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

  const body = gmailFetch_(url, { method: 'get', token: token }, 'list messages for ' + user);
  return { messages: body.messages || [], nextPageToken: body.nextPageToken || '' };
}

/**
 * Fetches one message as raw RFC822.
 *
 * @param {string} user
 * @param {string} token
 * @param {string} id
 * @returns {string} base64url-encoded RFC822
 */
function getSourceMessageRaw_(user, token, id) {
  const url = 'https://gmail.googleapis.com/gmail/v1/users/me/messages/' +
    encodeURIComponent(id) + '?format=raw';

  const body = gmailFetch_(url, { method: 'get', token: token }, 'get message ' + id);
  return body.raw;
}

/**
 * Imports one raw message into the destination mailbox.
 *
 * internalDateSource=dateHeader keeps the original sent date, so migrated mail
 * sorts correctly instead of all landing at the migration timestamp.
 * neverMarkSpam stops the destination's filters from burying mail the member
 * already received and read once.
 *
 * @param {string} user
 * @param {string} token
 * @param {string} raw - base64url RFC822
 * @param {string} labelId
 */
function importToDestination_(user, token, raw, labelId) {
  const url = 'https://gmail.googleapis.com/gmail/v1/users/' +
    encodeURIComponent(user) + '/messages/import' +
    '?internalDateSource=dateHeader&neverMarkSpam=true';

  gmailFetch_(url, {
    method: 'post',
    token: token,
    payload: JSON.stringify({ raw: raw, labelIds: labelId ? [labelId] : undefined })
  }, 'import message to ' + user);
}

/**
 * Finds or creates the migration label on the destination mailbox.
 *
 * Source labels are not carried across: recreating a cadet's whole label tree in
 * the senior mailbox is a much larger job, and one label makes the provenance
 * obvious and the whole import reversible with a single search. Returns '' on
 * failure — a missing label is not worth failing a migration over.
 *
 * @param {string} user
 * @param {string} token
 * @returns {string} label id, or ''
 */
function ensureDestinationLabel_(user, token) {
  const base = 'https://gmail.googleapis.com/gmail/v1/users/' + encodeURIComponent(user) + '/labels';

  try {
    const existing = gmailFetch_(base, { method: 'get', token: token }, 'list labels for ' + user);
    const found = (existing.labels || []).find(l => l.name === MIGRATE_LABEL_NAME_);
    if (found) return found.id;

    const created = gmailFetch_(base, {
      method: 'post',
      token: token,
      payload: JSON.stringify({
        name: MIGRATE_LABEL_NAME_,
        labelListVisibility: 'labelShow',
        messageListVisibility: 'show'
      })
    }, 'create label for ' + user);

    Logger.info('Migration label created', { user: user, labelId: created.id });
    return created.id;

  } catch (e) {
    Logger.warn('Could not resolve migration label — importing unlabelled', {
      user: user,
      errorMessage: e && e.message ? e.message : String(e)
    });
    return '';
  }
}

/** Retry budget for a single Gmail call. */
const MIGRATE_MAX_ATTEMPTS_ = 5;

/** First backoff wait; doubles per attempt. */
const MIGRATE_BACKOFF_BASE_MS_ = 2000;

/**
 * One Gmail REST call, with retry on transient failures.
 *
 * Deliberately NOT the shared executeWithRetry(). That helper is built for
 * low-volume Admin Directory work and does two things that are wrong here:
 *
 *  1. It sleeps CONFIG.API_DELAY_MS (3s) before EVERY call. This loop makes two
 *     Gmail calls per message, so that is 6s of pure sleep per message — a
 *     1800-message mailbox would take ~3 hours across ~40 continuation runs
 *     instead of ~4.
 *  2. It detects transient errors via e.details?.code, which only exists on
 *     errors thrown by Apps Script's advanced services. These are raw
 *     UrlFetchApp calls with muteHttpExceptions, so the status lives on the
 *     response, not on an exception — every error would read as code 0, be
 *     classed non-transient, and rethrow immediately. The 3s tax, no retry.
 *
 * So: read the status off the response, back off on 429/5xx, fail fast on
 * everything else (a 403 here means a missing DWD scope — retrying cannot help
 * and only delays a clear error).
 *
 * @param {string} url
 * @param {{method: string, token: string, payload: string=}} opts
 * @param {string} what - description for error messages
 * @returns {Object} parsed response
 */
function gmailFetch_(url, opts, what) {
  const params = {
    method: opts.method,
    headers: { Authorization: 'Bearer ' + opts.token },
    muteHttpExceptions: true
  };
  if (opts.payload) {
    params.contentType = 'application/json';
    params.payload = opts.payload;
  }

  let lastError = '';

  for (let attempt = 1; attempt <= MIGRATE_MAX_ATTEMPTS_; attempt++) {
    const resp = UrlFetchApp.fetch(url, params);
    const code = resp.getResponseCode();
    const text = resp.getContentText();

    if (code >= 200 && code < 300) {
      return text ? JSON.parse(text) : {};
    }

    lastError = `Gmail ${what} failed (${code}): ${text}`;

    // 429 = rate limited, 5xx = transient. Everything else is ours to fix.
    const transient = code === 429 || (code >= 500 && code < 600);
    if (!transient || attempt === MIGRATE_MAX_ATTEMPTS_) {
      throw new Error(lastError);
    }

    const delay = MIGRATE_BACKOFF_BASE_MS_ * Math.pow(2, attempt - 1);
    Logger.warn('Gmail call retrying after transient error', {
      what: what, attempt: attempt, code: code, waitMs: delay
    });
    Utilities.sleep(delay);
  }

  throw new Error(lastError);
}

// ============================================================================
// CONTINUATION
// ============================================================================

/**
 * Schedules a one-shot continuation a minute out.
 *
 * Guarded against stacking: an interrupted run can pause at both the row loop
 * and inside a mailbox, and two triggers racing would import the same pages
 * twice.
 */
function scheduleMigrationContinuation_() {
  const already = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === MIGRATE_CONTINUATION_FN_);

  if (already) {
    Logger.info('Migration continuation already scheduled');
    return;
  }

  ScriptApp.newTrigger(MIGRATE_CONTINUATION_FN_)
    .timeBased()
    .after(60 * 1000)
    .create();

  Logger.info('Migration continuation scheduled');
}

/** Removes any pending continuation triggers. Manual recovery aid. */
function clearMigrationContinuations() {
  let removed = 0;
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === MIGRATE_CONTINUATION_FN_)
    .forEach(t => { ScriptApp.deleteTrigger(t); removed++; });

  Logger.info('Migration continuations cleared', { removed: removed });
  return { removed: removed };
}

// ============================================================================
// PREVIEW
// ============================================================================

/**
 * Read-only. Reports what migration would move, and proves both credentials
 * work, without importing anything.
 *
 * Worth running before the first real migration: it exercises the source
 * gmail.readonly grant and the peer gmail.insert grant on every affected
 * mailbox, so a missing DWD scope surfaces here rather than halfway through
 * someone's mail.
 */
function previewCadetTransitionMigration() {
  const rows = readTransitions_();
  const cfg = getCrossTenantConfig_();
  let total = 0;

  for (const capid in rows) {
    const row = rows[capid];
    if (!isMigratable_(row)) {
      console.log(`${capid} | ${row.Name} | ${row.NewType} | skipped — ` +
        (row.SeniorEmail ? 'status ' + row.MigrationStatus : 'no destination account'));
      continue;
    }

    try {
      const srcToken = getImpersonatedToken_(
        row.CadetEmail, 'https://www.googleapis.com/auth/gmail.readonly'
      );
      const profile = gmailFetch_(
        'https://gmail.googleapis.com/gmail/v1/users/me/profile',
        { method: 'get', token: srcToken },
        'profile for ' + row.CadetEmail
      );

      // Prove the destination credential too — a token minted here beats
      // discovering the grant is missing mid-import.
      xtPeerToken_('https://www.googleapis.com/auth/gmail.insert', cfg);

      console.log(`${capid} | ${row.Name} | ${row.NewType} | ` +
        `${row.CadetEmail} -> ${row.SeniorEmail} | ${profile.messagesTotal} messages`);
      total += Number(profile.messagesTotal) || 0;

    } catch (e) {
      console.log(`${capid} | ${row.Name} | ERROR: ${e && e.message ? e.message : String(e)}`);
    }
  }

  console.log('');
  console.log('Total messages to migrate: ' + total);
  console.log('Nothing was imported — this is a preview.');
}
