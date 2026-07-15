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

/** Scope needed on the SOURCE (cadets) SA's DWD grant. */
const MIGRATE_SOURCE_SCOPE_ = 'https://www.googleapis.com/auth/gmail.readonly';

/**
 * Scopes needed on the PEER (seniors) SA's DWD grant, space separated.
 *
 * Two are required and it is not obvious why: gmail.insert authorizes
 * messages.import but NOT the labels API, and gmail.labels authorizes
 * labels.list/create but NOT importing. Requesting only the first silently
 * costs the label — ensureDestinationLabel_() catches the failure and imports
 * unlabelled, so it looks like it worked.
 */
const MIGRATE_DEST_SCOPES_ =
  'https://www.googleapis.com/auth/gmail.insert ' +
  'https://www.googleapis.com/auth/gmail.labels';

/**
 * Read scope for INSPECTING the destination — deliberately separate from
 * MIGRATE_DEST_SCOPES_ and deliberately minimal.
 *
 * gmail.metadata grants headers, counts and profile but NOT message bodies,
 * which is all a duplicate check needs. gmail.readonly would also work and would
 * additionally hand the peer SA the contents of every senior mailbox — too much
 * standing privilege for a sanity check, especially on an SA whose key is still
 * unrotated.
 *
 * Kept out of the migration token on purpose: a token requesting scopes that are
 * not ALL granted fails outright, so folding this in would mean a missing
 * metadata grant breaks migration itself. Inspection fails soft instead.
 */
const MIGRATE_INSPECT_SCOPE_ = 'https://www.googleapis.com/auth/gmail.metadata';

/**
 * Block migration if the destination already holds at least this share of the
 * source's message count — the signature of mail already moved by other means.
 *
 * A senior mailbox is never empty (welcome mail, whatever has arrived since), so
 * an absolute count proves nothing; the ratio is what matters. Paired with
 * MIGRATE_DUPLICATE_GUARD_MIN_ so a nearly-empty source cannot trip it.
 */
const MIGRATE_DUPLICATE_GUARD_RATIO_ = 0.5;

/** Destination must hold at least this many messages before the ratio applies. */
const MIGRATE_DUPLICATE_GUARD_MIN_ = 100;

/** Continuation trigger handler name. */
const MIGRATE_CONTINUATION_FN_ = 'continueCadetTransitionMigration';

/**
 * Script Property carrying what the pending continuation is scoped to.
 *
 * A trigger cannot take arguments, so without this the continuation has no idea
 * what the run that scheduled it was doing — and would default to "everything,
 * with notifications", silently escalating a deliberately narrow run into a
 * broad one. Holds {capid?, notify}.
 */
const MIGRATE_CONTINUATION_SCOPE_PROP_ = 'MIGRATE_CONTINUATION_SCOPE';

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
 * @param {boolean} [notify=true] - send completion emails
 * @returns {{migrated: number, incomplete: number, failed: number}}
 */
function migrateCadetTransitions(notify) {
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
      scheduleMigrationContinuation_({ notify: notify });
      incomplete++;
      break;
    }

    try {
      const result = migrateOneTransition_(row, started, { notify: notify });
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
 * Migrates ONE member, by CAPID. Everything else is left alone.
 *
 * For proving the pipeline on a single real mailbox before turning it loose on
 * the queue. Identical code path to migrateCadetTransitions() — same import,
 * same resumability, same completion email — just scoped to one row, so a defect
 * costs one mailbox to unpick instead of several.
 *
 * @param {string} capid
 * @param {boolean} [notify=true] - pass false to move the mail without emailing
 *   the member and their commander, for a first run you want to inspect quietly
 * @returns {{complete: boolean, imported: number}}
 */
function migrateSingleTransition(capid, notify) {
  if (TRANSITION_CONFIG.ROLE !== 'source') {
    throw new Error('Not the source tenant');
  }

  const rows = readTransitions_();
  const row = rows[String(capid)];
  if (!row) throw new Error('No transition row for CAPID ' + capid);

  if (!isMigratable_(row)) {
    throw new Error(`CAPID ${capid} is not migratable: ` +
      (row.SeniorEmail ? 'status is ' + row.MigrationStatus : 'no destination account'));
  }

  Logger.info('Single-member migration starting', {
    capid: capid, name: row.Name, from: row.CadetEmail, to: row.SeniorEmail,
    notify: notify !== false
  });

  const started = new Date();
  try {
    return migrateOneTransition_(row, started, {
      notify: notify !== false,
      single: true          // so a continuation resumes THIS member, not the queue
    });
  } catch (e) {
    setTransitionField_(row._rowNumber, 'MigrationStatus', TRANSITION_CONFIG.STATUS.FAILED);
    setTransitionField_(row._rowNumber, 'Notes',
      `Failed ${new Date().toISOString()}: ${e && e.message ? e.message : String(e)}`);
    throw e;
  }
}

/**
 * Continuation trigger target. Deletes the trigger that fired it, then resumes
 * whatever the previous run was doing — at the SAME scope and notify setting.
 *
 * Triggers take no arguments, so the scope comes from a Script Property. This
 * used to call migrateCadetTransitions() unconditionally, which meant a
 * deliberately narrow run — one member, notifications off — silently became the
 * whole queue with notifications on at its first pause. The narrower the run you
 * asked for, the more surprising the result.
 *
 * @param {Object} e - trigger event
 */
function continueCadetTransitionMigration(e) {
  if (e && e.triggerUid) {
    ScriptApp.getProjectTriggers()
      .filter(t => t.getUniqueId() === e.triggerUid)
      .forEach(t => ScriptApp.deleteTrigger(t));
  }

  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(MIGRATE_CONTINUATION_SCOPE_PROP_);
  props.deleteProperty(MIGRATE_CONTINUATION_SCOPE_PROP_);

  let scope = null;
  try {
    scope = raw ? JSON.parse(raw) : null;
  } catch (err) {
    Logger.warn('Unparseable continuation scope', { raw: raw });
  }

  if (!scope) {
    // No scope: either a trigger scheduled by an older version, or the property
    // was lost. We cannot know what the original run intended, so do the least
    // surprising thing — finish what is demonstrably already underway and start
    // nothing new. Notifications off, because a mailbox someone else began
    // migrating is not ours to announce.
    Logger.warn('Continuation has no scope — resuming in-progress rows only, ' +
      'starting nothing new, notifications suppressed');
    resumeInProgressOnly_();
    return;
  }

  if (scope.capid) {
    migrateSingleTransition(scope.capid, scope.notify);
  } else {
    migrateCadetTransitions(scope.notify);
  }
}

/**
 * Resumes only rows already mid-migration. Starts nothing, notifies nobody.
 *
 * The safe fallback when a continuation cannot tell what it was for.
 *
 * @returns {{resumed: number}}
 */
function resumeInProgressOnly_() {
  const started = new Date();
  const rows = readTransitions_();
  let resumed = 0;

  for (const capid in rows) {
    const row = rows[capid];
    if (row.MigrationStatus !== TRANSITION_CONFIG.STATUS.IN_PROGRESS) continue;
    if (!row.SeniorEmail) continue;

    try {
      migrateOneTransition_(row, started, { notify: false });
      resumed++;
    } catch (e) {
      setTransitionField_(row._rowNumber, 'MigrationStatus', TRANSITION_CONFIG.STATUS.FAILED);
      setTransitionField_(row._rowNumber, 'Notes',
        `Failed ${new Date().toISOString()}: ${e && e.message ? e.message : String(e)}`);
      Logger.error('Resume failed', {
        capid: capid, errorMessage: e && e.message ? e.message : String(e)
      });
    }
    break; // one at a time; the next pause reschedules
  }

  Logger.info('In-progress resume completed', { resumed: resumed });
  return { resumed: resumed };
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
 * @param {Object} [opts]
 * @param {string} [opts.query] - Gmail search limiting the set; the catch-up
 *   pass uses it to pick up only what arrived since the last run
 * @param {boolean} [opts.notify=true] - send the completion email
 * @param {boolean} [opts.single] - this run is scoped to THIS member only, so a
 *   continuation must resume only them rather than the whole queue
 * @returns {{complete: boolean, imported: number}}
 */
function migrateOneTransition_(row, started, opts) {
  opts = opts || {};
  const query = opts.query || '';
  const notify = opts.notify;

  const cfg = getCrossTenantConfig_();

  const sourceToken = getImpersonatedToken_(
    row.CadetEmail, 'https://www.googleapis.com/auth/gmail.readonly'
  );

  // Impersonate the MEMBER on the peer side, not the peer admin. Gmail has no
  // admin-level cross-user access — an admin token cannot write into someone
  // else's mailbox — so the third argument is what makes this work at all.
  const destToken = xtPeerToken_(MIGRATE_DEST_SCOPES_, cfg, row.SeniorEmail);

  // Duplicate guard — only on a FRESH start.
  //
  // Skipped on catch-up, where the destination is expected to already hold the
  // bulk of the mail — that is the point of that pass.
  //
  // Skipped on resume, and that one is not obvious: this function is re-entered
  // once per continuation, and a large mailbox takes ~18 of them. Partway
  // through, the destination legitimately holds half the source, which is
  // exactly what the guard looks for — so it would kill a migration that is
  // working correctly, the further along the more certainly. A row with a cursor
  // or a nonzero count has already passed this check on its first run.
  const isResume = !!row.LastCursor || !!(Number(row.MessagesMigrated) || 0);

  if (!query && !isResume) {
    const destCount = destinationMessageCount_(row.SeniorEmail, cfg);
    const srcCount = sourceMessageCount_(row.CadetEmail);

    if (looksAlreadyMigrated_(destCount, srcCount)) {
      throw new Error(
        `Destination already holds ${destCount} messages against a source of ` +
        `${srcCount} — this mail looks like it has already been moved. Importing ` +
        `would duplicate it. If this is wrong, migrate with the guard bypassed; ` +
        `if it is right, set MigrationStatus to COMPLETE and skip this row.`
      );
    }
    if (destCount !== null) {
      Logger.info('Duplicate guard passed', {
        capid: row.CAPID, destCount: destCount, sourceCount: srcCount
      });
    }
  }

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

      // Carry the scope forward, or the continuation widens the run.
      scheduleMigrationContinuation_(
        opts.single
          ? { capid: String(row.CAPID), notify: notify }
          : { notify: notify }
      );

      Logger.info('Mailbox migration paused — will resume', {
        capid: row.CAPID, imported: imported, scope: opts.single ? 'this member' : 'queue'
      });
      return { complete: false, imported: imported };
    }

    const page = listSourceMessages_(row.CadetEmail, sourceToken, pageToken, query);

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

  const migratedAt = new Date();

  // A query means this is a catch-up sweep over a mailbox already migrated, not
  // a first migration. Two things must NOT repeat:
  //
  //  - DeleteAfter must not move. Catch-up runs immediately before deletion, so
  //    re-stamping migratedAt + 14 would push the deadline out by another 14
  //    days every single time and the account would never actually close.
  //  - The completion email must not re-send. The member was told once; telling
  //    them again every sweep is noise, and it CCs their commander.
  const isCatchUp = !!query;

  setTransitionField_(row._rowNumber, 'MigrationStatus', TRANSITION_CONFIG.STATUS.COMPLETE);
  setTransitionField_(row._rowNumber, 'MigratedDate', migratedAt.toISOString());
  setTransitionField_(row._rowNumber, 'MessagesMigrated', imported);
  setTransitionField_(row._rowNumber, 'LastCursor', '');

  if (isCatchUp) {
    Logger.info('Catch-up sweep complete', {
      capid: row.CAPID, imported: imported
    });
    return { complete: true, imported: imported };
  }

  // The 90-day hold from detection was there to wait out National's fingerprint
  // processing — i.e. for members who might yet convert. This one has converted
  // and their mail is across, so that purpose is served. Pull the deletion in:
  // the forwarding group cannot take the old address until the account is gone,
  // so a longer wait just keeps the member in limbo for no benefit. The
  // remaining buffer is to catch a migration that reported success but was not.
  const deleteAfter = new Date(migratedAt);
  deleteAfter.setDate(deleteAfter.getDate() + TRANSITION_CONFIG.POST_MIGRATION_DELETE_DAYS);
  setTransitionField_(row._rowNumber, 'DeleteAfter', deleteAfter.toISOString());

  Logger.info('Mailbox migration complete', {
    capid: row.CAPID,
    from: row.CadetEmail,
    to: row.SeniorEmail,
    imported: imported,
    deleteAfter: deleteAfter.toISOString()
  });

  if (notify === false) {
    Logger.info('Notification suppressed for this run', { capid: row.CAPID });
    return { complete: true, imported: imported };
  }

  // Send-failure must not undo a good migration, so this is caught rather than
  // thrown: the mail is already safely across, and marking the row FAILED over
  // an email would block the deletion and re-run the whole import.
  try {
    sendTransitionCompleteEmail_(row, imported, deleteAfter);
    setTransitionField_(row._rowNumber, 'NotifiedDate', new Date().toISOString());
  } catch (e) {
    Logger.error('Transition complete, but notification failed', {
      capid: row.CAPID,
      errorMessage: e && e.message ? e.message : String(e)
    });
    setTransitionField_(row._rowNumber, 'Notes',
      `Migrated OK; notification failed ${new Date().toISOString()}: ${e && e.message ? e.message : String(e)}`);
  }

  return { complete: true, imported: imported };
}

// ============================================================================
// NOTIFICATION
// ============================================================================

/**
 * Tells the member their mail has moved and their old account is closing.
 *
 * Goes to three addresses, because the whole value of this email is reaching
 * them in time to warn their contacts:
 *  - the new senior mailbox, which they may not have logged into yet
 *  - their CAPWATCH personal address
 *  - their OLD cadet address, which is where they are most likely still reading
 *
 * The cadet address is included precisely because it is still live. An earlier
 * version excluded it on the grounds that a suspended account cannot receive —
 * but on this tenant these accounts are not suspended at all (suspension is not
 * running here), so mail sent there is delivered and read. Excluding the one
 * address they actually check would have defeated the point.
 *
 * Commander is CC'd, as the retention emails do.
 *
 * @param {Object} row - Transitions row
 * @param {number} messageCount - messages imported
 * @param {Date} deleteAfter - when the cadet account will be deleted
 */
function sendTransitionCompleteEmail_(row, messageCount, deleteAfter) {
  const capid = String(row.CAPID);

  // Rank/orgid are not on the sheet — pull them from CAPWATCH at send time
  // rather than widening the schema for two fields used in one place.
  const memberData = parseFile('Member');
  let rank = '';
  let lastName = row.Name || '';
  let orgid = '';
  for (let i = 0; i < memberData.length; i++) {
    if (String(memberData[i][0] || '').trim() === capid) {
      lastName = memberData[i][2] || lastName;
      rank = memberData[i][14] || '';
      orgid = memberData[i][11] || '';
      break;
    }
  }

  const personalEmail = createEmailMap()[capid] || '';
  const commander = orgid ? getCommanderInfo(orgid) : null;

  const recipients = [row.SeniorEmail, personalEmail, row.CadetEmail]
    .filter(Boolean)
    .filter((addr, i, all) => all.indexOf(addr) === i);

  if (!recipients.length) {
    throw new Error('No deliverable address for CAPID ' + capid);
  }

  // Full slash-prefixed name: subfoldered templates deploy with the folder in
  // the literal filename, and HtmlService needs it exactly.
  const htmlBody = HtmlService.createHtmlOutputFromFile(
    'recruiting-and-retention/TransitionCompleteEmail'
  )
    .getContent()
    .replace(/{{rank}}/g, rank || 'Senior Member')
    .replace(/{{lastName}}/g, lastName)
    .replace(/{{seniorEmail}}/g, row.SeniorEmail)
    .replace(/{{cadetEmail}}/g, row.CadetEmail)
    .replace(/{{messageCount}}/g, String(messageCount))
    .replace(/{{deleteDate}}/g, formatTransitionDate_(deleteAfter));

  executeWithRetry(() =>
    GmailApp.sendEmail(
      recipients.join(','),
      TRANSITION_CONFIG.EMAIL_SUBJECT,
      htmlBody,
      {
        htmlBody: htmlBody,
        cc: commander && commander.email ? commander.email : '',
        bcc: ITSUPPORT_EMAIL,
        replyTo: ITSUPPORT_EMAIL,
        from: AUTOMATION_SENDER_EMAIL,
        name: SENDER_NAME
      }
    )
  );

  Logger.info('Transition complete email sent', {
    capid: capid,
    to: recipients.join(','),
    commanderCc: commander && commander.email ? commander.email : 'none',
    messageCount: messageCount
  });
}

/**
 * Sends the completion email to migrated members who have not had it yet.
 *
 * Exists because migration and notification have to be separable. Mail can be
 * moved long before the deletion machinery is ready to honour the date the email
 * quotes, and an email promising "deleted on X, forwarding after" is a lie until
 * deletion and the forwarding group actually exist. So migrations run quietly
 * with notify=false, and this sends once the dates mean something.
 *
 * Driven off NotifiedDate rather than a status, so it cannot double-send: a row
 * with a date is done, blank means untold. Skips rows whose Notes carry a
 * catch-up failure — telling someone their mail is safely across while a
 * DO NOT DELETE marker sits on the row would be exactly backwards.
 *
 * @param {boolean} [dryRun=true] - report who would be mailed, send nothing
 * @returns {{sent: number, skipped: number, failed: number}}
 */
function notifyCompletedTransitions(dryRun) {
  if (TRANSITION_CONFIG.ROLE !== 'source') {
    Logger.info('Notification skipped — not the source tenant');
    return { sent: 0, skipped: 0, failed: 0 };
  }

  const isDry = dryRun !== false;   // default to dry run: this mails real people
  const rows = readTransitions_();
  let sent = 0;
  let skipped = 0;
  let failed = 0;

  for (const capid in rows) {
    const row = rows[capid];

    if (row.MigrationStatus !== TRANSITION_CONFIG.STATUS.COMPLETE) { skipped++; continue; }
    if (row.NotifiedDate) { skipped++; continue; }
    if (!row.SeniorEmail) { skipped++; continue; }

    if (String(row.Notes || '').indexOf('DO NOT DELETE') > -1) {
      console.log(`${capid} | ${row.Name} | SKIPPED — catch-up failed, mail may be missing`);
      skipped++;
      continue;
    }

    const deleteAfter = row.DeleteAfter ? new Date(row.DeleteAfter) : null;
    if (!deleteAfter) {
      console.log(`${capid} | ${row.Name} | SKIPPED — no DeleteAfter, the email needs a date`);
      skipped++;
      continue;
    }

    if (isDry) {
      console.log(`${capid} | ${row.Name} | would notify -> ${row.SeniorEmail}, ` +
        `${row.CadetEmail}, personal + commander | quoting delete date ` +
        formatTransitionDate_(deleteAfter));
      sent++;
      continue;
    }

    try {
      sendTransitionCompleteEmail_(row, Number(row.MessagesMigrated) || 0, deleteAfter);
      setTransitionField_(row._rowNumber, 'NotifiedDate', new Date().toISOString());
      sent++;
    } catch (e) {
      Logger.error('Notification failed', {
        capid: capid, errorMessage: e && e.message ? e.message : String(e)
      });
      failed++;
    }
  }

  console.log('');
  console.log(isDry
    ? `DRY RUN — ${sent} would be notified, ${skipped} skipped. Pass false to send.`
    : `Sent ${sent}, skipped ${skipped}, failed ${failed}.`);

  return { sent: sent, skipped: skipped, failed: failed };
}

/**
 * Formats a date for members to read, not for machines to parse.
 *
 * @param {Date} date
 * @returns {string} e.g. "29 July 2026"
 */
function formatTransitionDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'd MMMM yyyy');
}

// ============================================================================
// DESTINATION INSPECTION
// ============================================================================

/**
 * Message count already in the destination mailbox.
 *
 * Fails SOFT: returns null when the metadata scope is not granted, so a missing
 * grant degrades the duplicate guard to a warning rather than blocking every
 * migration. null means "unknown", which callers must not confuse with 0.
 *
 * @param {string} seniorEmail
 * @param {Object} cfg - from getCrossTenantConfig_()
 * @returns {?number} count, or null if it could not be determined
 */
function destinationMessageCount_(seniorEmail, cfg) {
  try {
    const token = xtPeerToken_(MIGRATE_INSPECT_SCOPE_, cfg, seniorEmail);
    const profile = gmailFetch_(
      'https://gmail.googleapis.com/gmail/v1/users/me/profile',
      { method: 'get', token: token },
      'destination profile for ' + seniorEmail
    );
    return Number(profile.messagesTotal) || 0;

  } catch (e) {
    Logger.warn('Could not read destination message count — duplicate guard inactive', {
      seniorEmail: seniorEmail,
      hint: 'add ' + MIGRATE_INSPECT_SCOPE_ + ' to the peer SA DWD grant',
      errorMessage: e && e.message ? e.message : String(e)
    });
    return null;
  }
}

/**
 * Decides whether the destination looks like it already holds this mail.
 *
 * Ratio rather than absolute count: a senior mailbox always has some mail, so
 * "not empty" proves nothing, whereas "already holds half as much as the source"
 * is hard to explain except by a migration that already happened.
 *
 * @param {?number} destCount - null when unknown
 * @param {number} sourceCount
 * @returns {boolean}
 */
function looksAlreadyMigrated_(destCount, sourceCount) {
  if (destCount === null) return false;              // unknown is not evidence
  if (destCount < MIGRATE_DUPLICATE_GUARD_MIN_) return false;
  if (!sourceCount) return false;
  return destCount >= sourceCount * MIGRATE_DUPLICATE_GUARD_RATIO_;
}

/**
 * Source mailbox message count.
 *
 * @param {string} cadetEmail
 * @returns {number}
 */
function sourceMessageCount_(cadetEmail) {
  const token = getImpersonatedToken_(cadetEmail, MIGRATE_SOURCE_SCOPE_);
  const profile = gmailFetch_(
    'https://gmail.googleapis.com/gmail/v1/users/me/profile',
    { method: 'get', token: token },
    'source profile for ' + cadetEmail
  );
  return Number(profile.messagesTotal) || 0;
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
function listSourceMessages_(user, token, pageToken, query) {
  const url = 'https://gmail.googleapis.com/gmail/v1/users/me/messages' +
    '?maxResults=' + MIGRATE_PAGE_SIZE_ +
    '&includeSpamTrash=false' +
    (query ? '&q=' + encodeURIComponent(query) : '') +
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
  // users/me, not the address: the token already impersonates this user, so 'me'
  // and the address are equivalent — but naming the address independently of the
  // token's subject is exactly the mismatch that caused the admin-token bug.
  // Keeping them coupled makes that class of error impossible.
  const url = 'https://gmail.googleapis.com/gmail/v1/users/me/messages/import' +
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
  // users/me — the token impersonates this user. See importToDestination_.
  const base = 'https://gmail.googleapis.com/gmail/v1/users/me/labels';

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
function scheduleMigrationContinuation_(scope) {
  const already = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === MIGRATE_CONTINUATION_FN_);

  if (already) {
    Logger.info('Migration continuation already scheduled');
    return;
  }

  // Written BEFORE the trigger exists: a trigger that fires without its scope
  // falls back to the conservative path, which is safe but would not resume what
  // was actually asked for.
  PropertiesService.getScriptProperties().setProperty(
    MIGRATE_CONTINUATION_SCOPE_PROP_, JSON.stringify(scope || {})
  );

  ScriptApp.newTrigger(MIGRATE_CONTINUATION_FN_)
    .timeBased()
    .after(60 * 1000)
    .create();

  Logger.info('Migration continuation scheduled', { scope: scope || {} });
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
      const srcToken = getImpersonatedToken_(row.CadetEmail, MIGRATE_SOURCE_SCOPE_);
      const profile = gmailFetch_(
        'https://gmail.googleapis.com/gmail/v1/users/me/profile',
        { method: 'get', token: srcToken },
        'profile for ' + row.CadetEmail
      );

      // Prove the DESTINATION credential by actually using it as the member.
      // Merely minting a token proves only that the scope is granted — which is
      // how the admin-subject bug survived the first version of this preview.
      //
      // labels.list, not users/me/profile: profile needs a READ scope
      // (gmail.readonly / gmail.metadata) and the destination token has neither
      // by design. Migration writes, it does not read, and the peer SA should
      // not hold read access to every senior mailbox just so a preview can print
      // a message count. gmail.labels covers labels.list, so this exercises a
      // real per-user call within the privilege migration actually needs.
      //
      // No address assertion needed here: the token exchange only succeeds if
      // the subject exists and the SA may impersonate it, and every destination
      // call uses users/me, so there is no subject/target mismatch of the kind
      // that caused the original bug.
      const destToken = xtPeerToken_(MIGRATE_DEST_SCOPES_, cfg, row.SeniorEmail);
      const destLabels = gmailFetch_(
        'https://gmail.googleapis.com/gmail/v1/users/me/labels',
        { method: 'get', token: destToken },
        'destination labels for ' + row.SeniorEmail
      );

      const hasLabel = (destLabels.labels || []).some(l => l.name === MIGRATE_LABEL_NAME_);
      const srcCount = Number(profile.messagesTotal) || 0;
      const destCount = destinationMessageCount_(row.SeniorEmail, cfg);

      let verdict;
      if (destCount === null) {
        verdict = 'dest OK (count unavailable — grant ' + MIGRATE_INSPECT_SCOPE_ +
          ' to enable the duplicate guard)';
      } else if (looksAlreadyMigrated_(destCount, srcCount)) {
        verdict = `WOULD BE BLOCKED — dest already holds ${destCount} vs source ${srcCount}. ` +
          'Looks already migrated; importing would duplicate.';
      } else {
        verdict = `dest OK (${destCount} already there)`;
      }
      if (hasLabel) verdict += ' [migration label already present]';

      console.log(`${capid} | ${row.Name} | ${row.NewType} | ` +
        `${row.CadetEmail} -> ${row.SeniorEmail} | ${srcCount} messages | ${verdict}`);

      if (!looksAlreadyMigrated_(destCount, srcCount)) total += srcCount;

    } catch (e) {
      console.log(`${capid} | ${row.Name} | ERROR: ${e && e.message ? e.message : String(e)}`);
    }
  }

  console.log('');
  console.log('Total messages to migrate: ' + total);
  console.log('Nothing was imported — this is a preview.');
}

/**
 * Brings across anything that landed in the cadet mailbox since it was migrated.
 *
 * The source account is NOT suspended and NOT closed — it keeps receiving mail
 * right up until deletion. The main migration is a point-in-time snapshot, so
 * without this, everything that arrived between migration and deletion would be
 * destroyed silently: the sender gets no bounce, the member never sees it, and
 * it is gone. This is the pass that makes the completion email's promise true.
 *
 * MUST run immediately before deletion. Safe to run repeatedly in between, and
 * worth doing so — each pass narrows the window that a final pass has to cover.
 *
 * Advances MigratedDate on success so the next run only looks at what is new.
 *
 * @returns {{caughtUp: number, imported: number}}
 */
function catchUpTransitionMail() {
  if (TRANSITION_CONFIG.ROLE !== 'source') {
    Logger.info('Catch-up skipped — not the source tenant');
    return { caughtUp: 0, imported: 0 };
  }

  const started = new Date();
  const rows = readTransitions_();
  let caughtUp = 0;
  let imported = 0;

  for (const capid in rows) {
    const row = rows[capid];
    if (row.MigrationStatus !== TRANSITION_CONFIG.STATUS.COMPLETE) continue;
    if (!row.MigratedDate || !row.SeniorEmail) continue;

    // Gmail's after: takes epoch seconds (the bare date form is day-granular,
    // which would re-import a whole day). Rewind a minute: same-second boundary
    // messages could otherwise slip through, and a couple of duplicates is a
    // better failure than silently losing mail.
    const since = Math.floor(new Date(row.MigratedDate).getTime() / 1000) - 60;
    const query = 'after:' + since;

    try {
      const before = Number(row.MessagesMigrated) || 0;

      // Reuse the full migration path: same resumability, same cursor handling,
      // same import semantics. Only the query differs.
      const result = migrateOneTransition_(row, started, { query: query, notify: false });
      const moved = result.imported - before;

      if (moved > 0) {
        Logger.info('Caught up new mail before deletion', {
          capid: capid, name: row.Name, newMessages: moved
        });
        imported += moved;
      }
      caughtUp++;

    } catch (e) {
      // Do NOT mark FAILED — the bulk migration succeeded and that status would
      // wrongly suggest it did not. But this MUST block deletion, since mail is
      // now known to be sitting in the source that is not in the destination.
      setTransitionField_(row._rowNumber, 'Notes',
        `Catch-up failed ${new Date().toISOString()} — DO NOT DELETE: ` +
        (e && e.message ? e.message : String(e)));
      Logger.error('Catch-up failed — deletion must not proceed', {
        capid: capid,
        errorMessage: e && e.message ? e.message : String(e)
      });
    }
  }

  Logger.info('Catch-up pass completed', {
    duration: new Date() - started + 'ms',
    rows: caughtUp,
    newMessages: imported
  });
  return { caughtUp: caughtUp, imported: imported };
}

/**
 * Records a mailbox as already migrated by other means, so the automation skips
 * the import but still runs the rest of the lifecycle over it.
 *
 * For mail moved by hand before this module existed. Sets the row up exactly as
 * a real migration would, so catch-up, deletion and the forwarding group all
 * behave normally — doing this by editing the sheet is easy to get subtly wrong,
 * because MigratedDate and DeleteAfter have to agree.
 *
 * @param {string} capid
 * @param {string} migratedDateIso - when the manual migration happened. THIS
 *   MATTERS: catch-up sweeps everything after it, so it is the boundary between
 *   "already copied" and "still to copy".
 *
 *   Err EARLY. Too early re-imports mail that is already there, costing
 *   duplicates the member can delete. Too late silently skips mail that arrived
 *   after the manual copy but before this row was marked, and that mail dies
 *   permanently when the account is deleted. If you cannot remember when you did
 *   it, guess generously in the early direction.
 *
 * @param {string} [note]
 * @returns {{capid: string, migratedDate: string, deleteAfter: string}}
 */
function markTransitionMigrated(capid, migratedDateIso, note) {
  if (!migratedDateIso) {
    throw new Error('migratedDateIso is required — it is the boundary catch-up ' +
      'sweeps from. Guess early rather than late: too early costs duplicates, ' +
      'too late loses mail permanently.');
  }

  const migratedAt = new Date(migratedDateIso);
  if (isNaN(migratedAt.getTime())) {
    throw new Error('Unparseable migratedDateIso: ' + migratedDateIso);
  }
  if (migratedAt > new Date()) {
    throw new Error('migratedDateIso is in the future: ' + migratedDateIso);
  }

  const rows = readTransitions_();
  const row = rows[String(capid)];
  if (!row) throw new Error('No transition row for CAPID ' + capid);

  // Deletion is measured from NOW, not from the manual migration: the 14 days
  // exist to give the member time to notice missing mail after being told, and
  // they are being told now.
  const deleteAfter = new Date();
  deleteAfter.setDate(deleteAfter.getDate() + TRANSITION_CONFIG.POST_MIGRATION_DELETE_DAYS);

  setTransitionField_(row._rowNumber, 'MigrationStatus', TRANSITION_CONFIG.STATUS.COMPLETE);
  setTransitionField_(row._rowNumber, 'MigratedDate', migratedAt.toISOString());
  setTransitionField_(row._rowNumber, 'DeleteAfter', deleteAfter.toISOString());
  setTransitionField_(row._rowNumber, 'LastCursor', '');
  setTransitionField_(row._rowNumber, 'Notes',
    `Migrated manually, recorded ${new Date().toISOString()}` + (note ? ': ' + note : ''));

  Logger.info('Transition marked as manually migrated', {
    capid: capid,
    name: row.Name,
    migratedDate: migratedAt.toISOString(),
    deleteAfter: deleteAfter.toISOString()
  });

  console.log(`Marked ${capid} (${row.Name}) COMPLETE.`);
  console.log(`  Catch-up will sweep everything after ${migratedAt.toISOString()}.`);
  console.log(`  No completion email was sent — send one, or tell them yourself,`);
  console.log(`  before the account is deleted after ${deleteAfter.toISOString()}.`);

  return {
    capid: String(capid),
    migratedDate: migratedAt.toISOString(),
    deleteAfter: deleteAfter.toISOString()
  };
}

/**
 * Clears FAILED rows back to PENDING so migration will retry them.
 *
 * FAILED is terminal by design — it blocks both retry and deletion until a human
 * has looked — but that left no way back once the cause was fixed. Use after
 * addressing whatever the Notes column reports.
 *
 * Also clears LastCursor: a row that failed mid-mailbox has an untrustworthy
 * cursor, and re-importing from the start risks duplicates where resuming from a
 * bad cursor risks silently skipping mail. Duplicates are recoverable.
 *
 * @returns {{reset: number}}
 */
function resetFailedTransitions() {
  const rows = readTransitions_();
  let reset = 0;

  for (const capid in rows) {
    const row = rows[capid];
    if (row.MigrationStatus !== TRANSITION_CONFIG.STATUS.FAILED) continue;

    setTransitionField_(row._rowNumber, 'MigrationStatus', TRANSITION_CONFIG.STATUS.PENDING);
    setTransitionField_(row._rowNumber, 'LastCursor', '');
    setTransitionField_(row._rowNumber, 'Notes',
      `Reset to PENDING ${new Date().toISOString()} (was: ${row.Notes || 'no note'})`);

    Logger.info('Transition row reset for retry', { capid: capid, name: row.Name });
    reset++;
  }

  Logger.info('Failed transitions reset', { reset: reset });
  return { reset: reset };
}
