/**
 * LSCode (FBI Background Check) Change Notification
 *
 * Version: 1.1.0
 * Date: 2026-07-15
 * Changes: Report the window each change was detected in, and state it in the
 *   digest. State file gains a per-member `seen` date (v2). Intended cadence is
 *   weekly. (1.0.0: new module. Notifies squadron commanders when a member under
 *   their command gains or loses their FBI background check.)
 *   See PCR_CHANGELOG.md.
 *
 * Authors: Isaac Wilson IV
 *
 * WHAT LSCode IS
 * Member.txt column `LSCode` is the FBI background check flag: 'A' means the
 * member has passed, blank means they have not. Seniors and FIFTY YEAR members
 * carry 'A'; cadets and PATRON members are blank. Nothing else in this codebase
 * reads the column.
 *
 * It is NOT a per-person "has a background check on file" flag. Cadets over 18
 * do undergo a check and still show blank (checked against CAPID 612148), so the
 * column reflects the senior-side record. This module therefore only means
 * anything on a tenant carrying senior members, which is what
 * PROFILE_.RUN_LSCODE_NOTIFICATIONS gates.
 *
 * WHY THIS IS A STANDALONE MODULE
 * updateAllMembers() already keeps a Drive snapshot (CurrentMembers.txt) and
 * diffs against it, so hooking in there looks tempting. Two reasons not to:
 *
 *   1. It is the account-provisioning job. If it sent commander mail and then
 *      threw before saveCurrentMemberData(), the next run would re-detect the
 *      same transitions and mail every commander a second time.
 *   2. forceUpdateAllMembers() writes that snapshot without diffing, so it would
 *      silently swallow pending transitions and those commanders would never be
 *      told.
 *
 * This module therefore keeps its own state file, runs on its own trigger, and
 * never touches provisioning. A failure here cannot affect account creation, and
 * a force-provision cannot suppress a notification.
 *
 * FIRST RUN IS SILENT BY CONSTRUCTION
 * A member absent from the state file is recorded without notifying. That covers
 * both the initial backfill (state file does not exist yet -> every member is
 * new -> zero mail) and genuinely new members, whose existing clearance is not
 * news to anyone. Only a member whose *recorded* value differs from their
 * current one produces a notification.
 *
 * WHY THE DIGEST DATES A WINDOW, NOT A DAY
 * CAPWATCH publishes no date for an LSCode change. `Member.txt` carries `DateMod`
 * (when the *record* was last modified) and `UsrID` (who modified it), but both
 * are record-level: an address edit moves `DateMod` just as a background check
 * does, so it cannot be quoted as the date a check cleared. There is no
 * background-check table anywhere in the extract either.
 *
 * So the only defensible date is our own: the change happened somewhere between
 * the last time we confirmed the old value and the run that saw the new one. The
 * digest says exactly that and no more. On a weekly cadence the window is a week,
 * which is the resolution this data actually supports.
 *
 * That window is per-member (`seen`), not a single global last-run date. A digest
 * that fails to send leaves its members pending with their original `seen`
 * intact, so when the retry lands a week later it still reports the week the
 * change was really detected, rather than the week of the retry.
 *
 * Setup:
 * 1. Set TENANT_PROFILE appropriately; this runs where RUN_LSCODE_NOTIFICATIONS
 *    is true (see config.gs).
 * 2. Run previewLSCodeChanges() first — it sends nothing and writes nothing.
 * 3. Run notifyLSCodeChanges() once by hand to lay down the baseline (silent).
 * 4. Add a time-driven trigger for notifyLSCodeChanges(), WEEKLY, on a day and
 *    hour that falls after getCapwatch() has refreshed the extract. The cadence
 *    sets the width of the reported window, so changing it changes what the
 *    digest claims — run it weekly and the digest says "detected 8–15 Jul".
 */

const LSCODE_NOTIFY_CONFIG = {
  // Own state file, deliberately not CurrentMembers.txt — see module header.
  STATE_FILE_NAME: 'LSCodeState.txt',

  // Bumped when the state file's shape changes. An unrecognised version is
  // re-baselined silently rather than guessed at: a misread state file would
  // either mail the whole wing or swallow real changes.
  STATE_VERSION: 2,

  // The one value that means "passed the FBI background check".
  CLEARED: 'A',

  SUBJECT_GRANTED: 'Background check cleared in your unit',
  SUBJECT_REVOKED: 'Background check no longer current in your unit',
  SUBJECT_MIXED: 'Background check status changes in your unit',

  // Spacing between commander digests, matching RETENTION_CONFIG.EMAIL_DELAY_MS.
  EMAIL_DELAY_MS: 1000
};

// ============================================================================
// MAIN
// ============================================================================

/**
 * Detects LSCode transitions since the last run and mails each affected
 * squadron commander a digest of the members under their command who changed.
 *
 * @returns {Object} Summary of the run
 */
function notifyLSCodeChanges() {
  return runLSCodeNotification_({ dryRun: false });
}

/**
 * Dry run: reports what notifyLSCodeChanges() would send, without sending any
 * mail or advancing the state file. Safe to run at any time.
 *
 * @returns {Object} Summary of the run
 */
function previewLSCodeChanges() {
  return runLSCodeNotification_({ dryRun: true });
}

/**
 * @param {Object} options - { dryRun: boolean }
 * @returns {Object} Summary of the run
 */
function runLSCodeNotification_(options) {
  const dryRun = !!(options && options.dryRun);

  if (!PROFILE_.RUN_LSCODE_NOTIFICATIONS) {
    Logger.info('LSCode notifications disabled for this tenant profile', {
      profile: TENANT_PROFILE
    });
    return { skipped: true };
  }

  clearCache(); // Ensure fresh CAPWATCH data
  const start = new Date();
  Logger.info(dryRun ? 'Starting LSCode change PREVIEW' : 'Starting LSCode change notification');

  const current = readMemberLSCodes_();
  if (!Object.keys(current).length) {
    // An empty Member.txt would otherwise read as "everyone departed" and wipe
    // the state file, making the next real run announce the whole wing.
    Logger.error('No members read from Member.txt — aborting without touching state');
    return { aborted: true, reason: 'no members' };
  }

  const prior = loadLSCodeState_();
  const isBaseline = !Object.keys(prior).length;
  const todayIso = isoDate_(start);

  const diff = diffLSCodes_(current, prior, todayIso);

  Logger.info('LSCode diff complete', {
    members: Object.keys(current).length,
    firstSeen: diff.firstSeen,
    unchanged: diff.unchanged,
    granted: diff.granted,
    revoked: diff.revoked,
    orgsAffected: Object.keys(diff.byOrg).length,
    baseline: isBaseline
  });

  if (isBaseline) {
    Logger.info('No prior state — laying down a silent baseline, no mail will be sent');
  }

  // Departed members fall out of state naturally: state is rebuilt from the
  // current roster. If they return they read as first-seen (silent), which is
  // right — a returning member's existing clearance is not news.
  //
  // Everyone here is recorded as confirmed today. Members whose digest fails to
  // send are rolled back below, which is what preserves their original window.
  const nextState = {};
  for (const capid in current) {
    nextState[capid] = { c: current[capid].lscode, seen: todayIso };
  }

  const commanders = buildCommanderMap_();
  const summary = {
    dryRun: dryRun,
    baseline: isBaseline,
    members: Object.keys(current).length,
    granted: diff.granted,
    revoked: diff.revoked,
    firstSeen: diff.firstSeen,
    // Carried so the returned summary accounts for every member read, not just
    // the interesting ones: members = firstSeen + unchanged + granted + revoked
    // (+ any non-boundary changes, which warn). A summary that cannot be
    // reconciled against the roster hides whichever bucket is wrong.
    unchanged: diff.unchanged,
    sent: 0,
    failedOrgs: [],
    noCommanderOrgs: [],
    startTime: start.toISOString()
  };

  for (const orgid in diff.byOrg) {
    const changes = diff.byOrg[orgid];
    const commander = commanders[String(orgid)];

    if (!commander || !commander.email) {
      // Do NOT advance state for these members. Leaving them pending means the
      // notification fires once a commander is assigned, rather than being lost
      // to a vacancy. It re-reports every run until then, which is the point:
      // an undeliverable compliance signal should stay visible.
      Logger.warn('No commander email for org with LSCode changes — leaving pending', {
        orgid: orgid,
        orgName: changes[0].orgName,
        memberCount: changes.length
      });
      summary.noCommanderOrgs.push({
        orgid: orgid,
        orgName: changes[0].orgName,
        members: changes.length
      });
      revertOrgState_(nextState, changes, prior);
      continue;
    }

    if (dryRun) {
      Logger.info('[PREVIEW] Would send LSCode digest', {
        orgid: orgid,
        orgName: changes[0].orgName,
        to: commander.email,
        members: changes.map(c =>
          c.capid + ' ' + c.direction + ' (' + formatWindow_(c.windowFrom, c.windowTo) + ')'
        )
      });
      summary.sent++;
      continue;
    }

    const ok = sendCommanderDigest_(commander, changes);
    if (ok) {
      summary.sent++;
    } else {
      // Per-org commit: a failed digest leaves its members at their prior value
      // so the next run retries them. Successful orgs still advance, so one bad
      // send cannot re-mail everybody else.
      summary.failedOrgs.push({ orgid: orgid, orgName: changes[0].orgName });
      revertOrgState_(nextState, changes, prior);
    }

    Utilities.sleep(LSCODE_NOTIFY_CONFIG.EMAIL_DELAY_MS);
  }

  if (!dryRun) {
    saveLSCodeState_(nextState);
  } else {
    Logger.info('[PREVIEW] State file not written');
  }

  summary.endTime = new Date().toISOString();
  summary.duration = new Date() - start;
  Logger.info(dryRun ? 'LSCode preview complete' : 'LSCode notification complete', summary);

  if (!dryRun && !isBaseline && (summary.failedOrgs.length || summary.noCommanderOrgs.length)) {
    sendLSCodeSummaryEmail_(summary);
  }

  return summary;
}

/**
 * Restores the prior recorded value for an org's changed members, so a run that
 * could not deliver their digest retries them next time.
 *
 * Restoring the whole prior record — including its `seen` date, not just the
 * value — is what makes the retry report the right week. Rewriting `seen` to
 * today would leave the change pending but make next week's digest claim it was
 * detected next week.
 *
 * @param {Object} nextState - State being built for this run
 * @param {Array<Object>} changes - Changed members for one org
 * @param {Object} prior - State as loaded at the start of this run
 * @returns {void}
 */
function revertOrgState_(nextState, changes, prior) {
  changes.forEach(function (change) {
    const record = prior[change.capid];
    if (record === undefined) {
      delete nextState[change.capid];
    } else {
      nextState[change.capid] = record;
    }
  });
}

// ============================================================================
// CAPWATCH READS
// ============================================================================

/**
 * Reads every member's LSCode from Member.txt, resolved by header name.
 *
 * parseFile() strips the header and hands back positional rows, which is how
 * this codebase has historically read CAPWATCH. Column 16 (Expiration) went
 * unverified for months that way. Resolve by name instead — the cost is one
 * extra read of a file that is already in Drive.
 *
 * @returns {Object} Map of CAPID to { capid, firstName, lastName, rank, type, orgid, orgName, lscode }
 */
function readMemberLSCodes_() {
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const it = folder.getFilesByName('Member.txt');
  if (!it.hasNext()) {
    Logger.error('Member.txt not found', { folderId: CONFIG.CAPWATCH_DATA_FOLDER_ID });
    return {};
  }

  const rows = Utilities.parseCsv(it.next().getBlob().getDataAsString());
  if (!rows || rows.length < 2) {
    Logger.error('Member.txt empty or header-only');
    return {};
  }

  const header = rows[0].map(h => String(h || '').trim());
  const idx = {
    capid: header.indexOf('CAPID'),
    lastName: header.indexOf('NameLast'),
    firstName: header.indexOf('NameFirst'),
    orgid: header.indexOf('ORGID'),
    rank: header.indexOf('Rank'),
    type: header.indexOf('Type'),
    status: header.indexOf('MbrStatus'),
    lscode: header.indexOf('LSCode')
  };

  // Fail loudly rather than silently reading the wrong column if CAP ever
  // reshapes the extract.
  const missing = Object.keys(idx).filter(k => idx[k] === -1);
  if (missing.length) {
    Logger.error('Member.txt is missing expected columns', {
      missing: missing,
      header: header.join(',')
    });
    return {};
  }

  const squadrons = getSquadrons();
  const excluded = PROFILE_.EXCLUDED_ORG_IDS || [];
  const out = {};
  let outOfScope = 0;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row[idx.capid]) continue;

    // Only members this tenant actually manages. A change on someone the tenant
    // does not carry is not the local commander's business.
    const status = String(row[idx.status] || '').trim();
    if (status !== 'ACTIVE') continue;

    const orgid = String(row[idx.orgid] || '').trim();
    if (excluded.indexOf(orgid) > -1) continue;

    // getSquadrons() is filtered to CONFIG.WING, so this doubles as the
    // in-tenant scope check: an org it does not know is not this tenant's to
    // report on, and its commander would not be ours to mail.
    const squadron = squadrons[orgid];
    if (!squadron) {
      outOfScope++;
      continue;
    }

    const capid = String(row[idx.capid]).trim();

    out[capid] = {
      capid: capid,
      firstName: String(row[idx.firstName] || '').trim(),
      lastName: String(row[idx.lastName] || '').trim(),
      rank: String(row[idx.rank] || '').trim(),
      type: String(row[idx.type] || '').trim(),
      orgid: orgid,
      orgName: squadron.name,
      charter: squadron.charter,
      lscode: String(row[idx.lscode] || '').trim()
    };
  }

  Logger.info('Member LSCodes read', {
    inScope: Object.keys(out).length,
    outOfScope: outOfScope
  });

  return out;
}

/**
 * Builds an ORGID -> commander map in one pass.
 *
 * getCommanderInfo() in SendRetentionEmail.gs answers the same question, but
 * rebuilds the whole email map on every call. Digesting per org would make that
 * O(orgs x contacts); this is one pass. Note Commanders.txt is nationwide, not
 * wing-scoped, so it is only ever read through orgids drawn from this tenant's
 * own Member.txt.
 *
 * @returns {Object} Map of ORGID to { capid, firstName, lastName, rank, email }
 */
function buildCommanderMap_() {
  const commanderData = parseFile('Commanders');
  const emailMap = createEmailMap();
  const map = {};

  for (let i = 0; i < commanderData.length; i++) {
    const row = commanderData[i];
    const orgid = String(row[0] || '').trim();
    const capid = String(row[4] || '').trim();
    if (!orgid || !capid) continue;

    map[orgid] = {
      capid: capid,
      lastName: row[8],
      firstName: row[9],
      rank: row[12],
      email: emailMap[capid] || null
    };
  }

  Logger.info('Commander map built', { orgs: Object.keys(map).length });
  return map;
}

// ============================================================================
// DIFF
// ============================================================================

/**
 * Compares current LSCodes against the recorded state.
 *
 * A member with no recorded value is first-seen and never notified — this is
 * what keeps the first run, and every new member afterwards, silent.
 *
 * Each change carries the window it was detected in: from the date we last
 * confirmed the old value (the member's recorded `seen`) to today. See the module
 * header for why a window rather than a date.
 *
 * @param {Object} current - Map of CAPID to member info incl. lscode
 * @param {Object} prior - Map of CAPID to { c, seen }
 * @param {string} todayIso - This run's date, 'yyyy-MM-dd'
 * @returns {Object} { byOrg, granted, revoked, firstSeen, unchanged }
 */
function diffLSCodes_(current, prior, todayIso) {
  const cleared = LSCODE_NOTIFY_CONFIG.CLEARED;
  const result = { byOrg: {}, granted: 0, revoked: 0, firstSeen: 0, unchanged: 0 };

  for (const capid in current) {
    const member = current[capid];
    const record = prior[capid];

    if (record === undefined) {
      result.firstSeen++;
      continue;
    }

    const previous = record.c;
    if (previous === member.lscode) {
      result.unchanged++;
      continue;
    }

    // '' -> 'A' is a grant; 'A' -> '' is a revocation. The column has no other
    // observed values, but anything non-'A' is treated as "not cleared" rather
    // than assumed away.
    const nowCleared = member.lscode === cleared;
    const wasCleared = previous === cleared;
    if (nowCleared === wasCleared) {
      // Changed, but not across the cleared boundary (e.g. '' -> some new code).
      // Not a background-check event; record it and move on rather than mailing
      // a commander about something this module does not understand.
      Logger.warn('LSCode changed without crossing the cleared boundary', {
        capsn: capid,
        from: previous,
        to: member.lscode
      });
      continue;
    }

    const direction = nowCleared ? 'GRANTED' : 'REVOKED';
    if (nowCleared) result.granted++; else result.revoked++;

    if (!result.byOrg[member.orgid]) result.byOrg[member.orgid] = [];
    result.byOrg[member.orgid].push({
      capid: capid,
      name: [member.rank, member.firstName, member.lastName].filter(String).join(' '),
      type: member.type,
      orgName: member.orgName,
      charter: member.charter,
      from: previous,
      to: member.lscode,
      direction: direction,
      // The change happened somewhere in here. `seen` is the last date we
      // confirmed the old value, so on a weekly cadence this spans a week.
      windowFrom: record.seen,
      windowTo: todayIso
    });
  }

  return result;
}

// ============================================================================
// STATE
// ============================================================================

/**
 * Loads the recorded state: CAPID -> { c: lscode, seen: 'yyyy-MM-dd' }, where
 * `seen` is the date we last confirmed that member held that value.
 *
 * @returns {Object} Recorded members map, or {} if this tenant has no usable state
 */
function loadLSCodeState_() {
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files = folder.getFilesByName(LSCODE_NOTIFY_CONFIG.STATE_FILE_NAME);

  if (!files.hasNext()) {
    Logger.info('No LSCode state file yet — this run will be a silent baseline');
    return {};
  }

  const content = files.next().getBlob().getDataAsString();
  if (!content) return {};

  let parsed;
  try {
    parsed = JSON.parse(content);
  } catch (e) {
    // Refusing to guess: returning {} here would look like a baseline and
    // silently swallow every pending change. Fail instead.
    Logger.error('LSCode state file is corrupt — refusing to run', {
      errorMessage: e.message,
      fileName: LSCODE_NOTIFY_CONFIG.STATE_FILE_NAME
    });
    throw new Error(
      'Cannot parse ' + LSCODE_NOTIFY_CONFIG.STATE_FILE_NAME + '. Fix or delete it ' +
      '(deleting re-baselines silently) before running again.'
    );
  }

  if (!parsed || parsed.version !== LSCODE_NOTIFY_CONFIG.STATE_VERSION) {
    // A v1 file (a flat CAPID -> lscode map) carries no `seen` date, so no window
    // could be reported for anything already pending in it. Re-baseline instead:
    // silent, and at worst it costs one cycle of notifications.
    Logger.warn('LSCode state version not recognised — re-baselining silently', {
      found: parsed ? parsed.version : null,
      expected: LSCODE_NOTIFY_CONFIG.STATE_VERSION
    });
    return {};
  }

  return parsed.members || {};
}

/**
 * Writes the members map, creating the file when absent.
 *
 * saveCurrentMemberData() warns and does nothing when its file is missing,
 * which would leave this module permanently baselining. Create it instead.
 *
 * @param {Object} members - CAPID to { c, seen }
 * @returns {void}
 */
function saveLSCodeState_(members) {
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files = folder.getFilesByName(LSCODE_NOTIFY_CONFIG.STATE_FILE_NAME);
  const content = JSON.stringify({
    version: LSCODE_NOTIFY_CONFIG.STATE_VERSION,
    written: isoDate_(new Date()),
    members: members
  });

  if (files.hasNext()) {
    files.next().setContent(content);
  } else {
    folder.createFile(LSCODE_NOTIFY_CONFIG.STATE_FILE_NAME, content);
  }

  Logger.info('LSCode state saved', {
    memberCount: Object.keys(members).length,
    fileName: LSCODE_NOTIFY_CONFIG.STATE_FILE_NAME
  });
}

/**
 * Formats a Date as 'yyyy-MM-dd' in the script's timezone. Dates are stored and
 * compared as calendar days: the window is only ever reported to day precision,
 * so a timestamp would imply accuracy this data does not have.
 *
 * @param {Date} date - Date to format
 * @returns {string} 'yyyy-MM-dd'
 */
function isoDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Renders a detection window for a commander, e.g. '8–15 Jul 2026',
 * '28 Jun – 5 Jul 2026', or '28 Dec 2025 – 4 Jan 2026'. Collapses the repeated
 * month and year when both ends share them.
 *
 * @param {string} fromIso - Window start, 'yyyy-MM-dd'
 * @param {string} toIso - Window end, 'yyyy-MM-dd'
 * @returns {string} Human-readable range
 */
function formatWindow_(fromIso, toIso) {
  const from = parseIsoDate_(fromIso);
  const to = parseIsoDate_(toIso);
  if (!from || !to) return '';

  const tz = Session.getScriptTimeZone();
  const fmt = (d, pattern) => Utilities.formatDate(d, tz, pattern);

  if (fromIso === toIso) return fmt(to, 'd MMM yyyy');

  const sameYear = fmt(from, 'yyyy') === fmt(to, 'yyyy');
  const sameMonth = sameYear && fmt(from, 'MM') === fmt(to, 'MM');

  if (sameMonth) return fmt(from, 'd') + '–' + fmt(to, 'd MMM yyyy');
  if (sameYear) return fmt(from, 'd MMM') + ' – ' + fmt(to, 'd MMM yyyy');
  return fmt(from, 'd MMM yyyy') + ' – ' + fmt(to, 'd MMM yyyy');
}

/**
 * @param {string} iso - 'yyyy-MM-dd'
 * @returns {Date|null} Parsed date, or null if unparseable
 */
function parseIsoDate_(iso) {
  const parts = String(iso || '').split('-');
  if (parts.length !== 3) return null;
  const year = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10);
  const day = parseInt(parts[2], 10);
  if (isNaN(year) || isNaN(month) || isNaN(day)) return null;
  return new Date(year, month - 1, day);
}

/**
 * Discards the recorded state so the next run re-baselines silently.
 * Use if the state file is corrupt or the roster has been rebuilt.
 *
 * @returns {void}
 */
function resetLSCodeState_() {
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files = folder.getFilesByName(LSCODE_NOTIFY_CONFIG.STATE_FILE_NAME);
  let removed = 0;
  while (files.hasNext()) {
    files.next().setTrashed(true);
    removed++;
  }
  Logger.info('LSCode state reset — next run re-baselines silently', { filesTrashed: removed });
}

// ============================================================================
// EMAIL
// ============================================================================

/**
 * Sends one commander a digest of every changed member under their command.
 *
 * @param {Object} commander - Commander object with email
 * @param {Array<Object>} changes - Changed members for this org
 * @returns {boolean} True if sent
 */
function sendCommanderDigest_(commander, changes) {
  const granted = changes.filter(c => c.direction === 'GRANTED');
  const revoked = changes.filter(c => c.direction === 'REVOKED');

  let subject;
  if (granted.length && revoked.length) subject = LSCODE_NOTIFY_CONFIG.SUBJECT_MIXED;
  else if (revoked.length) subject = LSCODE_NOTIFY_CONFIG.SUBJECT_REVOKED;
  else subject = LSCODE_NOTIFY_CONFIG.SUBJECT_GRANTED;

  // Charter (PCR-CA-070) rather than name: it is how commanders identify units.
  subject += ' — ' + changes[0].charter;

  try {
    const htmlBody = buildDigestHtml_(commander, granted, revoked);

    executeWithRetry(() =>
      GmailApp.sendEmail(
        commander.email,
        subject,
        'See the HTML version of this message.',
        {
          htmlBody: htmlBody,
          from: AUTOMATION_SENDER_EMAIL,
          name: SENDER_NAME,
          replyTo: ITSUPPORT_EMAIL
        }
      )
    );

    Logger.info('LSCode digest sent', {
      to: commander.email,
      orgName: changes[0].orgName,
      granted: granted.length,
      revoked: revoked.length
    });
    return true;

  } catch (e) {
    Logger.error('Failed to send LSCode digest', {
      to: commander.email,
      orgName: changes[0].orgName,
      errorMessage: e.message,
      errorCode: e.details?.code
    });
    return false;
  }
}

/**
 * @param {Object} commander - Commander object
 * @param {Array<Object>} granted - Members who gained their check
 * @param {Array<Object>} revoked - Members who lost it
 * @returns {string} HTML body
 */
function buildDigestHtml_(commander, granted, revoked) {
  const rows = list => list.map(m =>
    '<tr><td>' + escapeHtml_(m.name) + '</td><td>' + escapeHtml_(m.capid) +
    '</td><td>' + escapeHtml_(m.type) + '</td><td>' +
    escapeHtml_(formatWindow_(m.windowFrom, m.windowTo)) + '</td></tr>'
  ).join('');

  const header = '<tr><th>Member</th><th>CAPID</th><th>Type</th><th>Detected</th></tr>';

  let html =
    '<html><head><style>' +
    'body { font-family: Arial, sans-serif; color: #202124; }' +
    'h2 { color: #1a73e8; }' +
    'table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }' +
    'th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }' +
    'th { background-color: #1a73e8; color: white; }' +
    '.revoked { background-color: #fff3cd; padding: 10px; border-left: 4px solid #ffc107; }' +
    '.footer { font-size: 12px; color: #666; }' +
    '</style></head><body>' +
    '<p>' + escapeHtml_([commander.rank, commander.lastName].filter(String).join(' ')) + ',</p>' +
    '<p>The following members under your command have had a change to their FBI ' +
    'background check status in eServices since the last check.</p>';

  if (granted.length) {
    html +=
      '<h2>Background check now cleared (' + granted.length + ')</h2>' +
      '<table>' + header + rows(granted) + '</table>';
  }

  if (revoked.length) {
    html +=
      '<div class="revoked"><h2>Background check no longer current (' + revoked.length + ')</h2>' +
      '<p>These members previously showed a completed FBI background check and no ' +
      'longer do. This may reflect a records change rather than an action against ' +
      'the member — please verify in eServices before acting.</p></div>' +
      '<table>' + header + rows(revoked) + '</table>';
  }

  // Say plainly what "Detected" means. CAPWATCH publishes no date for an LSCode
  // change, so this is the window between our checks, not a date from eServices —
  // a commander reading a date in an official-looking email will otherwise assume
  // it came from the record.
  html +=
    '<hr><p class="footer">Automated notification derived from the CAPWATCH ' +
    'extract (Member.txt, LSCode). <strong>Detected</strong> is the period between ' +
    'our checks in which the change appeared — CAPWATCH does not publish the date a ' +
    'background check changed, so the exact date is not available here. eServices is ' +
    'authoritative. Questions: ' + escapeHtml_(ITSUPPORT_EMAIL) + '.</p></body></html>';

  return html;
}

/**
 * Reports undeliverable digests to IT support. Only sent when something needs a
 * human — a clean run stays quiet.
 *
 * @param {Object} summary - Run summary
 * @returns {void}
 */
function sendLSCodeSummaryEmail_(summary) {
  try {
    let html =
      '<html><body style="font-family: Arial, sans-serif;">' +
      '<h2>LSCode notification — items needing attention</h2>' +
      '<p>Sent ' + summary.sent + ' commander digest(s). ' +
      'Granted: ' + summary.granted + ', revoked: ' + summary.revoked + '.</p>';

    if (summary.noCommanderOrgs.length) {
      html +=
        '<h3>No commander on record (' + summary.noCommanderOrgs.length + ')</h3>' +
        '<p>These units have LSCode changes but no commander email in CAPWATCH. ' +
        'They remain pending and will notify once a commander is assigned.</p><ul>' +
        summary.noCommanderOrgs.map(o =>
          '<li>' + escapeHtml_(o.orgName) + ' (ORGID ' + escapeHtml_(o.orgid) +
          ') — ' + o.members + ' member(s)</li>'
        ).join('') + '</ul>';
    }

    if (summary.failedOrgs.length) {
      html +=
        '<h3>Send failed (' + summary.failedOrgs.length + ')</h3>' +
        '<p>These will be retried on the next run.</p><ul>' +
        summary.failedOrgs.map(o =>
          '<li>' + escapeHtml_(o.orgName) + ' (ORGID ' + escapeHtml_(o.orgid) + ')</li>'
        ).join('') + '</ul>';
    }

    html += '</body></html>';

    GmailApp.sendEmail(ITSUPPORT_EMAIL, 'LSCode notification — attention needed', 'See HTML version', {
      htmlBody: html,
      from: AUTOMATION_SENDER_EMAIL,
      name: SENDER_NAME
    });
  } catch (e) {
    Logger.error('Failed to send LSCode summary email', { errorMessage: e.message });
  }
}

/**
 * Escapes CAPWATCH-sourced text for inclusion in an HTML email body.
 *
 * @param {string} value - Raw value
 * @returns {string} Escaped value
 */
function escapeHtml_(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ============================================================================
// TEST
// ============================================================================

/**
 * Renders a digest from fabricated data and sends it to TEST_EMAIL, so the
 * template can be reviewed without waiting for a real transition.
 *
 * @returns {void}
 */
function testLSCodeDigestToTestEmail() {
  // A week-wide window, as a weekly trigger would produce.
  const to = isoDate_(new Date());
  const from = isoDate_(new Date(Date.now() - 7 * 24 * 60 * 60 * 1000));

  const commander = { rank: 'Maj', firstName: 'Pat', lastName: 'Example', email: TEST_EMAIL };
  const granted = [
    { name: 'Capt Jamie Rivera', capid: '123456', type: 'SENIOR', direction: 'GRANTED', windowFrom: from, windowTo: to },
    { name: '2d Lt Alex Chen', capid: '234567', type: 'SENIOR', direction: 'GRANTED', windowFrom: from, windowTo: to }
  ];
  const revoked = [
    { name: '1st Lt Sam Fitzgerald', capid: '345678', type: 'SENIOR', direction: 'REVOKED', windowFrom: from, windowTo: to }
  ];

  const htmlBody = buildDigestHtml_(commander, granted, revoked);
  GmailApp.sendEmail(TEST_EMAIL, 'TEST — LSCode digest preview', 'See HTML version', {
    htmlBody: htmlBody,
    from: AUTOMATION_SENDER_EMAIL,
    name: SENDER_NAME
  });

  Logger.info('Test LSCode digest sent', { recipient: TEST_EMAIL });
}
