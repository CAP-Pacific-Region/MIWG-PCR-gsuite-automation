/**
 * -------------------------------------------------------------------------
 * Version: 1.0.0
 * Date: 2026-07-26
 * Authors: Michigan Wing (MIWG) — Extended and Maintained by Lt Col Noel Luneau
 * Contributors: Maj Isaac Wilson IV, California Wing (1.0.0)
 * Changes: 1.0.0 — new module. Records every welcome email this system sends,
 *   so an account that never received one can be DETECTED rather than stumbled
 *   over. Pairs with WelcomeEmailResend.gs, which repairs what this finds.
 * -------------------------------------------------------------------------
 *
 * THE PROBLEM THIS CLOSES
 *
 * sendWelcomeEmail() has one call site: the insert branch of addOrUpdateUser().
 * An account created out-of-band (Admin console, GAM) never reaches it, so its
 * member never gets credentials — and every later sync, finding an account
 * already there, simply updates it. The gap is invisible. A welcome email that
 * throws on send (no deliverable recipient) is equally invisible: the failure is
 * caught and logged, and nothing revisits it.
 *
 * WHY A LEDGER, AND WHY IT MUST BE SEEDED
 *
 * "Was this member ever welcomed?" is not answerable from the directory: nothing
 * about an account records whether mail was sent to it. So this module writes it
 * down. sendWelcomeEmail() records each send here, and the scan reports members
 * holding an account with no entry.
 *
 * On the day the ledger goes live, NOBODY has an entry — the whole wing looks
 * unwelcomed. seedWelcomeLedger() establishes the baseline, and does it on the
 * one honest piece of evidence available in hindsight: an account that HAS BEEN
 * SIGNED INTO plainly received working credentials at some point, whatever the
 * route. Those are recorded as welcomed. Accounts that have never been signed
 * into cannot be judged either way, so the seed deliberately does NOT bury them
 * — they are reported as UNKNOWN, and that list is where an already-missed
 * member (the case that prompted all this) will be sitting.
 *
 * WHAT EACH VERDICT IS WORTH
 *   MISSED   — the account was created AFTER the baseline and no send was ever
 *              recorded. Near-certain: created out-of-band, or the send failed.
 *   UNKNOWN  — predates the baseline and has never been signed into. A genuine
 *              maybe: could be a member who never got credentials, could be one
 *              who got them and never logged in.
 *   WELCOMED — a recorded send, or seeded from login history.
 *
 * Only MISSED is mailed to IT. UNKNOWN is a review list, not a to-do list —
 * paging IT monthly about members who are merely quiet is how a notification
 * gets ignored. Note that the never-signed-in population is ALREADY surfaced to
 * unit command staff by notifications/RecoveryEmailNotify.gs (its LOGIN
 * condition); this module is not a second copy of that, and does not mail units.
 */

const WELCOME_AUDIT_CONFIG = {
  // Own state file, in the CAPWATCH data folder alongside the other state.
  // Nothing else writes it.
  LEDGER_FILE_NAME: 'WelcomeEmailLedger.txt',

  // Bumped when the ledger's shape changes. An unrecognised version is REFUSED
  // rather than re-baselined — see welcomeLedgerLoad_.
  LEDGER_VERSION: 1,

  // An account younger than this is never reported MISSED. Provisioning creates
  // the account and sends the welcome email as separate steps, and a resend may
  // be in flight; a two-day floor keeps the report free of work already in hand.
  NEW_ACCOUNT_GRACE_DAYS: 2,

  SUBJECT: 'Workspace accounts that never received a welcome email'
};

// ============================================================================
// LEDGER
// ============================================================================

/**
 * Loads the ledger.
 *
 * A missing file is normal exactly once (before the first seed) and is reported
 * as an empty, UNSEEDED ledger — which makes every member UNKNOWN rather than
 * MISSED, so a lost or not-yet-created file can never manufacture a wing-wide
 * list of false accusations. A file that exists but cannot be read is a
 * different matter and throws: silently continuing would either bury real
 * findings or invent them, depending on which way the guess fell.
 *
 * @returns {{seededAt: string, sent: Object}} `seededAt` is '' when unseeded;
 *   `sent` maps CAPID -> { on: iso-date, by: 'send'|'seed' }
 */
function welcomeLedgerLoad_() {
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files = folder.getFilesByName(WELCOME_AUDIT_CONFIG.LEDGER_FILE_NAME);

  if (!files.hasNext()) {
    Logger.info('No welcome-email ledger yet — run seedWelcomeLedger() to establish the baseline');
    return { seededAt: '', sent: {} };
  }

  const content = files.next().getBlob().getDataAsString();
  if (!content) return { seededAt: '', sent: {} };

  let parsed;
  try {
    parsed = JSON.parse(content);
  } catch (e) {
    Logger.error('Welcome-email ledger is corrupt — refusing to run', {
      errorMessage: e.message,
      fileName: WELCOME_AUDIT_CONFIG.LEDGER_FILE_NAME
    });
    throw new Error(
      'Cannot parse ' + WELCOME_AUDIT_CONFIG.LEDGER_FILE_NAME + '. Fix it, or delete it ' +
      'and re-run seedWelcomeLedger() (which re-baselines from login history).'
    );
  }

  if (!parsed || parsed.version !== WELCOME_AUDIT_CONFIG.LEDGER_VERSION) {
    throw new Error(
      WELCOME_AUDIT_CONFIG.LEDGER_FILE_NAME + ' has version ' +
      (parsed ? parsed.version : 'none') + ', expected ' + WELCOME_AUDIT_CONFIG.LEDGER_VERSION +
      '. Refusing to guess at its shape — reading it wrong would either hide real ' +
      'findings or report members who were welcomed.'
    );
  }

  return { seededAt: parsed.seededAt || '', sent: parsed.sent || {} };
}

/**
 * Writes the ledger, creating the file when absent.
 *
 * @param {{seededAt: string, sent: Object}} ledger
 * @returns {void}
 */
function welcomeLedgerSave_(ledger) {
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files = folder.getFilesByName(WELCOME_AUDIT_CONFIG.LEDGER_FILE_NAME);
  const content = JSON.stringify({
    version: WELCOME_AUDIT_CONFIG.LEDGER_VERSION,
    seededAt: ledger.seededAt || '',
    written: new Date().toISOString(),
    sent: ledger.sent || {}
  });

  if (files.hasNext()) {
    files.next().setContent(content);
  } else {
    folder.createFile(WELCOME_AUDIT_CONFIG.LEDGER_FILE_NAME, content);
  }
}

/**
 * Records that a welcome email was sent to this CAPID. Called by
 * sendWelcomeEmail() — so provisioning and the resend are both covered by the
 * one call, and any future sender is covered automatically.
 *
 * Read-modify-write per send, deliberately: accounts are created a handful at a
 * time, and a ledger entry that only exists in memory is lost the moment a run
 * dies mid-way — which would report a genuinely-welcomed member as MISSED.
 *
 * The CALLER must not let a failure here escape: the email has already gone out,
 * and a bookkeeping problem must never turn a successful send into a failed one.
 *
 * @param {string|number} capid
 * @returns {void}
 */
function welcomeLedgerRecordSent_(capid) {
  const ledger = welcomeLedgerLoad_();
  ledger.sent[String(capid).trim()] = {
    on: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    by: 'send'
  };
  welcomeLedgerSave_(ledger);
}

// ============================================================================
// CLASSIFICATION (pure)
// ============================================================================

/**
 * Decides, for every member, whether they appear to have been welcomed. Pure —
 * no Drive, no directory, no sends — so the verdicts are testable without a
 * tenant, and the scan and the notification can never disagree about them.
 *
 * @param {Object} members - CAPID -> member object (getMembers() output)
 * @param {Object} accountByCapid - CAPID -> { email, creationTime, lastLoginTime }
 * @param {{seededAt: string, sent: Object}} ledger
 * @param {Date} now - This run's clock
 * @param {Object} [opts]
 * @param {number} [opts.graceDays] - Override NEW_ACCOUNT_GRACE_DAYS
 * @returns {{missed: Array, unknown: Array, pending: Array, welcomed: number, noAccount: number}}
 *   Each array entry is { capid, name, type, charter, orgName, email, created,
 *   neverSignedIn, reason }, sorted oldest account first.
 */
function classifyWelcomeAudit_(members, accountByCapid, ledger, now, opts) {
  const graceDays = (opts && typeof opts.graceDays === 'number')
    ? opts.graceDays
    : WELCOME_AUDIT_CONFIG.NEW_ACCOUNT_GRACE_DAYS;
  const graceMs = graceDays * 24 * 60 * 60 * 1000;

  const seededAtMs = welcomeAuditTimestamp_(ledger && ledger.seededAt);
  const sent = (ledger && ledger.sent) || {};

  const out = { missed: [], unknown: [], pending: [], welcomed: 0, noAccount: 0 };

  Object.keys(members).forEach(function (capid) {
    const key = String(capid);
    const member = members[key];
    const account = accountByCapid[key];

    // No account is not this module's problem: the member either is not
    // provisioned yet or is held by the Level I gate. Reporting them here would
    // drown the real finding in people who never had an account to welcome.
    if (!account || !account.email) {
      out.noAccount++;
      return;
    }

    if (sent[key]) {
      out.welcomed++;
      return;
    }

    const createdMs = welcomeAuditTimestamp_(account.creationTime);
    const record = {
      capid: key,
      name: [member.rank, member.firstName, member.lastName].filter(String).join(' '),
      type: member.type || '',
      charter: member.charter || '',
      orgName: member.orgName || '',
      email: account.email,
      created: String(account.creationTime || '').slice(0, 10),
      neverSignedIn: welcomeAuditTimestamp_(account.lastLoginTime) <= 0,
      reason: ''
    };

    if (!seededAtMs) {
      // Unseeded ledger: everything is unknowable, and nothing may be accused.
      record.reason = 'no-baseline';
      out.unknown.push(record);
      return;
    }

    if (createdMs <= seededAtMs) {
      // Predates the baseline and the seed did not vouch for it, which means it
      // had no login history when the baseline was taken.
      record.reason = 'predates-baseline';
      out.unknown.push(record);
      return;
    }

    // Created while the ledger was live, and no send was recorded against it.
    if (now.getTime() - createdMs < graceMs) {
      record.reason = 'too-new-to-judge';
      out.pending.push(record);
      return;
    }

    record.reason = 'created-after-baseline-no-send-recorded';
    out.missed.push(record);
  });

  const byCreated = function (a, b) {
    return String(a.created).localeCompare(String(b.created)) ||
           String(a.capid).localeCompare(String(b.capid));
  };
  out.missed.sort(byCreated);
  out.unknown.sort(byCreated);
  out.pending.sort(byCreated);

  return out;
}

/**
 * @param {string} isoTimestamp
 * @returns {number} Milliseconds since epoch, or 0 when absent/unparseable.
 *   Google reports a never-signed-in account's lastLoginTime as the epoch, which
 *   lands on 0 here — exactly "never".
 */
function welcomeAuditTimestamp_(isoTimestamp) {
  const ms = new Date(String(isoTimestamp || '')).getTime();
  return isNaN(ms) ? 0 : ms;
}

/**
 * CAPID -> account, resolving a CAPID held by several accounts to the one most
 * recently signed into — the account the member actually uses. Judging an
 * abandoned twin would report a member who is perfectly fine.
 *
 * @param {Array<Object>} accounts - getActiveUsers() output
 * @returns {Object} CAPID -> account
 */
function welcomeAuditAccountMap_(accounts) {
  const byCapid = {};
  accounts.forEach(function (account) {
    if (!account || !account.capid) return;
    const capid = String(account.capid);
    const existing = byCapid[capid];
    if (!existing ||
        welcomeAuditTimestamp_(account.lastLoginTime) > welcomeAuditTimestamp_(existing.lastLoginTime)) {
      byCapid[capid] = account;
    }
  });
  return byCapid;
}

/**
 * Gathers what the scan and the notification both need. Shared so the two can
 * never be looking at different facts.
 *
 * @returns {{result: Object, ledger: Object, memberCount: number, accountCount: number}}
 */
function welcomeAuditRun_() {
  const ledger = welcomeLedgerLoad_();
  const members = getMembers();
  const accounts = getActiveUsers();
  const accountByCapid = welcomeAuditAccountMap_(accounts);
  const result = classifyWelcomeAudit_(members, accountByCapid, ledger, new Date());

  return {
    result: result,
    ledger: ledger,
    memberCount: Object.keys(members).length,
    accountCount: accounts.length
  };
}

// ============================================================================
// ENTRY POINTS
// ============================================================================

/**
 * ONE-TIME BASELINE. Records every member whose account has been signed into as
 * welcomed, so the scan starts from a truthful position instead of accusing the
 * entire wing.
 *
 * Accounts that have NEVER been signed into are deliberately left out. There is
 * no evidence either way for them, and seeding them would permanently hide the
 * very members this system exists to find. They report as UNKNOWN afterwards —
 * review that list once, resend where appropriate, and it stops mattering.
 *
 * **Defaults to a dry run.** Pass `false` to write.
 *
 * Re-seeding is REFUSED once a baseline exists: moving the baseline forward
 * reclassifies every confirmed MISSED account as merely UNKNOWN, quietly
 * discarding real findings. `seedWelcomeLedger(false, {force: true})` overrides,
 * for the case where the ledger was rebuilt from scratch on purpose.
 *
 * @param {boolean} [dryRun=true] - When true, reports and writes nothing
 * @param {Object} [opts]
 * @param {boolean} [opts.force=false] - Allow re-seeding over an existing baseline
 * @returns {{seeded: number, unknown: number, dryRun: boolean}}
 */
function seedWelcomeLedger(dryRun, opts) {
  const isDryRun = dryRun !== false;
  const force = opts && opts.force === true;
  const ledger = welcomeLedgerLoad_();

  if (ledger.seededAt && !force) {
    console.log('REFUSED: this ledger was already seeded at ' + ledger.seededAt + '.');
    console.log('Re-seeding moves the baseline forward, which turns every confirmed MISSED');
    console.log('account into an UNKNOWN one and discards the findings. Pass {force: true}');
    console.log('only if you rebuilt the ledger deliberately.');
    return { seeded: 0, unknown: 0, dryRun: isDryRun };
  }

  const members = getMembers();
  const accountByCapid = welcomeAuditAccountMap_(getActiveUsers());
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  let seeded = 0;
  let unknown = 0;
  const unknownRows = [];

  Object.keys(members).forEach(function (capid) {
    const key = String(capid);
    const account = accountByCapid[key];
    if (!account || !account.email) return;

    if (welcomeAuditTimestamp_(account.lastLoginTime) > 0) {
      // Signed in, therefore had working credentials at some point.
      ledger.sent[key] = { on: today, by: 'seed' };
      seeded++;
    } else {
      unknown++;
      unknownRows.push({
        capid: key,
        name: [members[key].rank, members[key].firstName, members[key].lastName].filter(String).join(' '),
        charter: members[key].charter || '',
        email: account.email,
        created: String(account.creationTime || '').slice(0, 10)
      });
    }
  });

  console.log('Welcome-email ledger seed' + (isDryRun ? ' — DRY RUN, nothing written' : ''));
  console.log('  Seeded as welcomed (account has login history): ' + seeded);
  console.log('  Left UNKNOWN (never signed in, cannot be judged): ' + unknown);

  if (unknownRows.length) {
    console.log('');
    console.log('  These are not accusations — a member who was welcomed and never logged in');
    console.log('  looks identical here. Review, and use resendWelcomeEmail(capid) where it fits:');
    unknownRows
      .sort(function (a, b) { return String(a.created).localeCompare(String(b.created)); })
      .forEach(function (r) {
        console.log('    ' + r.capid.padEnd(8) + ' created ' + (r.created || '?????????').padEnd(12) +
          r.charter.padEnd(14) + r.email);
      });
  }

  if (isDryRun) {
    console.log('');
    console.log('Nothing was written. Re-run as seedWelcomeLedger(false) to establish the baseline.');
    return { seeded: seeded, unknown: unknown, dryRun: true };
  }

  ledger.seededAt = new Date().toISOString();
  welcomeLedgerSave_(ledger);

  Logger.info('Welcome-email ledger seeded', {
    seeded: seeded,
    unknown: unknown,
    seededAt: ledger.seededAt
  });
  console.log('');
  console.log('Baseline set at ' + ledger.seededAt + '. From here, any account created without a');
  console.log('recorded welcome email is reported by scanUnwelcomedAccounts().');

  return { seeded: seeded, unknown: unknown, dryRun: false };
}

/**
 * READ-ONLY. Reports members holding an account with no recorded welcome email.
 * Changes nothing and sends nothing.
 *
 * @returns {Object} The classification result, for programmatic callers
 */
function scanUnwelcomedAccounts() {
  const run = welcomeAuditRun_();
  const r = run.result;

  console.log('Welcome-email audit (read-only)');
  console.log('  Baseline:  ' + (run.ledger.seededAt || 'NOT SEEDED — run seedWelcomeLedger() first'));
  console.log('  Members:   ' + run.memberCount + ' eligible, ' + run.accountCount + ' active accounts scanned');
  console.log('  Welcomed:  ' + r.welcomed);
  console.log('  No account: ' + r.noAccount + ' (provisioning/Level I, not this module)');
  console.log('');

  console.log('MISSED — created after the baseline, no welcome email recorded: ' + r.missed.length);
  r.missed.forEach(function (m) {
    console.log('  ' + m.capid.padEnd(8) + (m.created || '?').padEnd(12) + m.charter.padEnd(14) +
      m.email + (m.neverSignedIn ? '' : '  [has since signed in]'));
  });
  if (r.missed.length) {
    console.log('');
    console.log('  Fix each with: previewWelcomeEmailResend(capid) then resendWelcomeEmail(capid)');
    console.log('  An entry marked [has since signed in] got credentials some other way —');
    console.log('  the resend will refuse it, which is correct. Nothing to do for those.');
  }

  console.log('');
  console.log('PENDING — too new to judge (< ' + WELCOME_AUDIT_CONFIG.NEW_ACCOUNT_GRACE_DAYS +
    ' days old): ' + r.pending.length);

  console.log('');
  console.log('UNKNOWN — predates the baseline, never signed into: ' + r.unknown.length);
  console.log('  A genuine maybe. Not mailed to IT, and not a to-do list: a member who was');
  console.log('  welcomed and simply never logged in is indistinguishable from one who was');
  console.log('  never welcomed at all. Unit command staff already see this population via');
  console.log('  the monthly recovery-compliance digest.');
  r.unknown.forEach(function (m) {
    console.log('  ' + m.capid.padEnd(8) + (m.created || '?').padEnd(12) + m.charter.padEnd(14) + m.email);
  });

  return r;
}

/**
 * Emails IT the MISSED list — accounts created since the baseline with no
 * welcome email recorded against them. Silent when there is nothing to report,
 * so it is safe on a trigger.
 *
 * UNKNOWN is deliberately not mailed: it cannot be acted on without judgement,
 * and a monthly message that is mostly noise stops being read.
 *
 * @returns {{reported: number, sent: boolean}}
 */
function notifyUnwelcomedAccounts() {
  const run = welcomeAuditRun_();
  const missed = run.result.missed;

  if (!run.ledger.seededAt) {
    Logger.warn('Welcome-email audit skipped — ledger has no baseline', {
      remedy: 'Run seedWelcomeLedger(false) once'
    });
    return { reported: 0, sent: false };
  }

  if (!missed.length) {
    Logger.info('Welcome-email audit — nothing to report', {
      welcomed: run.result.welcomed,
      unknown: run.result.unknown.length
    });
    return { reported: 0, sent: false };
  }

  const rows = missed.map(function (m) {
    return '<tr>' +
      '<td>' + welcomeAuditEscapeHtml_(m.capid) + '</td>' +
      '<td>' + welcomeAuditEscapeHtml_(m.name) + '</td>' +
      '<td>' + welcomeAuditEscapeHtml_(m.charter) + '</td>' +
      '<td>' + welcomeAuditEscapeHtml_(m.email) + '</td>' +
      '<td>' + welcomeAuditEscapeHtml_(m.created) + '</td>' +
      '<td>' + (m.neverSignedIn ? 'never signed in' : 'has since signed in') + '</td>' +
      '</tr>';
  }).join('');

  const html =
    '<p>These Workspace accounts were created after the welcome-email ledger baseline ' +
    '(' + welcomeAuditEscapeHtml_(run.ledger.seededAt.slice(0, 10)) + ') and no welcome email was ever ' +
    'recorded for them. Each was almost certainly created outside provisioning — in the ' +
    'Admin console or by GAM — which never sends one. A failed send would also land here.</p>' +
    '<table border="1" cellpadding="6" cellspacing="0">' +
    '<tr><th>CAPID</th><th>Member</th><th>Unit</th><th>Account</th><th>Created</th><th>Sign-in</th></tr>' +
    rows +
    '</table>' +
    '<p><strong>To fix:</strong> run <code>previewWelcomeEmailResend(capid)</code>, then ' +
    '<code>resendWelcomeEmail(capid)</code>. That resets the account password and mails ' +
    'the welcome email carrying the new one.</p>' +
    '<p>Rows marked <em>has since signed in</em> obtained credentials some other way. The ' +
    'resend will refuse them, which is correct — resetting a working account would lock ' +
    'the member out. Nothing needs doing for those.</p>';

  MailApp.sendEmail({
    to: ITSUPPORT_EMAIL,
    subject: WELCOME_AUDIT_CONFIG.SUBJECT + ' (' + missed.length + ')',
    htmlBody: html
  });

  Logger.info('Welcome-email audit reported to IT', {
    reported: missed.length,
    to: ITSUPPORT_EMAIL
  });

  return { reported: missed.length, sent: true };
}

/**
 * Installs the monthly audit. Replaces any existing trigger for the same
 * handler, so running it twice does not double up.
 *
 * Run this **as the automation account** — triggers are owned by and visible
 * only to their creator, and the report is sent by whoever owns the trigger.
 *
 * @returns {void}
 */
function installWelcomeAuditMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'notifyUnwelcomedAccounts') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('notifyUnwelcomedAccounts')
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();

  Logger.info('Installed monthly welcome-email audit trigger', {
    handler: 'notifyUnwelcomedAccounts',
    schedule: '1st of each month ~08:00 America/Los_Angeles',
    note: 'Confirm in the Triggers panel that the owner is the automation account'
  });
}

/**
 * @param {*} value
 * @returns {string} HTML-escaped
 */
function welcomeAuditEscapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
