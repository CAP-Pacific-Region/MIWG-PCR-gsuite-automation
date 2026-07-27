/*******************************************************
 * Bad Parent-Contact Email Notification Module
 *
 * Version: 1.0.0
 * Filename: ParentEmailNotify.gs
 * Saved: 2026-07-27
 *
 * Mails a unit's Commander and Deputy Commander for Cadets the parent/guardian
 * email addresses on their cadets' records that GOOGLE WILL NOT ACCEPT, so the
 * unit can correct them in eServices.
 *
 * WHY THE DATA COMES FROM A LEDGER RATHER THAN A CHECK
 * An address is only discovered to be bad by asking Google. A well-formed
 * gmail.com address that simply does not exist looks identical to a good one
 * until members.insert refuses it with 404 "Resource Not Found" — Google
 * verifies addresses in its own domain, and there is no read-only way to ask.
 * The squadron sync is therefore the only place the truth is ever known, and it
 * writes what it learned to RejectedMemberAddresses.json. This module reads that
 * and mails nobody anything it did not observe.
 *
 * Consequences worth knowing:
 *   - Until updateAllSquadronGroups() has run at least once, the ledger does not
 *     exist and this reports nothing. That is deliberate: an empty ledger means
 *     "not yet looked", not "nothing wrong", and the two must not be confused.
 *   - Only addresses on lists the sync manages are ever tested. A cadet in a unit
 *     with no managed .parents list is not covered.
 *   - A corrected address stops being rejected, so it drops out of the ledger on
 *     the next complete run and stops being reported without anyone pruning it.
 *
 * SUPPRESSION is per member AND per address, not per member. A cadet with two
 * bad parent addresses who has one fixed should still be told about the other;
 * keying on the member alone would hide it for three months.
 *
 * RECIPIENTS are the Commander and the Deputy Commander for Cadets, and nobody
 * else. On the cadets tenant both are SENIOR members whose accounts live on the
 * senior domain — recipient resolution is shared with RecoveryEmailNotify.gs,
 * which already solves that cross-tenant lookup.
 *******************************************************/

const PARENT_EMAIL_NOTIFY_CONFIG = {
  // Written by SquadronGroups.gs. This module only ever reads it.
  LEDGER_FILE_NAME: 'RejectedMemberAddresses.json',

  // Own state file. Nothing else writes it: this suppression window is the only
  // thing between a standing data problem and a monthly nag.
  STATE_FILE_NAME: 'ParentEmailNotifyState.txt',
  STATE_VERSION: 1,

  // Do not report the same member/address pair again until this many months
  // have passed.
  SUPPRESSION_MONTHS: 3,

  // Exactly the two roles asked for. Deliberately narrower than
  // RECOVERY_NOTIFY_CONFIG.RECIPIENT_DUTY_TITLES: a parent contact belongs to a
  // cadet, so the cadet-side command staff are the ones who act on it.
  RECIPIENT_ROLES: ['Commander', 'Deputy Commander for Cadets'],

  SUBJECT: 'Parent contact emails that need correcting in your unit',

  // Where the correction is made.
  SELF_SERVICE_URL: 'https://www.capnhq.gov/CAP.PersonnelInfo.Web/',

  // Spacing between unit digests, matching the other notification modules.
  EMAIL_DELAY_MS: 1000
};

// ============================================================================
// MAIN
// ============================================================================

/**
 * Reports unusable parent-contact addresses to unit command staff and records
 * who was told.
 *
 * @returns {Object} Summary of the run
 */
function notifyBadParentEmails() {
  return runParentEmailNotification_({ dryRun: false });
}

/**
 * Dry run: reports what notifyBadParentEmails() would send, without sending any
 * mail or touching the state file. Safe to run at any time.
 *
 * @returns {Object} Summary of the run
 */
function previewBadParentEmails() {
  return runParentEmailNotification_({ dryRun: true });
}

/**
 * @param {Object} options - { dryRun: boolean }
 * @returns {Object} Summary of the run
 */
function runParentEmailNotification_(options) {
  const dryRun = !!(options && options.dryRun);
  const summary = {
    dryRun: dryRun,
    ledgerAddresses: 0,
    resolved: 0,
    suppressed: 0,
    unresolved: 0,
    units: 0,
    sent: 0,
    failed: 0
  };

  if (!PROFILE_.RUN_PARENT_EMAIL_NOTIFICATIONS) {
    Logger.info('Parent-email notifications are off for this tenant — nothing to do');
    return summary;
  }

  const ledger = peLoadLedger_();
  summary.ledgerAddresses = ledger.length;
  if (ledger.length === 0) {
    Logger.info('No rejected parent addresses on record — nothing to report', {
      note: 'The ledger is written by updateAllSquadronGroups(). No ledger means not yet looked.'
    });
    return summary;
  }

  const byOrg = peResolveToUnits_(ledger, summary);
  const orgids = Object.keys(byOrg);
  if (orgids.length === 0) {
    Logger.info('No rejected address resolved to a current cadet — nothing to report', {
      ledgerAddresses: ledger.length,
      unresolved: summary.unresolved
    });
    return summary;
  }

  const today = peIsoDate_(new Date());
  const state = peLoadState_();
  const reportable = {};

  orgids.forEach(function (orgid) {
    const rows = byOrg[orgid].filter(function (row) {
      if (peIsSuppressed_(state[row.key], today)) {
        summary.suppressed++;
        return false;
      }
      return true;
    });
    if (rows.length) reportable[orgid] = rows;
  });

  const reportableOrgs = Object.keys(reportable);
  summary.units = reportableOrgs.length;

  if (reportableOrgs.length === 0) {
    Logger.info('Every unusable parent address is inside its suppression window', {
      suppressed: summary.suppressed
    });
    return summary;
  }

  const recipientsByOrg = peBuildRecipients_();

  reportableOrgs.sort().forEach(function (orgid, index) {
    const rows = reportable[orgid];
    const recipients = recipientsByOrg[orgid] || [];

    if (recipients.length === 0) {
      // No commander and no DCC on record. Saying so is the useful outcome —
      // silently dropping the unit would make the digest look complete.
      Logger.warn('No command staff on record for a unit with unusable parent addresses', {
        orgid: orgid,
        charter: rows[0].charter,
        addresses: rows.length
      });
      summary.failed++;
      return;
    }

    if (dryRun) {
      Logger.info('💡 [Dry-Run] Would send parent-email digest', {
        orgid: orgid,
        charter: rows[0].charter,
        to: recipients.map(r => r.email),
        addresses: rows.map(r => ({ cadet: r.capid, address: r.address, reason: r.reason }))
      });
      summary.sent++;
      return;
    }

    if (peSendDigest_(orgid, recipients, rows)) {
      summary.sent++;
      rows.forEach(function (row) { state[row.key] = today; });
    } else {
      summary.failed++;
    }

    if (index < reportableOrgs.length - 1) {
      Utilities.sleep(PARENT_EMAIL_NOTIFY_CONFIG.EMAIL_DELAY_MS);
    }
  });

  if (!dryRun) peSaveState_(state);

  Logger.info('Parent-email notification run complete', summary);
  return summary;
}

// ============================================================================
// LEDGER AND RESOLUTION
// ============================================================================

/**
 * Reads the rejected-address ledger written by the squadron sync.
 *
 * A missing ledger is not an error — it means the sync has not run here yet.
 * A corrupt one IS: silently returning nothing would look exactly like "no
 * problems found", which is the one wrong answer this module can give.
 *
 * @returns {Array<Object>} [{ member, reason, groups, firstSeen, lastSeen }]
 */
function peLoadLedger_() {
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files = folder.getFilesByName(PARENT_EMAIL_NOTIFY_CONFIG.LEDGER_FILE_NAME);
  if (!files.hasNext()) return [];

  const content = files.next().getBlob().getDataAsString();
  if (!content) return [];

  let parsed;
  try {
    parsed = JSON.parse(content);
  } catch (e) {
    Logger.error('Rejected-address ledger is corrupt — refusing to run', {
      errorMessage: e.message,
      fileName: PARENT_EMAIL_NOTIFY_CONFIG.LEDGER_FILE_NAME
    });
    throw new Error(
      'Cannot parse ' + PARENT_EMAIL_NOTIFY_CONFIG.LEDGER_FILE_NAME +
      '. Delete it and re-run updateAllSquadronGroups() to rebuild it.'
    );
  }

  return (parsed && parsed.addresses) || [];
}

/**
 * Maps each rejected address back to the cadets whose record carries it, and
 * those cadets to their unit.
 *
 * The ledger records an ADDRESS, not a person — the squadron sync only ever knew
 * the address it was refused. The link back to a cadet lives in MbrContact, so
 * it is rebuilt here against the current extract. That has a useful side effect:
 * a cadet who has since left, or whose contact row was corrected, resolves to
 * nothing and is not reported.
 *
 * CADETS COME FROM THE RAW EXTRACT, NOT getMembers(). On the cadets tenant
 * CADET_LITE filters every cadet below the account-holding grade out of
 * getMembers() — on CAWG that is over 1,600 of them, and they are exactly the
 * population this digest exists for: a cadet-lite member holds no account, so
 * their parent's address is the ONLY way the unit's list reaches that family.
 * Resolving against the filtered set silently dropped ten of thirteen addresses
 * on the first live preview. Reading Member.txt directly is what makes them
 * visible, the same reason rcBuildRecipientDirectory_ reads it raw.
 *
 * One address can belong to several cadets — siblings share a parent — so this
 * can produce more rows than the ledger has addresses.
 *
 * @param {Array<Object>} ledger - peLoadLedger_() output
 * @param {Object} summary - Mutated with resolved/unresolved counts
 * @returns {Object} ORGID -> [{ key, capid, cadetName, address, reason, firstSeen, orgid, charter }]
 */
function peResolveToUnits_(ledger, summary) {
  const wanted = {};
  ledger.forEach(function (row) {
    const address = String(row.member || '').trim().toLowerCase();
    if (address) wanted[address] = row;
  });

  // Member.txt columns: [0] CAPID, [2] last, [3] first, [11] ORGID, [13] unit,
  // [21] type, [24] status — the same layout buildCadetLiteExcludedCadetsByOrgid_
  // reads in SquadronGroups.gs.
  const cadets = {};
  parseFile('Member').forEach(function (row) {
    const capid = String(row[0] || '').trim();
    const orgid = String(row[11] || '').trim();
    const unit = String(row[13] || '').trim();
    const type = String(row[21] || '').trim().toUpperCase();
    const status = String(row[24] || '').trim().toUpperCase();

    if (!capid || !orgid) return;
    if (status !== 'ACTIVE') return;
    if (type !== 'CADET') return;
    if (unit === '0' || unit === '000' || unit === '999') return;

    cadets[capid] = {
      orgid: orgid,
      lastName: String(row[2] || '').trim(),
      firstName: String(row[3] || '').trim()
    };
  });

  const squadrons = getSquadrons();
  const byOrg = {};
  const matchedAddresses = {};

  parseFile('MbrContact').forEach(function (row) {
    const capid = String(row[0] || '').trim();
    const contactType = String(row[1] || '').trim();
    const value = String(row[3] || '').trim().toLowerCase();
    const doNotContact = String(row[6] || '').trim();

    if (!capid || contactType !== 'CADET PARENT EMAIL') return;
    if (doNotContact === 'True') return;
    if (!wanted[value]) return;

    const member = cadets[capid];
    if (!member) return;                       // no longer an active cadet here

    const orgid = member.orgid;

    const squadron = squadrons[orgid];
    matchedAddresses[value] = true;

    if (!byOrg[orgid]) byOrg[orgid] = [];
    byOrg[orgid].push({
      key: capid + '|' + value,
      capid: capid,
      cadetName: [member.firstName, member.lastName].filter(Boolean).join(' ').trim() || capid,
      address: value,
      reason: wanted[value].reason || 'rejected',
      firstSeen: wanted[value].firstSeen || '',
      orgid: orgid,
      charter: (squadron && squadron.charter) || orgid,
      orgName: (squadron && squadron.name) || ''
    });

    summary.resolved++;
  });

  summary.unresolved = Object.keys(wanted).filter(a => !matchedAddresses[a]).length;

  if (summary.unresolved > 0) {
    // Usually benign — the address was corrected, or the cadet left — but a
    // large number would mean the ledger has gone stale against the extract.
    Logger.info('Rejected addresses that no longer match an active cadet', {
      count: summary.unresolved,
      note: 'Corrected in eServices, or the cadet is no longer in this tenant.'
    });
  }

  return byOrg;
}

// ============================================================================
// RECIPIENTS
// ============================================================================

/**
 * ORGID -> [{ capid, firstName, lastName, rank, role, email }] for the two roles
 * this digest addresses.
 *
 * Built on RecoveryEmailNotify.gs's directory, which already handles the part
 * that is genuinely hard: on the cadets tenant a cadet unit's commander is a
 * SENIOR, absent from getMembers() and holding an account on the other domain.
 * Duplicating that would mean maintaining two copies of the same subtle lookup.
 *
 * @returns {Object} ORGID to an array of recipients
 */
function peBuildRecipients_() {
  if (typeof rcBuildRecipientDirectory_ !== 'function') {
    throw new Error('RecoveryEmailNotify.gs is required for recipient resolution');
  }

  const roles = PARENT_EMAIL_NOTIFY_CONFIG.RECIPIENT_ROLES;
  const full = rcBuildRecipientDirectory_(getActiveUsers());
  const filtered = {};

  Object.keys(full).forEach(function (orgid) {
    const keep = (full[orgid] || []).filter(r => roles.indexOf(r.role) > -1);
    if (keep.length) filtered[orgid] = keep;
  });

  return filtered;
}

// ============================================================================
// SUPPRESSION STATE
// ============================================================================

/**
 * @returns {Object} 'capid|address' -> 'yyyy-MM-dd' last reported
 */
function peLoadState_() {
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files = folder.getFilesByName(PARENT_EMAIL_NOTIFY_CONFIG.STATE_FILE_NAME);
  if (!files.hasNext()) {
    Logger.info('No parent-email notify state yet — every unusable address is reportable');
    return {};
  }

  const content = files.next().getBlob().getDataAsString();
  if (!content) return {};

  let parsed;
  try {
    parsed = JSON.parse(content);
  } catch (e) {
    // Returning {} would silently drop every suppression window and re-mail the
    // whole wing. Fail instead.
    Logger.error('Parent-email notify state is corrupt — refusing to run', {
      errorMessage: e.message,
      fileName: PARENT_EMAIL_NOTIFY_CONFIG.STATE_FILE_NAME
    });
    throw new Error(
      'Cannot parse ' + PARENT_EMAIL_NOTIFY_CONFIG.STATE_FILE_NAME + '. Fix or delete it ' +
      '(deleting re-reports every unusable address) before running again.'
    );
  }

  if (!parsed || parsed.version !== PARENT_EMAIL_NOTIFY_CONFIG.STATE_VERSION) {
    Logger.warn('Parent-email notify state version not recognised — starting fresh', {
      found: parsed ? parsed.version : null,
      expected: PARENT_EMAIL_NOTIFY_CONFIG.STATE_VERSION
    });
    return {};
  }

  return parsed.reported || {};
}

/**
 * @param {Object} reported - 'capid|address' -> 'yyyy-MM-dd'
 * @returns {void}
 */
function peSaveState_(reported) {
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files = folder.getFilesByName(PARENT_EMAIL_NOTIFY_CONFIG.STATE_FILE_NAME);
  const content = JSON.stringify({
    version: PARENT_EMAIL_NOTIFY_CONFIG.STATE_VERSION,
    written: peIsoDate_(new Date()),
    reported: reported
  });

  if (files.hasNext()) files.next().setContent(content);
  else folder.createFile(PARENT_EMAIL_NOTIFY_CONFIG.STATE_FILE_NAME, content);

  Logger.info('Parent-email notify state saved', {
    entries: Object.keys(reported).length,
    fileName: PARENT_EMAIL_NOTIFY_CONFIG.STATE_FILE_NAME
  });
}

/**
 * Discards the recorded state. The next run reports every unusable address
 * again, ignoring any suppression window.
 *
 * @returns {void}
 */
function resetParentEmailNotifyState() {
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files = folder.getFilesByName(PARENT_EMAIL_NOTIFY_CONFIG.STATE_FILE_NAME);
  let removed = 0;
  while (files.hasNext()) {
    files.next().setTrashed(true);
    removed++;
  }
  Logger.info('Parent-email notify state reset — next run reports everything again', {
    filesTrashed: removed
  });
}

/**
 * Suppression is measured in CALENDAR MONTHS, not days: a member reported on
 * the 3rd becomes eligible on the 3rd, three months later, which is what a
 * monthly cadence implies. Date rollover is handled by the Date constructor.
 *
 * @param {string} lastNotifiedIso - Date last reported, 'yyyy-MM-dd'
 * @param {string} todayIso - This run's date, 'yyyy-MM-dd'
 * @returns {boolean} True if this member/address should be skipped
 */
function peIsSuppressed_(lastNotifiedIso, todayIso) {
  const last = peParseIsoDate_(lastNotifiedIso);
  const today = peParseIsoDate_(todayIso);
  if (!last || !today) return false;

  const eligibleAgain = new Date(
    last.getFullYear(),
    last.getMonth() + PARENT_EMAIL_NOTIFY_CONFIG.SUPPRESSION_MONTHS,
    last.getDate()
  );

  return today < eligibleAgain;
}

/**
 * @param {Date} date - Date to format
 * @returns {string} 'yyyy-MM-dd' in the script's timezone
 */
function peIsoDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * @param {string} iso - 'yyyy-MM-dd'
 * @returns {Date|null} Parsed date, or null if unparseable
 */
function peParseIsoDate_(iso) {
  const parts = String(iso || '').split('-');
  if (parts.length !== 3) return null;
  const year = Number(parts[0]);
  const month = Number(parts[1]);
  const day = Number(parts[2]);
  if (!isFinite(year) || !isFinite(month) || !isFinite(day)) return null;
  return new Date(year, month - 1, day);
}

// ============================================================================
// SENDING
// ============================================================================

/**
 * Sends one unit's digest. The commander is the addressee; the Deputy Commander
 * for Cadets is copied.
 *
 * @param {string} orgid - Unit ORGID
 * @param {Array<Object>} recipients - Command staff for this unit
 * @param {Array<Object>} rows - Unusable addresses for this unit
 * @returns {boolean} True if sent
 */
function peSendDigest_(orgid, recipients, rows) {
  const selected = rcSelectAddressees_(recipients);
  const addressee = selected.addressee;
  const cc = selected.cc;
  const subject = PARENT_EMAIL_NOTIFY_CONFIG.SUBJECT + ' — ' + rows[0].charter;

  try {
    const options = {
      htmlBody: peBuildDigestHtml_(addressee, rows),
      from: AUTOMATION_SENDER_EMAIL,
      name: SENDER_NAME,
      replyTo: ITSUPPORT_EMAIL
    };
    if (cc.length) options.cc = cc.join(',');

    executeWithRetry(() =>
      GmailApp.sendEmail(addressee.email, subject, 'See the HTML version of this message.', options)
    );

    Logger.info('Parent-email digest sent', {
      to: addressee.email,
      cc: cc.length,
      orgid: orgid,
      charter: rows[0].charter,
      addresses: rows.length
    });
    return true;

  } catch (e) {
    Logger.error('Failed to send parent-email digest', {
      to: addressee.email,
      orgid: orgid,
      charter: rows[0].charter,
      errorMessage: e.message,
      errorCode: e.details?.code
    });
    return false;
  }
}

/**
 * Renders one unit's digest.
 *
 * The address is shown verbatim, because the recipient's job is to compare it
 * against what the parent actually uses and spot the difference — a masked or
 * summarised address would make the email useless.
 *
 * @param {Object} addressee - Recipient the digest is addressed to
 * @param {Array<Object>} rows - Unusable addresses for this unit
 * @returns {string} HTML body
 */
function peBuildDigestHtml_(addressee, rows) {
  const esc = rcEscapeHtml_;
  const url = PARENT_EMAIL_NOTIFY_CONFIG.SELF_SERVICE_URL;
  const anyMalformed = rows.some(r => r.reason === 'malformed');
  const anyMissing = rows.some(r => r.reason !== 'malformed');

  let html = '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222">';
  html += '<p>' + esc(addressee.rank ? addressee.rank + ' ' + addressee.lastName : addressee.lastName) + ',</p>';

  html += '<p>The parent/guardian email addresses below are on your cadets\' records in eServices, ' +
    'but <strong>Google will not accept them</strong>. Mail sent to your unit\'s parents list does ' +
    'not reach these families, and will not until the address is corrected.</p>';

  html += '<table cellpadding="6" cellspacing="0" border="0" style="border-collapse:collapse;font-size:13px">';
  html += '<tr style="background:#f2f2f2">' +
    '<th align="left">Cadet</th><th align="left">CAPID</th>' +
    '<th align="left">Address on file</th><th align="left">Problem</th></tr>';

  rows.forEach(function (row, i) {
    const shade = i % 2 ? '#ffffff' : '#fafafa';
    const problem = row.reason === 'malformed'
      ? 'Not a valid address'
      : 'No such Google account';
    html += '<tr style="background:' + shade + '">' +
      '<td>' + esc(row.cadetName) + '</td>' +
      '<td>' + esc(row.capid) + '</td>' +
      '<td><code>' + esc(row.address) + '</code></td>' +
      '<td>' + esc(problem) + '</td></tr>';
  });
  html += '</table>';

  html += '<p><strong>What to do:</strong> confirm the correct address with the family, then update ' +
    'the cadet\'s parent contact in <a href="' + url + '">eServices</a>. The correction is picked up ' +
    'automatically — nothing needs to be sent back to us.</p>';

  if (anyMissing) {
    html += '<p><em>“No such Google account”</em> means Google checked and found no account at that ' +
      'address. Usually a typo, or an address the family stopped using.</p>';
  }
  if (anyMalformed) {
    html += '<p><em>“Not a valid address”</em> means the address cannot be delivered to as written — ' +
      'a doubled dot, a stray character, or a Gmail username containing something Gmail does not allow ' +
      '(only letters, digits and dots are permitted, and <code>+tags</code> cannot be used here).</p>';
  }

  html += '<p style="color:#666;font-size:12px">You are receiving this as the unit Commander or Deputy ' +
    'Commander for Cadets. Each address is reported at most once every ' +
    PARENT_EMAIL_NOTIFY_CONFIG.SUPPRESSION_MONTHS + ' months. Questions: ' +
    '<a href="mailto:' + esc(ITSUPPORT_EMAIL) + '">' + esc(ITSUPPORT_EMAIL) + '</a>.</p>';
  html += '</div>';

  return html;
}

// ============================================================================
// TRIGGER
// ============================================================================

/**
 * Installs the monthly trigger, replacing any existing one for this handler.
 *
 * ⚠ RUN THIS AS THE AUTOMATION ACCOUNT. Triggers are owned by whoever creates
 * them, and the digest sends with the automation account's Send-As alias — a
 * trigger owned by anyone else fails on every send.
 *
 * @returns {void}
 */
function installParentEmailMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'notifyBadParentEmails') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('notifyBadParentEmails')
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();

  Logger.info('Installed monthly parent-email trigger', {
    handler: 'notifyBadParentEmails',
    schedule: '1st of each month ~08:00 America/Los_Angeles',
    note: 'Confirm in the Triggers panel that the owner is the automation account'
  });
}
