/**
 * Recovery-Email Compliance Notification
 *
 * Version: 1.0.1
 * Date: 2026-07-19
 * Changes: Resolve recipients from this tenant's own directory before deriving
 *   the address, and share one addressee/Cc reduction between the send and the
 *   preview. Both came out of the first live preview: the preview printed the
 *   raw duty list (one person appeared four times in a unit that holds several
 *   duties), and derivation alone produced five classes of dead address —
 *   an apostrophe stripped in the real account, a '.3' duplicate-account suffix,
 *   a CAPWATCH nickname vs the legal first name, a surname changed since the
 *   account was created, and a middle name concatenated into the first name.
 *   (1.0.0: new module. Mails each unit's commander, personnel officer(s) and
 *   deputy commander(s) a monthly digest of members under their direct command
 *   whose CAPWATCH email setup would leave them unable to reset their Workspace
 *   password.) See PCR_CHANGELOG.md.
 *
 * Authors: Isaac Wilson IV
 *
 * WHAT THIS REPORTS
 * A member resets their Workspace password through a personal, non-tenant
 * recovery address. You cannot recover an account from the account it locks, so
 * a member whose only address on file is their CAP address has no way back in.
 * The wing's intended CAPWATCH setup is:
 *
 *   PRIMARY   = the member's CAP (tenant) address
 *   SECONDARY = a personal, non-CAP address
 *
 * Two independent conditions are reported, and a member can trip either or both:
 *
 *   flagPrimary  - there is no tenant address in the PRIMARY slot. That covers a
 *                  personal address sitting in PRIMARY *and* no PRIMARY at all.
 *   flagRecovery - there is no usable personal recovery address anywhere.
 *
 * flagRecovery deliberately reuses `member.recoveryEmail`, which
 * UpdateMembers.gs already derives as
 *   firstPersonalEmail_([secondaryEmail, primaryEmailValue]) || parentEmail || ''
 * That derivation is the same one that populates the account's real Workspace
 * recovery address, so this module reports on exactly what recovery will use
 * rather than on a second, subtly different opinion of it. Two consequences are
 * intentional:
 *
 *   - A CADET whose own secondary is empty but who has a parent/guardian email on
 *     file is NOT flagged for recovery. That parent address is a working recovery
 *     path, and flagging it would bury the real gaps under most of the cadet wing.
 *   - A member with a personal address in PRIMARY and nothing in SECONDARY is not
 *     flagged for recovery (they can recover), but IS flagged on primary — which
 *     is the correction that actually needs making.
 *
 * WHY IT IS NOT A CHANGE DETECTOR (and why the first run is loud)
 * LSCodeNotify.gs diffs a value against recorded state and is silent on its first
 * run by construction. This module is the opposite kind of thing: it reports a
 * STANDING condition, so the first run legitimately surfaces the entire existing
 * backlog. That is the point of it — but it means the first real run can mail a
 * lot of units at once. Run previewRecoveryEmailCompliance() first and look at
 * the volume before arming anything.
 *
 * THE THREE-MONTH RULE
 * A standing condition would otherwise re-mail the same commander about the same
 * member every single month until the member acts, which is how a notification
 * gets filtered to trash. State records the date each member was last reported,
 * and a member is skipped until SUPPRESSION_MONTHS have passed. A member who
 * becomes compliant is dropped from state entirely, so if they later regress they
 * are reported on the next run rather than sitting silently inside a stale
 * suppression window.
 *
 * WHO GETS MAILED
 * The unit's commander (To), with its personnel officers and deputy commanders
 * (Cc) — primary and assistant alike. "Direct command" means the member's own
 * ORGID: a member is reported to their own unit's staff, never to a parent org.
 *
 * Recipient addresses are the CAP account (first.last@<command domain>), falling
 * back to the CAPWATCH PRIMARY address only when a name is unusable. This is
 * deliberate: the whole premise of the module is that CAPWATCH primary addresses
 * are often wrong or personal, so it should not depend on one to reach the
 * command staff it is reporting to.
 *
 * CONFIG.COMMAND_EMAIL_DOMAIN exists for the cadets tenant. Command staff are
 * senior members, so on a cadet tenant their account lives on the SENIOR domain
 * (@cawgcap.org), not the tenant's own (@cawgcadets.org). It defaults to
 * CONFIG.EMAIL_DOMAIN, which is correct for the seniors tenant; set
 * TENANT_COMMAND_EMAIL_DOMAIN on cadets. Note the address is derived, not
 * verified against a directory — a senior with no account, or an unusual account
 * name, yields an address that silently goes nowhere.
 *
 * SENDING IDENTITY — this bites, so read it before scheduling.
 * Each digest sets `from: AUTOMATION_SENDER_EMAIL`, and Gmail permits that only
 * when the *executing* account owns that address as a verified Send-As alias. The
 * monthly trigger must therefore be owned by the automation account, exactly as
 * LSCodeNotify.gs documents. The IT failure-summary is deliberately sent WITHOUT
 * `from` so that this specific misconfiguration still gets reported rather than
 * failing silently along with everything else.
 *
 * Setup:
 * 1. Set TENANT_PROFILE; this runs where RUN_RECOVERY_EMAIL_NOTIFICATIONS is true
 *    (seniors and cadets; see config.gs).
 * 2. On the cadets tenant, set TENANT_COMMAND_EMAIL_DOMAIN=@cawgcap.org first.
 * 3. Run previewRecoveryEmailCompliance() — sends nothing, writes nothing. Check
 *    the volume and that the recipient addresses look right.
 * 4. Run notifyRecoveryEmailCompliance() once by hand, as the automation account.
 * 5. As the automation account, installRecoveryComplianceMonthlyTrigger().
 */

const RECOVERY_NOTIFY_CONFIG = {
  // Own state file. Nothing else may write it: this module's suppression window
  // is the only thing standing between a standing condition and a monthly nag.
  STATE_FILE_NAME: 'RecoveryComplianceState.txt',

  // Bumped when the state file's shape changes. An unrecognised version is
  // re-baselined rather than guessed at — see rcLoadState_.
  STATE_VERSION: 1,

  // Do not report the same member again until this many months have passed.
  SUPPRESSION_MONTHS: 3,

  // Duty positions that receive the digest alongside the commander. Assistants
  // are included: the assistant personnel officer is frequently the person who
  // actually does this work.
  RECIPIENT_DUTY_TITLES: [
    'Personnel Officer',
    'Deputy Commander',
    'Deputy Commander for Seniors',
    'Deputy Commander for Cadets'
  ],

  SUBJECT: 'Member email records needing correction in your unit',

  // Where a commander or personnel officer can make the correction themselves.
  SELF_SERVICE_URL: 'https://www.capnhq.gov/CAP.PersonnelInfo.Web/',

  // Spacing between unit digests, matching RETENTION_CONFIG.EMAIL_DELAY_MS.
  EMAIL_DELAY_MS: 1000
};

// ============================================================================
// MAIN
// ============================================================================

/**
 * Reports non-compliant members to their unit's command staff and records who
 * was told, so nobody is reported again for SUPPRESSION_MONTHS.
 *
 * @returns {Object} Summary of the run
 */
function notifyRecoveryEmailCompliance() {
  return runRecoveryNotification_({ dryRun: false });
}

/**
 * Dry run: reports what notifyRecoveryEmailCompliance() would send, without
 * sending any mail or touching the state file. Safe to run at any time.
 *
 * @returns {Object} Summary of the run
 */
function previewRecoveryEmailCompliance() {
  return runRecoveryNotification_({ dryRun: true });
}

/**
 * @param {Object} options - { dryRun: boolean }
 * @returns {Object} Summary of the run
 */
function runRecoveryNotification_(options) {
  const dryRun = !!(options && options.dryRun);

  if (!PROFILE_.RUN_RECOVERY_EMAIL_NOTIFICATIONS) {
    Logger.info('Recovery-email notifications disabled for this tenant profile', {
      profile: TENANT_PROFILE
    });
    return { skipped: true };
  }

  clearCache(); // Ensure fresh CAPWATCH data
  const start = new Date();
  const todayIso = rcIsoDate_(start);
  Logger.info(dryRun
    ? 'Starting recovery-email compliance PREVIEW'
    : 'Starting recovery-email compliance notification');

  // Duty positions are not needed on the members being *evaluated* (recipients
  // are resolved separately), and cadet-lite members are excluded by the same
  // filter that keeps them from getting accounts — a member with no account has
  // no password to reset.
  const members = getMembers(CONFIG.MEMBER_TYPES.ACTIVE, false, false);
  const capids = Object.keys(members);

  if (!capids.length) {
    // An empty roster would otherwise read as "everyone is compliant" and wipe
    // the suppression state, so the next real run re-mails the whole wing.
    Logger.error('No members returned — aborting without touching state');
    return { aborted: true, reason: 'no members' };
  }

  const evaluation = rcEvaluateMembers_(members);
  const prior = rcLoadState_();

  Logger.info('Recovery-email evaluation complete', {
    members: capids.length,
    evaluated: evaluation.evaluated,
    compliant: evaluation.compliant,
    flagged: Object.keys(evaluation.flagged).length,
    skippedNoCapwatchContact: evaluation.skipped
  });

  // Build the state this run will leave behind.
  //
  // Compliant and departed members are simply absent from `flagged`, so they
  // drop out of state here — that is what lets a member who regresses later be
  // reported immediately instead of inside a stale suppression window.
  const nextState = {};
  const toNotify = {};   // orgid -> [flagged member, ...]
  let suppressed = 0;

  Object.keys(evaluation.flagged).forEach(function (capid) {
    const member = evaluation.flagged[capid];
    const record = prior[capid];

    if (record && rcIsSuppressed_(record.lastNotified, todayIso)) {
      // Still inside the three-month window: carry the record forward untouched
      // so the window continues to run from the date they were actually told.
      suppressed++;
      nextState[capid] = record;
      return;
    }

    nextState[capid] = { lastNotified: todayIso, reasons: member.reasons };
    if (!toNotify[member.orgid]) toNotify[member.orgid] = [];
    toNotify[member.orgid].push(member);
  });

  const recipientsByOrg = rcBuildRecipientDirectory_();
  const summary = {
    dryRun: dryRun,
    members: capids.length,
    evaluated: evaluation.evaluated,
    compliant: evaluation.compliant,
    skipped: evaluation.skipped,
    flagged: Object.keys(evaluation.flagged).length,
    suppressed: suppressed,
    reported: 0,
    sent: 0,
    failedOrgs: [],
    noRecipientOrgs: [],
    startTime: start.toISOString()
  };

  Object.keys(toNotify).forEach(function (orgid) {
    const flagged = toNotify[orgid];
    const recipients = recipientsByOrg[String(orgid)] || [];

    if (!recipients.length) {
      // Do NOT advance state for these members. Leaving them pending means the
      // digest fires once a commander or personnel officer is assigned, rather
      // than being lost to a vacancy — an undeliverable compliance signal should
      // stay visible.
      Logger.warn('No reachable command staff for unit with flagged members — leaving pending', {
        orgid: orgid,
        orgName: flagged[0].orgName,
        memberCount: flagged.length
      });
      summary.noRecipientOrgs.push({
        orgid: orgid,
        orgName: flagged[0].orgName,
        members: flagged.length
      });
      rcRevertOrgState_(nextState, flagged, prior);
      return;
    }

    if (dryRun) {
      // Same reduction the real send performs, so the preview shows the actual
      // addressee and Cc rather than the raw (repeating) duty list.
      const selected = rcSelectAddressees_(recipients);
      Logger.info('[PREVIEW] Would send recovery-email digest', {
        orgid: orgid,
        orgName: flagged[0].orgName,
        to: selected.addressee.email,
        cc: selected.cc,
        members: flagged.map(f => f.capid + ' [' + f.reasons + ']')
      });
      summary.sent++;
      summary.reported += flagged.length;
      return;
    }

    const ok = rcSendDigest_(orgid, recipients, flagged);
    if (ok) {
      summary.sent++;
      summary.reported += flagged.length;
    } else {
      // Per-unit commit: a failed digest leaves its members at their prior record
      // so the next run retries them, while delivered units still advance.
      summary.failedOrgs.push({ orgid: orgid, orgName: flagged[0].orgName });
      rcRevertOrgState_(nextState, flagged, prior);
    }

    Utilities.sleep(RECOVERY_NOTIFY_CONFIG.EMAIL_DELAY_MS);
  });

  if (!dryRun) {
    rcSaveState_(nextState);
  } else {
    Logger.info('[PREVIEW] State file not written');
  }

  summary.endTime = new Date().toISOString();
  summary.duration = new Date() - start;
  Logger.info(dryRun
    ? 'Recovery-email compliance preview complete'
    : 'Recovery-email compliance notification complete', summary);

  if (!dryRun && (summary.failedOrgs.length || summary.noRecipientOrgs.length)) {
    rcSendSummaryEmail_(summary);
  }

  return summary;
}

/**
 * Restores the prior record for a unit's members, so a run that could not
 * deliver their digest retries them next time.
 *
 * Restoring the whole prior record — not just deleting — matters: a member who
 * was previously reported keeps their original lastNotified, so the suppression
 * window continues to run from when they were actually told.
 *
 * @param {Object} nextState - State being built for this run
 * @param {Array<Object>} flagged - Flagged members for one unit
 * @param {Object} prior - State as loaded at the start of this run
 * @returns {void}
 */
function rcRevertOrgState_(nextState, flagged, prior) {
  flagged.forEach(function (member) {
    const record = prior[member.capid];
    if (record === undefined) {
      delete nextState[member.capid];
    } else {
      nextState[member.capid] = record;
    }
  });
}

// ============================================================================
// EVALUATION
// ============================================================================

/**
 * Applies the two compliance conditions to every member.
 *
 * Members merged from the ManualMembers sheet are skipped. They never pass
 * through addContactInfo() (loadManualMembers() merges them afterwards), so they
 * carry no `recoveryEmail` at all — evaluated naively, every one of them would be
 * reported as having neither a CAP primary nor a recovery address, which is an
 * artefact of where they came from rather than anything a commander can fix.
 * `typeof recoveryEmail === 'string'` is precisely "went through the CAPWATCH
 * contact derivation", because that derivation assigns it to every member it
 * sees, including '' when nothing qualifies.
 *
 * @param {Object} members - CAPID -> member object from getMembers()
 * @returns {Object} { flagged, evaluated, compliant, skipped }
 */
function rcEvaluateMembers_(members) {
  const result = { flagged: {}, evaluated: 0, compliant: 0, skipped: 0 };

  Object.keys(members).forEach(function (capid) {
    const m = members[capid];

    if (typeof m.recoveryEmail !== 'string') {
      result.skipped++;
      return;
    }

    result.evaluated++;

    const flagPrimary = !rcIsTenantDomainEmail_(m.primaryEmailValue);
    const flagRecovery = !m.recoveryEmail;

    if (!flagPrimary && !flagRecovery) {
      result.compliant++;
      return;
    }

    result.flagged[capid] = {
      capid: String(capid),
      name: [m.rank, m.firstName, m.lastName].filter(String).join(' '),
      type: m.type || '',
      orgid: String(m.orgid),
      orgName: m.orgName || '',
      charter: m.charter || '',
      flagPrimary: flagPrimary,
      flagRecovery: flagRecovery,
      reasons: flagPrimary && flagRecovery ? 'BOTH' : (flagPrimary ? 'PRIMARY' : 'RECOVERY')
    };
  });

  return result;
}

/**
 * True when `email` is on one of this tenant's own domains.
 *
 * Mirrors the domain normalisation in UpdateMembers.gs firstPersonalEmail_():
 * case-insensitive, leading '@' optional, so both '@cawg.cap.gov' and
 * 'cawg.cap.gov' forms of the configured domain match. An absent address is not
 * a tenant address, which is what makes "no PRIMARY at all" report as a missing
 * CAP primary rather than passing silently.
 *
 * @param {string} email - Address to test
 * @returns {boolean} True if the address is on CONFIG.DOMAIN / SECONDARY_EMAIL_DOMAIN
 */
function rcIsTenantDomainEmail_(email) {
  if (!email) return false;

  const address = String(email);
  const at = address.lastIndexOf('@');
  if (at < 0) return false;

  const domain = address.slice(at + 1).toLowerCase();
  const tenantDomains = [CONFIG.DOMAIN, CONFIG.SECONDARY_EMAIL_DOMAIN]
    .map(d => String(d || '').trim().toLowerCase().replace(/^@/, ''))
    .filter(Boolean);

  return tenantDomains.indexOf(domain) > -1;
}

/**
 * True while a member is still inside their suppression window.
 *
 * Counted in calendar months rather than days so the digest lands on the same
 * day-of-month it did three months earlier, which is what a monthly cadence
 * implies. Date rollover is handled by the Date constructor.
 *
 * @param {string} lastNotifiedIso - Date last reported, 'yyyy-MM-dd'
 * @param {string} todayIso - This run's date, 'yyyy-MM-dd'
 * @returns {boolean} True if the member should be skipped this run
 */
function rcIsSuppressed_(lastNotifiedIso, todayIso) {
  const last = rcParseIsoDate_(lastNotifiedIso);
  const today = rcParseIsoDate_(todayIso);
  if (!last || !today) return false;

  const eligibleAgain = new Date(
    last.getFullYear(),
    last.getMonth() + RECOVERY_NOTIFY_CONFIG.SUPPRESSION_MONTHS,
    last.getDate()
  );

  return today < eligibleAgain;
}

// ============================================================================
// RECIPIENTS
// ============================================================================

/**
 * Builds ORGID -> [{ capid, firstName, lastName, rank, role, email }] for every
 * unit, covering the commander plus the duty positions in RECIPIENT_DUTY_TITLES.
 *
 * Commanders.txt carries its own name columns, so commanders are read straight
 * from it. Duty holders carry only a CAPID, so their names come from a map built
 * over raw Member.txt. That map is deliberately built from the *raw* extract
 * rather than getMembers(): on the cadets tenant getMembers() returns cadets
 * only, but a cadet unit's commander and personnel officer are seniors — present
 * in the extract (they are assigned to the cadet ORGID) yet filtered out of the
 * member set. Reading raw is what makes cross-tenant recipients resolvable.
 *
 * Note Commanders.txt is nationwide, so this is only ever indexed by ORGIDs
 * drawn from this tenant's own members.
 *
 * @returns {Object} Map of ORGID to an array of recipients
 */
function rcBuildRecipientDirectory_() {
  const nameByCapid = {};
  parseFile('Member').forEach(function (row) {
    const capid = String(row[0] || '').trim();
    if (!capid) return;
    nameByCapid[capid] = {
      lastName: String(row[2] || '').trim(),
      firstName: String(row[3] || '').trim(),
      rank: String(row[14] || '').trim()
    };
  });

  const emailMap = createEmailMap();
  const directoryMap = rcBuildCommandDirectoryMap_();
  const byOrg = {};

  const add = function (orgid, capid, role, nameInfo) {
    if (!orgid || !capid) return;

    const info = nameInfo || nameByCapid[capid] || {};
    const email = rcResolveRecipientEmail_(info, capid, emailMap, directoryMap);
    if (!email) {
      Logger.warn('Command staff member has no resolvable email — skipping', {
        capsn: capid, orgid: orgid, role: role
      });
      return;
    }

    if (!byOrg[orgid]) byOrg[orgid] = [];
    byOrg[orgid].push({
      capid: capid,
      firstName: info.firstName || '',
      lastName: info.lastName || '',
      rank: info.rank || '',
      role: role,
      email: email
    });
  };

  // Commanders.txt: ORGID=0, CAPID=4, NameLast=8, NameFirst=9, Rank=12
  parseFile('Commanders').forEach(function (row) {
    add(String(row[0] || '').trim(), String(row[4] || '').trim(), 'Commander', {
      lastName: String(row[8] || '').trim(),
      firstName: String(row[9] || '').trim(),
      rank: String(row[12] || '').trim()
    });
  });

  // DutyPosition.txt: CAPID=0, Duty=1, Asst=4, ORGID=7. The ORGID is the org the
  // duty is HELD at, which is what "direct command" means here — a squadron
  // member holding a wing duty is not wing staff for this purpose.
  const titles = RECOVERY_NOTIFY_CONFIG.RECIPIENT_DUTY_TITLES;
  parseFile('DutyPosition').forEach(function (row) {
    const title = String(row[1] || '').trim();
    if (titles.indexOf(title) === -1) return;
    const role = String(row[4] || '').trim() === '1' ? title + ' (Assistant)' : title;
    add(String(row[7] || '').trim(), String(row[0] || '').trim(), role);
  });

  Logger.info('Recipient directory built', { orgs: Object.keys(byOrg).length });
  return byOrg;
}

/**
 * Builds CAPID -> real Workspace address from this tenant's own directory, for
 * use as the authoritative recipient address.
 *
 * Only built when command staff actually live on THIS tenant, i.e. the command
 * domain is this tenant's own. On the cadets tenant it is not: command staff are
 * seniors whose accounts are in the *other* tenant's directory, which this
 * script cannot read — and a local hit there would be some cadet-domain account
 * rather than the mailbox the commander reads. So the cadets tenant deliberately
 * skips this and derives instead.
 *
 * Why it matters where it does apply: rcDeriveCommandEmail_ reproduces the
 * DEFAULT account name (addOrUpdateUser builds the same `first.last` string), so
 * it is right for most people and wrong for exactly those whose account is not
 * the default — a `.2` duplicate, a manual creation, a rename. Those are
 * invisible to derivation and would be mailed into the void.
 *
 * Failure is non-fatal: a directory read that throws (scope/permission) falls
 * back to derivation rather than aborting a run that is otherwise fine.
 *
 * @returns {Object} CAPID -> primary email, or {} when not applicable
 */
function rcBuildCommandDirectoryMap_() {
  const commandDomain = String(CONFIG.COMMAND_EMAIL_DOMAIN || '').trim().toLowerCase().replace(/^@/, '');
  const ownDomain = String(CONFIG.EMAIL_DOMAIN || '').trim().toLowerCase().replace(/^@/, '');

  if (!commandDomain || commandDomain !== ownDomain) {
    Logger.info('Command staff are not on this tenant — deriving their addresses', {
      commandDomain: commandDomain,
      tenantDomain: ownDomain
    });
    return {};
  }

  try {
    const map = {};
    getActiveUsers().forEach(function (u) {
      if (u && u.capid && u.email) map[String(u.capid)] = String(u.email);
    });
    Logger.info('Command directory map built from this tenant', { accounts: Object.keys(map).length });
    return map;
  } catch (e) {
    Logger.warn('Could not read the directory — falling back to derived addresses', {
      errorMessage: e.message
    });
    return {};
  }
}

/**
 * Resolves the address to reach a command-staff member at, in order:
 *   1. their real Workspace account (authoritative, when readable — see
 *      rcBuildCommandDirectoryMap_)
 *   2. the derived CAP account first.last@<command domain>
 *   3. their CAPWATCH PRIMARY, only when a name is unusable
 *
 * CAPWATCH primary is last on purpose rather than being preferred: a wrong or
 * personal CAPWATCH primary is the very thing this module exists to report, so
 * it is a genuine last resort.
 *
 * @param {Object} info - { firstName, lastName }
 * @param {string} capid - Recipient CAPID
 * @param {Object} emailMap - CAPID -> CAPWATCH primary address
 * @param {Object} [directoryMap] - CAPID -> real Workspace address
 * @returns {string|null} Address, or null if no route yields one
 */
function rcResolveRecipientEmail_(info, capid, emailMap, directoryMap) {
  return (directoryMap && directoryMap[capid]) ||
    rcDeriveCommandEmail_(info) ||
    emailMap[capid] ||
    null;
}

/**
 * Builds first.last@<command domain>, matching the account-naming convention
 * assignManagerEmails() in UpdateMembers.gs falls back to.
 *
 * CONFIG.COMMAND_EMAIL_DOMAIN is the SENIOR domain on a cadet tenant — see the
 * module header. The address is derived, never verified.
 *
 * @param {Object} info - { firstName, lastName }
 * @returns {string|null} Derived address, or null if it cannot be built
 */
function rcDeriveCommandEmail_(info) {
  const first = String((info && info.firstName) || '').toLowerCase().replace(/\s+/g, '');
  const last = String((info && info.lastName) || '').toLowerCase().replace(/\s+/g, '');
  if (!first || !last) return null;

  const domain = String(CONFIG.COMMAND_EMAIL_DOMAIN || '').trim().replace(/^@/, '');
  if (!domain) return null;

  return first + '.' + last + '@' + domain;
}

// ============================================================================
// STATE
// ============================================================================

/**
 * Loads the recorded state: CAPID -> { lastNotified: 'yyyy-MM-dd', reasons }.
 *
 * @returns {Object} Recorded members map, or {} if this tenant has no usable state
 */
function rcLoadState_() {
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files = folder.getFilesByName(RECOVERY_NOTIFY_CONFIG.STATE_FILE_NAME);

  if (!files.hasNext()) {
    Logger.info('No recovery-compliance state file yet — every flagged member is reportable');
    return {};
  }

  const content = files.next().getBlob().getDataAsString();
  if (!content) return {};

  let parsed;
  try {
    parsed = JSON.parse(content);
  } catch (e) {
    // Refusing to guess: returning {} here would silently drop every suppression
    // window and re-mail the whole wing. Fail instead.
    Logger.error('Recovery-compliance state file is corrupt — refusing to run', {
      errorMessage: e.message,
      fileName: RECOVERY_NOTIFY_CONFIG.STATE_FILE_NAME
    });
    throw new Error(
      'Cannot parse ' + RECOVERY_NOTIFY_CONFIG.STATE_FILE_NAME + '. Fix or delete it ' +
      '(deleting re-reports every non-compliant member) before running again.'
    );
  }

  if (!parsed || parsed.version !== RECOVERY_NOTIFY_CONFIG.STATE_VERSION) {
    // At worst this costs one extra round of digests; the alternative is reading
    // an unknown shape and suppressing people we never actually told.
    Logger.warn('Recovery-compliance state version not recognised — starting fresh', {
      found: parsed ? parsed.version : null,
      expected: RECOVERY_NOTIFY_CONFIG.STATE_VERSION
    });
    return {};
  }

  return parsed.members || {};
}

/**
 * Writes the members map, creating the file when absent.
 *
 * @param {Object} members - CAPID to { lastNotified, reasons }
 * @returns {void}
 */
function rcSaveState_(members) {
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files = folder.getFilesByName(RECOVERY_NOTIFY_CONFIG.STATE_FILE_NAME);
  const content = JSON.stringify({
    version: RECOVERY_NOTIFY_CONFIG.STATE_VERSION,
    written: rcIsoDate_(new Date()),
    members: members
  });

  if (files.hasNext()) {
    files.next().setContent(content);
  } else {
    folder.createFile(RECOVERY_NOTIFY_CONFIG.STATE_FILE_NAME, content);
  }

  Logger.info('Recovery-compliance state saved', {
    memberCount: Object.keys(members).length,
    fileName: RECOVERY_NOTIFY_CONFIG.STATE_FILE_NAME
  });
}

/**
 * Discards the recorded state. The next run reports every non-compliant member
 * again, ignoring any suppression window they were inside.
 *
 * @returns {void}
 */
function rcResetState_() {
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files = folder.getFilesByName(RECOVERY_NOTIFY_CONFIG.STATE_FILE_NAME);
  let removed = 0;
  while (files.hasNext()) {
    files.next().setTrashed(true);
    removed++;
  }
  Logger.info('Recovery-compliance state reset — next run reports everyone again', {
    filesTrashed: removed
  });
}

/**
 * Formats a Date as 'yyyy-MM-dd' in the script's timezone. Suppression is a
 * calendar-day question, so a timestamp would imply precision it does not have.
 *
 * @param {Date} date - Date to format
 * @returns {string} 'yyyy-MM-dd'
 */
function rcIsoDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * @param {string} iso - 'yyyy-MM-dd'
 * @returns {Date|null} Parsed date, or null if unparseable
 */
function rcParseIsoDate_(iso) {
  const parts = String(iso || '').split('-');
  if (parts.length !== 3) return null;
  const year = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10);
  const day = parseInt(parts[2], 10);
  if (isNaN(year) || isNaN(month) || isNaN(day)) return null;
  return new Date(year, month - 1, day);
}

// ============================================================================
// EMAIL
// ============================================================================

/**
 * Reduces a unit's raw command-staff list to who is actually mailed: one
 * addressee and a deduplicated Cc list.
 *
 * Shared by the real send AND the preview, deliberately. One person routinely
 * holds several of these duties (commander *and* personnel officer, or an
 * assistant slot alongside a primary one), so the raw list repeats them — a
 * preview that printed the raw list would show four copies of one name and
 * would not be showing what the run actually does. The preview is the surface
 * an operator checks recipients on, so it has to run the same reduction.
 *
 * @param {Array<Object>} recipients - Command staff for one unit, may repeat
 * @returns {Object} { addressee, cc } — cc is an array of addresses
 */
function rcSelectAddressees_(recipients) {
  const seen = {};
  const deduped = [];
  recipients.forEach(function (r) {
    const key = String(r.email).toLowerCase();
    if (seen[key]) return;
    seen[key] = true;
    deduped.push(r);
  });

  // Address the commander when there is one; otherwise whoever is left is the
  // right person to be asking.
  let addressee = null;
  for (let i = 0; i < deduped.length; i++) {
    if (deduped[i].role === 'Commander') { addressee = deduped[i]; break; }
  }
  if (!addressee) addressee = deduped[0];

  const cc = deduped
    .filter(r => String(r.email).toLowerCase() !== String(addressee.email).toLowerCase())
    .map(r => r.email);

  return { addressee: addressee, cc: cc };
}

/**
 * Sends one unit's command staff a digest of the flagged members under their
 * direct command. The commander is the addressee; everyone else is copied.
 *
 * @param {string} orgid - Unit ORGID
 * @param {Array<Object>} recipients - Command staff for this unit
 * @param {Array<Object>} flagged - Flagged members for this unit
 * @returns {boolean} True if sent
 */
function rcSendDigest_(orgid, recipients, flagged) {
  const selected = rcSelectAddressees_(recipients);
  const addressee = selected.addressee;
  const cc = selected.cc;

  const subject = RECOVERY_NOTIFY_CONFIG.SUBJECT + ' — ' + flagged[0].charter;

  try {
    const options = {
      htmlBody: rcBuildDigestHtml_(addressee, flagged),
      from: AUTOMATION_SENDER_EMAIL,
      name: SENDER_NAME,
      replyTo: ITSUPPORT_EMAIL
    };
    if (cc.length) options.cc = cc.join(',');

    executeWithRetry(() =>
      GmailApp.sendEmail(
        addressee.email,
        subject,
        'See the HTML version of this message.',
        options
      )
    );

    Logger.info('Recovery-email digest sent', {
      to: addressee.email,
      cc: cc.length,
      orgid: orgid,
      orgName: flagged[0].orgName,
      members: flagged.length
    });
    return true;

  } catch (e) {
    Logger.error('Failed to send recovery-email digest', {
      to: addressee.email,
      orgid: orgid,
      orgName: flagged[0].orgName,
      errorMessage: e.message,
      errorCode: e.details?.code
    });
    return false;
  }
}

/**
 * @param {Object} addressee - Recipient the digest is addressed to
 * @param {Array<Object>} flagged - Flagged members for this unit
 * @returns {string} HTML body
 */
function rcBuildDigestHtml_(addressee, flagged) {
  const url = RECOVERY_NOTIFY_CONFIG.SELF_SERVICE_URL;

  const rows = flagged.map(function (m) {
    const issues = [];
    if (m.flagPrimary) issues.push('CAP email is not their PRIMARY email');
    if (m.flagRecovery) issues.push('No personal (non-CAP) email on file');
    return '<tr><td>' + rcEscapeHtml_(m.name) + '</td><td>' + rcEscapeHtml_(m.capid) +
      '</td><td>' + rcEscapeHtml_(m.type) + '</td><td>' +
      issues.map(rcEscapeHtml_).join('<br>') + '</td></tr>';
  }).join('');

  return '<html><head><style>' +
    'body { font-family: Arial, sans-serif; color: #202124; }' +
    'h2 { color: #1a73e8; }' +
    'table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }' +
    'th, td { border: 1px solid #ddd; padding: 8px; text-align: left; vertical-align: top; }' +
    'th { background-color: #1a73e8; color: white; }' +
    '.action { background-color: #e8f0fe; padding: 10px; border-left: 4px solid #1a73e8; }' +
    '.footer { font-size: 12px; color: #666; }' +
    '</style></head><body>' +

    '<p>' + rcEscapeHtml_([addressee.rank, addressee.lastName].filter(String).join(' ')) + ',</p>' +

    '<p>The following ' + flagged.length + ' member(s) under your direct command have an ' +
    'email record in eServices that will prevent them from resetting their ' +
    rcEscapeHtml_(CONFIG.ORG_LABEL) + ' account password.</p>' +

    '<p>Password resets are sent to a personal, non-' + rcEscapeHtml_(CONFIG.ORG_LABEL) +
    ' address — a member locked out of their account cannot receive a reset at that same ' +
    'account. Members should have their ' + rcEscapeHtml_(CONFIG.ORG_LABEL) +
    ' address listed as their <strong>primary</strong> email and a personal address listed ' +
    'as their <strong>secondary</strong> email.</p>' +

    '<h2>Members needing correction (' + flagged.length + ')</h2>' +
    '<table><tr><th>Member</th><th>CAPID</th><th>Type</th><th>Issue</th></tr>' +
    rows + '</table>' +

    '<div class="action"><p><strong>What to do:</strong> please contact these members and ask ' +
    'them to correct their email addresses in eServices. If you are certain of a member\'s ' +
    'information, you may make the change on their behalf — as a commander or personnel ' +
    'officer you have the privileges to do so at ' +
    '<a href="' + url + '">' + rcEscapeHtml_(url) + '</a>.</p></div>' +

    '<hr><p class="footer">Automated notification derived from the CAPWATCH extract ' +
    '(MbrContact). eServices is authoritative. To avoid repeat notices, a member reported ' +
    'here will not be reported again for ' + RECOVERY_NOTIFY_CONFIG.SUPPRESSION_MONTHS +
    ' months, even if the issue is not yet corrected. Questions: ' +
    rcEscapeHtml_(ITSUPPORT_EMAIL) + '.</p></body></html>';
}

/**
 * Reports undeliverable digests to IT support. Only sent when something needs a
 * human — a clean run stays quiet.
 *
 * Deliberately does NOT set `from: AUTOMATION_SENDER_EMAIL`. The most likely
 * reason a whole run fails is that the executing account lacks that Send-As
 * alias; if this alarm used the same `from` it would fail the same way and the
 * failure would never be reported. See LSCodeNotify.gs, where this was proven
 * live.
 *
 * @param {Object} summary - Run summary
 * @returns {void}
 */
function rcSendSummaryEmail_(summary) {
  try {
    let html =
      '<html><body style="font-family: Arial, sans-serif;">' +
      '<h2>Recovery-email compliance — items needing attention</h2>' +
      '<p>Sent ' + summary.sent + ' unit digest(s) covering ' + summary.reported +
      ' member(s). Flagged: ' + summary.flagged + ', suppressed (reported within the last ' +
      RECOVERY_NOTIFY_CONFIG.SUPPRESSION_MONTHS + ' months): ' + summary.suppressed + '.</p>';

    if (summary.noRecipientOrgs.length) {
      html +=
        '<h3>No reachable command staff (' + summary.noRecipientOrgs.length + ')</h3>' +
        '<p>These units have flagged members but no commander, personnel officer or deputy ' +
        'with a resolvable email. They remain pending and will notify once one is assigned.</p><ul>' +
        summary.noRecipientOrgs.map(o =>
          '<li>' + rcEscapeHtml_(o.orgName) + ' (ORGID ' + rcEscapeHtml_(o.orgid) +
          ') — ' + o.members + ' member(s)</li>'
        ).join('') + '</ul>';
    }

    if (summary.failedOrgs.length) {
      html +=
        '<h3>Send failed (' + summary.failedOrgs.length + ')</h3>' +
        '<p>These will be retried on the next run.</p><ul>' +
        summary.failedOrgs.map(o =>
          '<li>' + rcEscapeHtml_(o.orgName) + ' (ORGID ' + rcEscapeHtml_(o.orgid) + ')</li>'
        ).join('') + '</ul>';
    }

    html += '</body></html>';

    // No `from:` — see the function doc.
    GmailApp.sendEmail(ITSUPPORT_EMAIL, 'Recovery-email compliance — attention needed',
      'See HTML version', { htmlBody: html, name: SENDER_NAME });
  } catch (e) {
    Logger.error('Failed to send recovery-compliance summary email', { errorMessage: e.message });
  }
}

/**
 * Escapes CAPWATCH-sourced text for inclusion in an HTML email body.
 *
 * @param {string} value - Raw value
 * @returns {string} Escaped value
 */
function rcEscapeHtml_(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ============================================================================
// SETUP
// ============================================================================

/**
 * Installs the monthly time-driven trigger for notifyRecoveryEmailCompliance().
 *
 * MUST be run while signed in as the automation account. A time-driven trigger
 * runs as whoever creates it, and only the automation account owns the
 * AUTOMATION_SENDER_EMAIL Send-As alias the digests require (see SENDING IDENTITY
 * in the module header) — created under any other identity, every digest fails
 * with "Invalid argument". Confirm the owner in the Triggers panel afterward.
 *
 * Idempotent: removes any existing triggers for this handler first, so re-running
 * never stacks duplicates, and leaves triggers for other functions alone.
 *
 * Scheduled for the 1st of the month, after getCapwatch() has refreshed the
 * extract. Unlike LSCodeNotify's weekly trigger, the cadence here does not change
 * what the digest claims — only how often a unit hears from it. The three-month
 * suppression window is independent of the trigger interval.
 *
 * @returns {void}
 */
function installRecoveryComplianceMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'notifyRecoveryEmailCompliance') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('notifyRecoveryEmailCompliance')
    .timeBased()
    .onMonthDay(1)
    .atHour(7)
    .create();

  Logger.info('Installed monthly recovery-compliance trigger', {
    handler: 'notifyRecoveryEmailCompliance',
    schedule: '1st of each month ~07:00 America/Los_Angeles',
    note: 'Confirm in the Triggers panel that the owner is the automation account'
  });
}

// ============================================================================
// TEST
// ============================================================================

/**
 * Renders a digest from fabricated data and sends it to TEST_EMAIL, so the
 * template can be reviewed without waiting for a real run.
 *
 * @returns {void}
 */
function testRecoveryDigestToTestEmail() {
  const addressee = { rank: 'Maj', firstName: 'Pat', lastName: 'Example', email: TEST_EMAIL };
  const flagged = [
    {
      capid: '123456', name: 'Capt Jamie Rivera', type: 'SENIOR', charter: 'PCR-CA-070',
      orgName: 'Example Squadron', flagPrimary: true, flagRecovery: false
    },
    {
      capid: '234567', name: '2d Lt Alex Chen', type: 'SENIOR', charter: 'PCR-CA-070',
      orgName: 'Example Squadron', flagPrimary: false, flagRecovery: true
    },
    {
      capid: '345678', name: 'C/CMSgt Sam Fitzgerald', type: 'CADET', charter: 'PCR-CA-070',
      orgName: 'Example Squadron', flagPrimary: true, flagRecovery: true
    }
  ];

  GmailApp.sendEmail(TEST_EMAIL, 'TEST — recovery-email compliance digest preview',
    'See HTML version', {
      htmlBody: rcBuildDigestHtml_(addressee, flagged),
      from: AUTOMATION_SENDER_EMAIL,
      name: SENDER_NAME
    });

  Logger.info('Test recovery-compliance digest sent', { recipient: TEST_EMAIL });
}
