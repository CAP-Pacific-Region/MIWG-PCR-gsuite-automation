/**
 * Account-Compliance Notification (recovery email, 2SV, first sign-in)
 *
 * Version: 1.2.1
 * Date: 2026-07-23
 * Changes: 1.2.1 — the never-signed-in guidance routes members to the support
 *   portal (SUPPORT_TICKET_URL) instead of the IT mailbox, matching the 2SV
 *   block; came out of the first one-unit live test.
 *   1.2.0 — new testRecoveryDigestForOrg(orgid, recipient): renders ONE
 *   unit's real digest (post-suppression, exactly what the next run would send
 *   that unit) and mails it to a test recipient only. Reads state, writes
 *   nothing — nobody lands on the cooldown. No `from` override, so it runs
 *   from any signed-in account. Call from a scratch.gs wrapper to pass args.
 *   1.1.0 — Two new Workspace-sourced conditions join the digest: accounts with
 *   2-Step Verification not enabled, and accounts never signed into 60+ days
 *   after creation. Both read the tenant's own directory (getActiveUsers now
 *   carries isEnrolledIn2Sv/lastLoginTime/creationTime); a directory read that
 *   fails or returns nothing aborts the run without touching state, for the
 *   same reason an empty roster does. Suppression is now PER ISSUE CATEGORY
 *   (EMAIL / TWOSV / LOGIN) so a member already inside the email window is
 *   still reported when a 2SV gap appears — state v2, with v1 records migrated
 *   in place (their lastNotified becomes the EMAIL category's date).
 *   1.0.1: resolve recipients from this tenant's own directory before deriving
 *   the address, and share one addressee/Cc reduction between the send and the
 *   preview — both out of the first live preview, where the raw duty list
 *   repeated people and derivation alone produced five classes of dead address.
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
 * Four independent conditions are reported, and a member can trip any of them:
 *
 *   flagPrimary  - there is no tenant address in the PRIMARY slot. That covers a
 *                  personal address sitting in PRIMARY *and* no PRIMARY at all.
 *   flagRecovery - there is no usable personal recovery address anywhere.
 *   flagTwoSv    - their Workspace account is in use (has been signed into) but
 *                  2-Step Verification is not enabled on it.
 *   flagLogin    - their Workspace account was created FIRST_LOGIN_GRACE_DAYS or
 *                  more ago and has never been signed into at all.
 *
 * The first two come from CAPWATCH; the last two from this tenant's own
 * directory, joined by CAPID. A member with no resolvable account is simply not
 * evaluated on the account conditions — there is nothing to enroll or sign into.
 * A never-signed-into account is deliberately NOT also flagged for 2SV: you
 * cannot enroll an account you have never entered, and one row saying "has never
 * signed in" already implies the rest.
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
 * gets filtered to trash. State records, PER ISSUE CATEGORY, the date each
 * member was last reported for it, and that category is skipped until
 * SUPPRESSION_MONTHS have passed. The categories are EMAIL (flagPrimary and/or
 * flagRecovery — they share one window, as they did before categories existed),
 * TWOSV and LOGIN. Per-category matters: a member told about their email record
 * in June whose 2SV lapse is noticed in July is reported for 2SV in July, not
 * silently absorbed into the email window. A category that becomes compliant is
 * dropped from state, so if it later regresses it is reported on the next run
 * rather than sitting silently inside a stale suppression window.
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

  // Bumped when the state file's shape changes. Version 1 (member-level date) is
  // migrated in place; anything else unrecognised is re-baselined rather than
  // guessed at — see rcLoadState_.
  STATE_VERSION: 2,

  // Do not report the same member FOR THE SAME ISSUE CATEGORY again until this
  // many months have passed. Categories: EMAIL, TWOSV, LOGIN.
  SUPPRESSION_MONTHS: 3,

  // An account this old that has never been signed into is reported. Younger
  // accounts get the benefit of the doubt — new members take a while to start.
  FIRST_LOGIN_GRACE_DAYS: 60,

  // Duty positions that receive the digest alongside the commander. Assistants
  // are included: the assistant personnel officer is frequently the person who
  // actually does this work.
  RECIPIENT_DUTY_TITLES: [
    'Personnel Officer',
    'Deputy Commander',
    'Deputy Commander for Seniors',
    'Deputy Commander for Cadets'
  ],

  // Broadened from "email records" when the 2SV and first-sign-in conditions
  // joined the digest — the issues are no longer only about eServices email.
  SUBJECT: 'Member account issues needing attention in your unit',

  // Where a commander or personnel officer can make the eServices email
  // correction themselves.
  SELF_SERVICE_URL: 'https://www.capnhq.gov/CAP.PersonnelInfo.Web/',

  // Where a member enables 2-Step Verification on their own account.
  TWOSV_URL: 'https://myaccount.google.com/signinoptions/two-step-verification',

  // Where a member files a support ticket — needed when 2SV enforcement has
  // already locked them out, so they must be exempted before they can sign in
  // and enroll. Region-wide portal, so it is correct for every tenant.
  SUPPORT_TICKET_URL: 'https://support.pcrcap.org',

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

  // The account conditions (2SV, never signed in) read this tenant's own
  // directory. A read that fails or comes back empty would make every member
  // look account-less — dropping every TWOSV/LOGIN record from state and
  // re-mailing those members the moment the directory recovers — so it aborts
  // for exactly the reason an empty roster does.
  let accounts;
  try {
    accounts = getActiveUsers();
  } catch (e) {
    Logger.error('Could not read the directory — aborting without touching state', {
      errorMessage: e.message
    });
    return { aborted: true, reason: 'directory unavailable' };
  }
  if (!accounts.length) {
    Logger.error('Directory returned no accounts — aborting without touching state');
    return { aborted: true, reason: 'directory empty' };
  }

  const accountByCapid = rcBuildAccountMap_(accounts);
  const evaluation = rcEvaluateMembers_(members, accountByCapid, start);
  const prior = rcLoadState_();

  Logger.info('Account-compliance evaluation complete', {
    members: capids.length,
    evaluated: evaluation.evaluated,
    compliant: evaluation.compliant,
    flagged: Object.keys(evaluation.flagged).length,
    skippedNoCapwatchContact: evaluation.skipped
  });

  // Build the state this run will leave behind.
  //
  // Compliant and departed members are simply absent from `flagged`, and a
  // category no longer tripped is absent from `categories`, so both drop out of
  // state here — that is what lets a regression be reported immediately instead
  // of inside a stale suppression window.
  const nextState = {};
  const toNotify = {};   // orgid -> [flagged member, ...]
  let suppressed = 0;

  Object.keys(evaluation.flagged).forEach(function (capid) {
    const member = evaluation.flagged[capid];
    const priorCategories = (prior[capid] && prior[capid].categories) || {};

    // Each issue category runs its own window. A category still inside it
    // carries its recorded date forward untouched, so the window continues to
    // run from the date the unit was actually told about THAT issue; a category
    // outside it (or new) is stamped today and reported.
    const categories = {};
    const reportable = [];
    member.categories.forEach(function (category) {
      if (rcIsSuppressed_(priorCategories[category], todayIso)) {
        categories[category] = priorCategories[category];
      } else {
        categories[category] = todayIso;
        reportable.push(category);
      }
    });

    nextState[capid] = { categories: categories };

    if (!reportable.length) {
      suppressed++;
      return;
    }

    // The digest shows only the categories being reported this run; issues
    // still inside their window stay quiet rather than becoming a monthly nag.
    member.reportCategories = reportable;
    if (!toNotify[member.orgid]) toNotify[member.orgid] = [];
    toNotify[member.orgid].push(member);
  });

  const recipientsByOrg = rcBuildRecipientDirectory_(accounts);
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
        members: flagged.map(f => f.capid + ' [' + f.reportCategories.join('+') + ']')
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
 * was previously reported keeps each category's original date, so those
 * suppression windows continue to run from when the unit was actually told.
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
 * Reduces the directory listing to CAPID -> account, keeping — when a CAPID
 * holds several accounts (the known duplicate-account problem) — the one most
 * recently signed into. That is the account the member actually uses, so it is
 * the one whose 2SV and sign-in facts describe the member rather than an
 * abandoned twin; judging the dead twin would flag people who are fine.
 *
 * @param {Array<Object>} accounts - getActiveUsers() output
 * @returns {Object} CAPID -> { email, isEnrolledIn2Sv, lastLoginTime, creationTime }
 */
function rcBuildAccountMap_(accounts) {
  const byCapid = {};
  accounts.forEach(function (account) {
    if (!account || !account.capid) return;
    const capid = String(account.capid);
    const existing = byCapid[capid];
    if (!existing ||
        rcTimestamp_(account.lastLoginTime) > rcTimestamp_(existing.lastLoginTime)) {
      byCapid[capid] = account;
    }
  });
  return byCapid;
}

/**
 * True when the account has actually been signed into. The directory reports a
 * never-used account's lastLoginTime as the epoch, so "parses to a positive
 * timestamp" is exactly "has been signed into".
 *
 * @param {Object} account - Entry from rcBuildAccountMap_
 * @returns {boolean} True if the account has ever been signed into
 */
function rcHasLoggedIn_(account) {
  return rcTimestamp_(account.lastLoginTime) > 0;
}

/**
 * @param {string} isoTimestamp - Directory timestamp, possibly absent
 * @returns {number} Milliseconds since epoch, or 0 when absent/unparseable
 */
function rcTimestamp_(isoTimestamp) {
  const ms = new Date(String(isoTimestamp || '')).getTime();
  return isNaN(ms) ? 0 : ms;
}

/**
 * Applies the compliance conditions to every member: the two CAPWATCH email
 * conditions, and the two account conditions joined from the directory.
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
 * A member with no account in the map is not evaluated on the account
 * conditions: there is nothing to enroll or sign into. A never-signed-into
 * account is flagged for LOGIN only, never also for 2SV — you cannot enroll an
 * account you have never entered.
 *
 * @param {Object} members - CAPID -> member object from getMembers()
 * @param {Object} accountByCapid - CAPID -> account from rcBuildAccountMap_()
 * @param {Date} now - This run's clock, for the first-sign-in grace period
 * @returns {Object} { flagged, evaluated, compliant, skipped }
 */
function rcEvaluateMembers_(members, accountByCapid, now) {
  const result = { flagged: {}, evaluated: 0, compliant: 0, skipped: 0 };
  const graceMs = RECOVERY_NOTIFY_CONFIG.FIRST_LOGIN_GRACE_DAYS * 24 * 60 * 60 * 1000;

  Object.keys(members).forEach(function (capid) {
    const m = members[capid];

    if (typeof m.recoveryEmail !== 'string') {
      result.skipped++;
      return;
    }

    result.evaluated++;

    const flagPrimary = !rcIsTenantDomainEmail_(m.primaryEmailValue);
    const flagRecovery = !m.recoveryEmail;

    const account = accountByCapid[String(capid)];
    let flagTwoSv = false;
    let flagLogin = false;
    if (account) {
      if (rcHasLoggedIn_(account)) {
        flagTwoSv = !account.isEnrolledIn2Sv;
      } else {
        const created = rcTimestamp_(account.creationTime);
        flagLogin = created > 0 && (now.getTime() - created) >= graceMs;
      }
    }

    if (!flagPrimary && !flagRecovery && !flagTwoSv && !flagLogin) {
      result.compliant++;
      return;
    }

    // The suppression categories: PRIMARY and RECOVERY share the EMAIL window
    // (as they did before categories existed); the account conditions each get
    // their own.
    const categories = [];
    if (flagPrimary || flagRecovery) categories.push('EMAIL');
    if (flagTwoSv) categories.push('TWOSV');
    if (flagLogin) categories.push('LOGIN');

    result.flagged[capid] = {
      capid: String(capid),
      name: [m.rank, m.firstName, m.lastName].filter(String).join(' '),
      type: m.type || '',
      orgid: String(m.orgid),
      orgName: m.orgName || '',
      charter: m.charter || '',
      flagPrimary: flagPrimary,
      flagRecovery: flagRecovery,
      flagTwoSv: flagTwoSv,
      flagLogin: flagLogin,
      accountCreated: account && flagLogin ? String(account.creationTime).slice(0, 10) : '',
      categories: categories,
      reasons: [
        flagPrimary ? 'PRIMARY' : '',
        flagRecovery ? 'RECOVERY' : '',
        flagTwoSv ? '2SV' : '',
        flagLogin ? 'LOGIN' : ''
      ].filter(String).join('+')
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
 * @param {Array<Object>} accounts - getActiveUsers() output, fetched once by the run
 * @returns {Object} Map of ORGID to an array of recipients
 */
function rcBuildRecipientDirectory_(accounts) {
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
  const directoryMap = rcBuildCommandDirectoryMap_(accounts);
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
 * The listing is fetched once by the run (the account conditions need it too)
 * and passed in, so this never issues its own directory read.
 *
 * @param {Array<Object>} accounts - getActiveUsers() output
 * @returns {Object} CAPID -> primary email, or {} when not applicable
 */
function rcBuildCommandDirectoryMap_(accounts) {
  const commandDomain = String(CONFIG.COMMAND_EMAIL_DOMAIN || '').trim().toLowerCase().replace(/^@/, '');
  const ownDomain = String(CONFIG.EMAIL_DOMAIN || '').trim().toLowerCase().replace(/^@/, '');

  if (!commandDomain || commandDomain !== ownDomain) {
    Logger.info('Command staff are not on this tenant — deriving their addresses', {
      commandDomain: commandDomain,
      tenantDomain: ownDomain
    });
    return {};
  }

  const map = {};
  (accounts || []).forEach(function (u) {
    if (u && u.capid && u.email) map[String(u.capid)] = String(u.email);
  });
  Logger.info('Command directory map built from this tenant', { accounts: Object.keys(map).length });
  return map;
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
 * Loads the recorded state: CAPID -> { categories: { EMAIL|TWOSV|LOGIN:
 * 'yyyy-MM-dd' } }, each date being when the member was last reported for that
 * category.
 *
 * Version 1 records (one member-level lastNotified, email conditions only) are
 * migrated in place: their date becomes the EMAIL category's date, so windows
 * already running when the account conditions were added keep running instead
 * of re-mailing every previously-reported member.
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

  if (parsed && parsed.version === 1) {
    // v1 stored one member-level date for what is now the EMAIL category.
    const migrated = {};
    Object.keys(parsed.members || {}).forEach(function (capid) {
      const record = parsed.members[capid];
      if (record && record.lastNotified) {
        migrated[capid] = { categories: { EMAIL: record.lastNotified } };
      }
    });
    Logger.info('Migrated v1 recovery-compliance state to per-category records', {
      memberCount: Object.keys(migrated).length
    });
    return migrated;
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
 * @param {Object} members - CAPID to { categories: { EMAIL|TWOSV|LOGIN: 'yyyy-MM-dd' } }
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
 * Renders one unit's digest. Rows show only the categories being REPORTED this
 * run (member.reportCategories); an issue still inside its suppression window
 * stays out of the email entirely, so the digest never turns into a monthly
 * re-listing of things the commander was already told.
 *
 * The guidance blocks are conditional on the issues actually present, so a
 * digest that is all 2SV does not open with a lecture about eServices email
 * slots.
 *
 * @param {Object} addressee - Recipient the digest is addressed to
 * @param {Array<Object>} flagged - Flagged members for this unit
 * @returns {string} HTML body
 */
function rcBuildDigestHtml_(addressee, flagged) {
  const url = RECOVERY_NOTIFY_CONFIG.SELF_SERVICE_URL;
  const twoSvUrl = RECOVERY_NOTIFY_CONFIG.TWOSV_URL;
  const label = rcEscapeHtml_(CONFIG.ORG_LABEL);

  let anyEmail = false;
  let anyTwoSv = false;
  let anyLogin = false;

  const rows = flagged.map(function (m) {
    // testRecoveryDigestToTestEmail and older callers pass no reportCategories;
    // falling back to `categories` renders every current issue.
    const cats = m.reportCategories || m.categories || [];
    const issues = [];
    if (cats.indexOf('EMAIL') > -1) {
      anyEmail = true;
      if (m.flagPrimary) issues.push('CAP email is not their PRIMARY email');
      if (m.flagRecovery) issues.push('No personal (non-CAP) email on file');
    }
    if (cats.indexOf('TWOSV') > -1) {
      anyTwoSv = true;
      issues.push('2-Step Verification is not enabled on their account');
    }
    if (cats.indexOf('LOGIN') > -1) {
      anyLogin = true;
      issues.push('Has never signed in to their account' +
        (m.accountCreated ? ' (created ' + m.accountCreated + ')' : ''));
    }
    return '<tr><td>' + rcEscapeHtml_(m.name) + '</td><td>' + rcEscapeHtml_(m.capid) +
      '</td><td>' + rcEscapeHtml_(m.type) + '</td><td>' +
      issues.map(rcEscapeHtml_).join('<br>') + '</td></tr>';
  }).join('');

  let guidance = '';

  if (anyEmail) {
    guidance +=
      '<div class="action"><p><strong>Email record in eServices:</strong> password resets are ' +
      'sent to a personal, non-' + label + ' address — a member locked out of their account ' +
      'cannot receive a reset at that same account. Members should have their ' + label +
      ' address listed as their <strong>primary</strong> email and a personal address listed ' +
      'as their <strong>secondary</strong> email. Please contact these members and ask them ' +
      'to correct their email addresses in eServices. If you are certain of a member\'s ' +
      'information, you may make the change on their behalf — as a commander or personnel ' +
      'officer you have the privileges to do so at ' +
      '<a href="' + url + '">' + rcEscapeHtml_(url) + '</a>.</p></div>' +

    '<p></p>';
  }

  if (anyTwoSv) {
    const ticketUrl = RECOVERY_NOTIFY_CONFIG.SUPPORT_TICKET_URL;
    guidance +=
      '<div class="action"><p><strong>2-Step Verification:</strong> 2-Step Verification ' +
      'protects a member\'s ' + label + ' account even if their password is guessed or ' +
      'stolen, and it is expected on every account. Please ask these members to turn it on ' +
      'at <a href="' + twoSvUrl + '">' + rcEscapeHtml_(twoSvUrl) + '</a> while signed in to ' +
      'their ' + label + ' account — it takes about two minutes with the phone they already ' +
      'carry. If a member is already locked out of their account for not having 2-Step ' +
      'Verification enabled, they may need to file a support ticket at ' +
      '<a href="' + ticketUrl + '">' + rcEscapeHtml_(ticketUrl) + '</a> to have their ' +
      'account temporarily exempted so they can sign in and enroll.</p></div>' +

    '<p></p>';
  }

  if (anyLogin) {
    const ticketUrl = RECOVERY_NOTIFY_CONFIG.SUPPORT_TICKET_URL;
    guidance +=
      '<div class="action"><p><strong>Never signed in:</strong> these members\' ' + label +
      ' accounts were created over ' + RECOVERY_NOTIFY_CONFIG.FIRST_LOGIN_GRACE_DAYS +
      ' days ago and have never been used, so they are missing unit and wing ' +
      'communications sent to them. Please ask them to sign in at ' +
      '<a href="https://mail.google.com">mail.google.com</a> with their ' + label +
      ' address. A member who does not know their password can use the "Forgot password" ' +
      'link (the reset goes to the personal address in eServices) or file a support ' +
      'ticket at <a href="' + ticketUrl + '">' + rcEscapeHtml_(ticketUrl) + '</a>.</p></div>' +

    '<p></p>';
  }

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
    'issue with their ' + label + ' account or their email record in eServices that needs ' +
    'attention. Each issue, and what to do about it, is explained below the table.</p>' +

    '<h2>Members needing attention (' + flagged.length + ')</h2>' +
    '<table><tr><th>Member</th><th>CAPID</th><th>Type</th><th>Issue</th></tr>' +
    rows + '</table>' +

    guidance +

    '<hr><p class="footer">Automated notification derived from the CAPWATCH extract and the ' +
    label + ' account directory. eServices is authoritative for email records. To avoid ' +
    'repeat notices, an issue reported here will not be reported again for ' +
    RECOVERY_NOTIFY_CONFIG.SUPPRESSION_MONTHS + ' months, even if it is not yet corrected. ' +
    'Questions: ' + rcEscapeHtml_(ITSUPPORT_EMAIL) + '.</p></body></html>';
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
 * Renders ONE unit's REAL digest — exactly the rows the next run would report
 * for that unit — and mails it to a test recipient, and nobody else.
 *
 * Reads the suppression state, so what you see is what the unit's command
 * staff would get (issues inside their window are excluded, exactly as the
 * real send excludes them) — but WRITES NOTHING: nobody is added to the
 * cooldown, so the real run afterwards still reports everyone shown here.
 *
 * Sends WITHOUT the automation `from` override, deliberately: this is run by
 * hand while reviewing, and it must work from whatever account is signed in
 * rather than requiring the automation account's Send-As alias.
 *
 * The editor's Run dropdown cannot pass arguments — call it from a scratch.gs
 * wrapper, e.g.:
 *   function myUnitDigestTest() { testRecoveryDigestForOrg('1299'); }
 *
 * @param {string} orgid - Unit ORGID to render
 * @param {string} [recipient] - Where to send it; defaults to TEST_EMAIL
 * @returns {Object} { orgid, reportable, sent }
 */
function testRecoveryDigestForOrg(orgid, recipient) {
  const targetOrg = String(orgid || '').trim();
  if (!targetOrg) {
    throw new Error(
      'Pass a unit ORGID — run from a scratch.gs wrapper, e.g. ' +
      "function myUnitDigestTest() { testRecoveryDigestForOrg('1299'); }");
  }
  const to = String(recipient || TEST_EMAIL || '').trim();
  if (!to) throw new Error('No recipient: pass one, or set TENANT_TEST_EMAIL.');

  if (!PROFILE_.RUN_RECOVERY_EMAIL_NOTIFICATIONS) {
    Logger.info('Recovery-email notifications disabled for this tenant profile', {
      profile: TENANT_PROFILE
    });
    return { skipped: true };
  }

  clearCache();
  const start = new Date();
  const todayIso = rcIsoDate_(start);

  const members = getMembers(CONFIG.MEMBER_TYPES.ACTIVE, false, false);
  const accounts = getActiveUsers();
  const evaluation = rcEvaluateMembers_(members, rcBuildAccountMap_(accounts), start);
  const prior = rcLoadState_();

  // Same per-category reduction the real run performs, so the rendered rows are
  // the rows the commander would actually receive — nothing suppressed sneaks
  // back in just because this is a test.
  const flagged = [];
  Object.keys(evaluation.flagged).forEach(function (capid) {
    const member = evaluation.flagged[capid];
    if (String(member.orgid) !== targetOrg) return;
    const priorCategories = (prior[capid] && prior[capid].categories) || {};
    const reportable = member.categories.filter(
      c => !rcIsSuppressed_(priorCategories[c], todayIso));
    if (!reportable.length) return;
    member.reportCategories = reportable;
    flagged.push(member);
  });

  if (!flagged.length) {
    Logger.info('Test digest: nothing reportable for this unit right now', { orgid: targetOrg });
    return { orgid: targetOrg, reportable: 0, sent: 0 };
  }

  // The real addressee, so the greeting renders as the commander would see it —
  // but the mail goes ONLY to the test recipient, with no Cc.
  const recipients = rcBuildRecipientDirectory_(accounts)[targetOrg] || [];
  const addressee = recipients.length
    ? rcSelectAddressees_(recipients).addressee
    : { rank: '', lastName: '(no command staff found)' };

  GmailApp.sendEmail(to,
    'TEST — ' + RECOVERY_NOTIFY_CONFIG.SUBJECT + ' — ' + flagged[0].charter,
    'See the HTML version of this message.',
    { htmlBody: rcBuildDigestHtml_(addressee, flagged), name: SENDER_NAME });

  Logger.info('Test digest sent for one unit — state NOT written', {
    orgid: targetOrg,
    orgName: flagged[0].orgName,
    to: to,
    members: flagged.map(f => f.capid + ' [' + f.reportCategories.join('+') + ']')
  });
  return { orgid: targetOrg, reportable: flagged.length, sent: 1 };
}

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
      orgName: 'Example Squadron', flagPrimary: true, flagRecovery: false,
      categories: ['EMAIL']
    },
    {
      capid: '234567', name: '2d Lt Alex Chen', type: 'SENIOR', charter: 'PCR-CA-070',
      orgName: 'Example Squadron', flagPrimary: false, flagRecovery: true,
      flagTwoSv: true, categories: ['EMAIL', 'TWOSV']
    },
    {
      capid: '345678', name: 'C/CMSgt Sam Fitzgerald', type: 'CADET', charter: 'PCR-CA-070',
      orgName: 'Example Squadron', flagLogin: true, accountCreated: '2026-04-01',
      categories: ['LOGIN']
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
