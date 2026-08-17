/***********************************************
 * File: TwoSvSetupGroup.gs
 * Description: Nightly prune of the 2SV setup group — a member leaves it the
 *   moment they enroll in 2-Step Verification, or after the grace window
 *   expires, whichever comes first.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-08-16
 * Changes: 1.0.0 — new module.
 ***********************************************/

/**
 * WHAT THIS IS FOR
 *
 * The 2SV setup group is a holding pen. A member who cannot sign in because
 * enforcement is on, or who is being walked through enrollment by the help desk,
 * gets parked in it: membership exempts them from 2SV enforcement long enough to
 * set it up. admin-webapp/ (the help-desk app) is what puts most of them there,
 * one member at a time.
 *
 * Nothing took them back out. An exemption granted for an afternoon quietly
 * became permanent, and the group — whose entire purpose is to be small and
 * temporary — accumulated members who are the least protected accounts in the
 * wing precisely because they are in it. This closes that loop nightly:
 *
 *   - enrolled in 2SV      -> removed, the exemption did its job
 *   - GRACE_DAYS elapsed   -> removed, the exemption expired
 *
 * whichever lands first.
 *
 * WHY A LEDGER
 *
 * "When was this member added to the group?" is not answerable from the
 * directory: AdminDirectory.Members carries no join timestamp, and the Reports
 * API keeps admin activity for six months but is a different (and much heavier)
 * read. So this module writes down what it sees. Each run stamps a first-seen
 * date on members it has not seen before, and the grace clock runs from there.
 *
 * The consequence, stated plainly because it matters on day one: members ALREADY
 * in the group when this first runs have their clock started at that first run,
 * not at whenever they were really added. They get a full grace window rather
 * than being removed immediately. That is the safe direction to be wrong in —
 * the alternative is stripping an exemption out from under someone mid-setup —
 * and it self-corrects after one window.
 *
 * WHAT IT WILL NOT DO
 *
 *   - Touch any group but the configured one. TENANT_2SV_SETUP_GROUP blank
 *     disables the whole module rather than guessing at an address.
 *   - Act on a nested group, or on any member that is not a USER.
 *   - Remove anyone on a failed 2SV read. An account the directory would not
 *     report on is left to the grace clock alone (which is time-based and needs
 *     no directory read), so a transient API failure cannot manufacture a
 *     removal — nor keep a stale exemption alive forever.
 */

const TWO_SV_SETUP_CONFIG = {
  /**
   * How long an exemption may live without 2SV showing up. Seven days: long
   * enough for a member to find a phone and a quiet moment, short enough that an
   * account without 2SV is not exempt from enforcement for a month.
   */
  GRACE_DAYS: 7,

  // Own state file, in the CAPWATCH data folder alongside the other state.
  // Nothing else writes it.
  LEDGER_FILE_NAME: 'TwoSvSetupGroupLedger.txt',

  // Bumped when the ledger's shape changes. An unrecognised version is REFUSED
  // rather than re-baselined — see twoSvLedgerLoad_.
  LEDGER_VERSION: 1,

  // Members read per Directory page.
  PAGE_SIZE: 200
};

/**
 * The group, from Script Properties. BLANK DISABLES THIS MODULE — see
 * TENANT_2SV_SETUP_GROUP in config.gs.
 *
 * A FUNCTION, not a constant in the object above, and that is not a style
 * choice. Apps Script evaluates every file's top level in editor order, and
 * subfoldered files sort ahead of config.gs — so reading TENANT out here throws
 * `ReferenceError: TENANT is not defined` before config.gs has run. A top-level
 * throw aborts the load of the WHOLE project, taking every unrelated function
 * down with it (seen live on seniors, 2026-08-16). Read tenant config inside a
 * function body, where the whole project is already loaded.
 *
 * NOTE this is a SECOND property naming the same group: admin-webapp/ is its own
 * Apps Script project with its own Script Properties, and reads it as
 * WEBAPP_2SV_SETUP_GROUP. Set both, to the same address.
 *
 * @returns {string} lowercased group address, or '' when unconfigured
 */
function twoSvGroupEmail_() {
  return String(TENANT.TWO_SV_SETUP_GROUP || '').trim().toLowerCase();
}

// ============================================================================
// LEDGER
// ============================================================================

/**
 * Loads the first-seen ledger.
 *
 * A missing file is normal exactly once, and is reported as empty: every current
 * member is then seen for the first time today and gets a full grace window. A
 * file that exists but cannot be read is a different matter and throws — reading
 * it wrong would either strip exemptions early or reset everyone's clock.
 *
 * @returns {{firstSeen: Object}} maps lowercased email -> iso-date first seen
 */
function twoSvLedgerLoad_() {
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files = folder.getFilesByName(TWO_SV_SETUP_CONFIG.LEDGER_FILE_NAME);

  if (!files.hasNext()) {
    Logger.info('No 2SV setup-group ledger yet — this run establishes it', {
      fileName: TWO_SV_SETUP_CONFIG.LEDGER_FILE_NAME
    });
    return { firstSeen: {} };
  }

  const content = files.next().getBlob().getDataAsString();
  if (!content) return { firstSeen: {} };

  let parsed;
  try {
    parsed = JSON.parse(content);
  } catch (e) {
    Logger.error('2SV setup-group ledger is corrupt — refusing to run', {
      errorMessage: e.message,
      fileName: TWO_SV_SETUP_CONFIG.LEDGER_FILE_NAME
    });
    throw new Error(
      'Cannot parse ' + TWO_SV_SETUP_CONFIG.LEDGER_FILE_NAME + '. Fix it, or delete it — ' +
      'deleting restarts every current member\'s grace window from the next run.'
    );
  }

  if (!parsed || parsed.version !== TWO_SV_SETUP_CONFIG.LEDGER_VERSION) {
    throw new Error(
      TWO_SV_SETUP_CONFIG.LEDGER_FILE_NAME + ' has version ' +
      (parsed ? parsed.version : 'none') + ', expected ' + TWO_SV_SETUP_CONFIG.LEDGER_VERSION +
      '. Refusing to guess at its shape — reading it wrong would remove exemptions early.'
    );
  }

  return { firstSeen: parsed.firstSeen || {} };
}

/**
 * Writes the ledger, creating the file when absent.
 *
 * @param {{firstSeen: Object}} ledger
 * @returns {void}
 */
function twoSvLedgerSave_(ledger) {
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files = folder.getFilesByName(TWO_SV_SETUP_CONFIG.LEDGER_FILE_NAME);
  const content = JSON.stringify({
    version: TWO_SV_SETUP_CONFIG.LEDGER_VERSION,
    written: new Date().toISOString(),
    firstSeen: ledger.firstSeen || {}
  });

  if (files.hasNext()) {
    files.next().setContent(content);
  } else {
    folder.createFile(TWO_SV_SETUP_CONFIG.LEDGER_FILE_NAME, content, MimeType.PLAIN_TEXT);
  }
}

// ============================================================================
// THE DECISION
// ============================================================================

/**
 * Decides, for the group's current membership, who leaves and who stays.
 *
 * Pure: no Google calls, no clock of its own. Everything it needs is an
 * argument, which is what makes the policy testable (test/TwoSvSetupGroup.test.js).
 *
 * @param {Array<{email: string, enrolledIn2Sv: (boolean|null)}>} members - current
 *   direct USER members. `enrolledIn2Sv` null means the account's 2SV state could
 *   not be read this run.
 * @param {Object} firstSeen - lowercased email -> iso-date, from the ledger
 * @param {Date} now
 * @param {{graceDays: (number|undefined)}} [options]
 * @returns {{remove: Array<Object>, keep: Array<Object>, firstSeen: Object, added: number}}
 *   `firstSeen` is the ledger to write back: stamped for newcomers, and pruned of
 *   anyone no longer in the group so a re-add starts a fresh window.
 */
function evaluateTwoSvSetupGroup_(members, firstSeen, now, options) {
  const opts = options || {};
  const graceDays = opts.graceDays == null ? TWO_SV_SETUP_CONFIG.GRACE_DAYS : opts.graceDays;
  const today = twoSvDateStamp_(now);

  const nextFirstSeen = {};
  const remove = [];
  const keep = [];
  let added = 0;

  for (let i = 0; i < members.length; i++) {
    const email = String(members[i].email || '').trim().toLowerCase();
    if (!email) continue;

    const seenOn = firstSeen[email] || today;
    if (!firstSeen[email]) added++;
    nextFirstSeen[email] = seenOn;

    const days = twoSvDaysBetween_(seenOn, now);
    const entry = {
      email: email,
      addedOn: seenOn,
      days: days,
      enrolledIn2Sv: members[i].enrolledIn2Sv
    };

    // 2SV first: it is the outcome the group exists to produce, and it holds
    // even on the day the member was added.
    if (members[i].enrolledIn2Sv === true) {
      entry.reason = '2sv-enrolled';
      remove.push(entry);
      continue;
    }

    if (days >= graceDays) {
      // Reached on an unreadable 2SV state too. The exemption is what expires
      // here, and that does not depend on knowing whether they enrolled.
      entry.reason = 'grace-expired';
      remove.push(entry);
      continue;
    }

    entry.reason = members[i].enrolledIn2Sv === null ? 'within-grace-2sv-unknown' : 'within-grace';
    keep.push(entry);
  }

  return { remove: remove, keep: keep, firstSeen: nextFirstSeen, added: added };
}

/**
 * Whole days elapsed from an iso date-stamp to `now`, floored at 0.
 *
 * Date-stamp rather than timestamp granularity, so a member added yesterday
 * afternoon is one day in this morning — the same arithmetic a person does
 * reading the ledger.
 *
 * @param {string} isoDate - 'YYYY-MM-DD'
 * @param {Date} now
 * @returns {number}
 */
function twoSvDaysBetween_(isoDate, now) {
  const then = new Date(String(isoDate) + 'T00:00:00Z');
  if (isNaN(then.getTime())) return 0;
  const today = new Date(twoSvDateStamp_(now) + 'T00:00:00Z');
  const days = Math.floor((today.getTime() - then.getTime()) / 86400000);
  return days > 0 ? days : 0;
}

/**
 * 'YYYY-MM-DD' in UTC, so a ledger written from one timezone reads the same from
 * another.
 *
 * @param {Date} date
 * @returns {string}
 */
function twoSvDateStamp_(date) {
  return date.toISOString().slice(0, 10);
}

// ============================================================================
// DIRECTORY READS
// ============================================================================

/**
 * Lists the group's direct USER members, with each one's 2SV state.
 *
 * Per-member Users.get rather than getActiveUsers(): that helper filters to
 * non-suspended, non-admin accounts carrying a CAPID external ID, and every one
 * of those exclusions describes an account that can legitimately be sitting in
 * this group. Asking about the members we actually have is both accurate and
 * cheap — the group is meant to hold a handful of people.
 *
 * @param {string} groupEmail
 * @returns {Array<{email: string, enrolledIn2Sv: (boolean|null), suspended: boolean}>}
 */
function twoSvReadGroupMembers_(groupEmail) {
  const out = [];
  let pageToken = '';
  let nested = 0;

  do {
    const page = executeWithRetry(() => AdminDirectory.Members.list(groupEmail, {
      maxResults: TWO_SV_SETUP_CONFIG.PAGE_SIZE,
      pageToken: pageToken
    }));

    const members = page.members || [];
    for (let i = 0; i < members.length; i++) {
      const email = String(members[i].email || '').trim().toLowerCase();
      if (!email) continue;

      // A nested group is somebody's deliberate structural choice, and removing
      // it would take an unknown number of people's exemption with it.
      if (members[i].type && members[i].type !== 'USER') {
        nested++;
        continue;
      }

      out.push(twoSvAccountState_(email));
    }

    pageToken = page.nextPageToken || '';
  } while (pageToken);

  if (nested > 0) {
    Logger.warn('2SV setup group contains non-user members — left alone', {
      group: groupEmail,
      count: nested
    });
  }

  return out;
}

/**
 * Reads one account's 2SV state.
 *
 * A failed read (deleted account, external address, API trouble) yields null
 * rather than false: "not enrolled" and "could not tell" are different facts,
 * and only the first should ever end an exemption early.
 *
 * @param {string} email
 * @returns {{email: string, enrolledIn2Sv: (boolean|null), suspended: boolean}}
 */
function twoSvAccountState_(email) {
  try {
    const user = executeWithRetry(() => AdminDirectory.Users.get(email, {
      fields: 'primaryEmail,isEnrolledIn2Sv,suspended'
    }), CONFIG.API_RETRY_ATTEMPTS, [ERROR_CODES.NOT_FOUND]);

    return {
      email: email,
      enrolledIn2Sv: user.isEnrolledIn2Sv === true,
      suspended: user.suspended === true
    };
  } catch (e) {
    Logger.warn('Could not read 2SV state — grace clock still applies', {
      email: email,
      errorMessage: e.message
    });
    return { email: email, enrolledIn2Sv: null, suspended: false };
  }
}

// ============================================================================
// ENTRY POINTS
// ============================================================================

/**
 * Nightly: cross-references the 2SV setup group against 2SV enrollment and
 * removes members who no longer need the exemption.
 *
 * Schedule it AFTER updateAllMembers() (see ADMIN_GUIDE §8) — not because it
 * reads CAPWATCH (it does not), but so a same-morning account exists before its
 * 2SV state is asked about.
 *
 * @param {{dryRun: (boolean|undefined), graceDays: (number|undefined)}} [options]
 * @returns {{group: string, checked: number, removed: number, failed: number,
 *   kept: number, added: number, dryRun: boolean, skipped: (string|undefined)}}
 */
function pruneTwoSvSetupGroup(options) {
  const opts = options || {};
  const dryRun = opts.dryRun === true;
  const group = twoSvGroupEmail_();

  if (!group) {
    Logger.info('TENANT_2SV_SETUP_GROUP is not set — nothing to prune');
    return { group: '', checked: 0, removed: 0, failed: 0, kept: 0, added: 0,
             dryRun: dryRun, skipped: 'not-configured' };
  }

  Logger.info('2SV setup-group prune starting', { group: group, dryRun: dryRun });

  const members = twoSvReadGroupMembers_(group);
  const ledger = twoSvLedgerLoad_();
  const verdict = evaluateTwoSvSetupGroup_(members, ledger.firstSeen, new Date(), opts);

  let removed = 0;
  let failed = 0;

  // Anyone actually out of the group. Their ledger entry goes, so a member who
  // is parked again later gets a fresh window rather than inheriting an expired
  // one and being removed the same night.
  const gone = {};

  for (let i = 0; i < verdict.remove.length; i++) {
    const entry = verdict.remove[i];

    if (dryRun) {
      Logger.info('Would remove from 2SV setup group', {
        email: entry.email, reason: entry.reason, addedOn: entry.addedOn, days: entry.days
      });
      continue;
    }

    try {
      executeWithRetry(() => AdminDirectory.Members.remove(group, entry.email));
      removed++;
      gone[entry.email] = true;
      Logger.info('Removed from 2SV setup group', {
        email: entry.email, reason: entry.reason, addedOn: entry.addedOn, days: entry.days
      });
    } catch (e) {
      failed++;
      // Keep the ledger entry: a removal that failed is still an exemption in
      // force, and the next run must judge it on the original date, not restart
      // its clock.
      Logger.error('Could not remove from 2SV setup group', {
        email: entry.email,
        reason: entry.reason,
        errorMessage: e.message,
        errorCode: e.details?.code
      });
    }
  }

  if (!dryRun) {
    // Written after the removals, so a member whose removal FAILED keeps their
    // original first-seen date and is judged on it again tomorrow — rather than
    // having their clock restarted by the very failure that left them exempt.
    const stillTracked = {};
    Object.keys(verdict.firstSeen).forEach(function (email) {
      if (!gone[email]) stillTracked[email] = verdict.firstSeen[email];
    });
    twoSvLedgerSave_({ firstSeen: stillTracked });
  }

  const summary = {
    group: group,
    checked: members.length,
    removed: dryRun ? 0 : removed,
    failed: failed,
    kept: verdict.keep.length,
    added: verdict.added,
    dryRun: dryRun
  };

  Logger.info('2SV setup-group prune complete', summary);

  // A dry run's whole point is the list, and the editor shows the return value
  // where a reader has to go hunting for log lines.
  if (dryRun) summary.wouldRemove = verdict.remove;

  return summary;
}

/**
 * Same read, same decisions, no removals — run this once before arming the
 * trigger to see who the first night would take out.
 *
 * @returns {Object} the same summary pruneTwoSvSetupGroup() returns
 */
function previewTwoSvSetupGroup() {
  return pruneTwoSvSetupGroup({ dryRun: true });
}
