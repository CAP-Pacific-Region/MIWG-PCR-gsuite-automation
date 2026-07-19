/**
 * Duplicate Account Guard & Cleanup
 *
 * Version: 1.2.0
 * Date: 2026-07-19
 * Author: Maj Isaac Wilson IV, California Wing
 *
 * Changes: 1.2.0 — two fixes found by reviewing the first live scan.
 *   (a) chooseAuthoritativeAccount_ ranks on the lastLogin TIMESTAMP, not a
 *   has-ever-signed-in boolean. One cadets pair has BOTH accounts signed into — one
 *   used days ago, one months ago — and the boolean tied them, so "newest created"
 *   picked the stale account and marked the ACTIVELY USED one for retirement.
 *   (b) suspendOrphanDuplicates no longer re-decides KEEP/retire. It had re-ranked
 *   using a canonical first.last from CAPWATCH that the scan's preview does not use,
 *   so the account an admin reviewed as KEEP could differ from the one cleanup kept.
 *   It now consumes the scan's own decision — preview and action are the same call.
 *   1.1.0 — corrected by real scan data from the cadets tenant.
 *   (a) chooseAuthoritativeAccount_ now ranks LOGIN HISTORY above the canonical
 *   first.last name. In the 28 duplicate groups found there, the account members
 *   actually sign into is overwhelmingly the OLDER, oddly-named one, and the newer
 *   canonically-named twin has never been signed into — so the 1.0.0 ordering would
 *   have retired accounts in active use.
 *   (b) Added buildProvisioningEmailByCapid_, the CAPID map provisioning now uses.
 *   Suspending a dead twin is not enough on its own: provisioning still derives that
 *   twin's address, Users.update SUCCEEDS against it, and the guard (which only fires
 *   on a 404) never runs — so the twin would be unsuspended and re-maintained every
 *   run. The map has to resolve the CAPID to the in-use account for a retirement to
 *   stick.
 *
 * TWO JOBS
 *   1. PREVENTION (called from UpdateMembers.gs addOrUpdateUser, v1.17.0):
 *      findExistingAccountsByCapid_() + chooseAuthoritativeAccount_() let the
 *      create path do a live directory lookup by CAPID before inserting, so a
 *      member whose real account is suspended / tagged under a different externalId
 *      type / created out-of-band (last.first) is UPDATED in place instead of
 *      getting a second account. See the 1.17.0 note in UpdateMembers.gs.
 *
 *   2. CLEANUP (suspendOrphanDuplicates, run by hand): retires the extra accounts
 *      that already exist. Retirement = retype the organization externalId to a
 *      RETIRED marker and suspend. NOTHING is deleted — deletion is permanent on
 *      this edition (no Archived-User licenses; see LICENSE_CONFIG) and a
 *      never-signed-in orphan has no mail worth destroying, but suspend-then-review
 *      is still the safe default. Defaults to a DRY RUN.
 *
 * WHY RETYPE THE externalId, NOT JUST SUSPEND
 *   reactivateRenewedMembers() scans suspended users (getInactiveUsers) and
 *   un-suspends any whose CAPID is active in CAPWATCH — and an orphan shares its
 *   CAPID with the live member, so a plain suspend would be reversed on the next
 *   reactivation run, resurrecting the duplicate. Both getInactiveUsers() and the
 *   provisioning map read the CAPID ONLY from externalIds[type='organization'].
 *   Retyping that entry to {type:'custom', customType:'duplicate_retired_capid'}
 *   makes the orphan invisible to reactivation, to getInactiveUsers, to the map,
 *   AND to this guard's own lookup (provisioningCapidsFromUser_ skips the marker) —
 *   while preserving the CAPID on the account for a human audit trail. Fully
 *   reversible: change the type back to 'organization' and unsuspend.
 *
 * SAFETY / SCOPE
 *   Prevention helpers are read-only except the in-place update the caller already
 *   intended. Cleanup writes only when called with dryRun === false. Uses scopes
 *   already in src/appsscript.json (admin.directory.user[.readonly]).
 */

/** externalId type/customType used to mark a CAPID on a retired duplicate. */
var RETIRED_CAPID_EXTERNALID_TYPE = 'custom';
var RETIRED_CAPID_CUSTOM_TYPE = 'duplicate_retired_capid';

/** A CAPID on these tenants is a 5-7 digit number. */
var DUP_GUARD_CAPID_RE = /^\d{5,7}$/;

/** True for the externalId entry retireOrphanAccount_ writes to park a CAPID. */
function isRetiredCapidExternalId_(id) {
  return !!id &&
    id.type === RETIRED_CAPID_EXTERNALID_TYPE &&
    id.customType === RETIRED_CAPID_CUSTOM_TYPE;
}

/**
 * CAPIDs that mark a Directory User as a LIVE, provisionable account for a member.
 * Reads every externalId value that looks like a CAPID plus a top-level employeeId,
 * but deliberately IGNORES the retired marker so a retired duplicate is not treated
 * as an existing account. Pure. Returns a de-duplicated array (usually length 1).
 *
 * @param {Object} user Directory User (projection:'full')
 * @returns {string[]}
 */
function provisioningCapidsFromUser_(user) {
  const found = {};
  (user.externalIds || []).forEach(function (id) {
    if (isRetiredCapidExternalId_(id)) return;
    const v = String(id && id.value != null ? id.value : '').trim();
    if (DUP_GUARD_CAPID_RE.test(v)) found[v] = true;
  });
  const emp = String(user.employeeId != null ? user.employeeId : '').trim();
  if (DUP_GUARD_CAPID_RE.test(emp)) found[emp] = true;
  return Object.keys(found);
}

/**
 * Live directory lookup for accounts carrying a CAPID, INCLUDING suspended ones and
 * every externalId type — the two things the provisioning map (getActiveUsers) is
 * blind to. Retired-marker-only accounts are excluded (they are already handled).
 *
 * @param {string|number} capid
 * @returns {Array<{email,suspended,archived,created,neverSignedIn,orgUnitPath}>}
 */
function findExistingAccountsByCapid_(capid) {
  const wanted = String(capid).trim();
  const out = [];
  if (!DUP_GUARD_CAPID_RE.test(wanted)) return out;

  let pageToken = null;
  try {
    do {
      const page = AdminDirectory.Users.list({
        customer: 'my_customer',
        query: 'externalId=' + wanted,   // matches any externalId value, suspended included
        projection: 'full',
        maxResults: 200,
        pageToken: pageToken
      });
      (page.users || []).forEach(function (u) {
        // externalId= can be loose; confirm a non-retired CAPID carrier really matches.
        if (provisioningCapidsFromUser_(u).indexOf(wanted) === -1) return;
        out.push({
          email: u.primaryEmail,
          suspended: !!u.suspended,
          archived: !!u.archived,
          created: u.creationTime || null,
          lastLogin: u.lastLoginTime || null,   // chooseAuthoritativeAccount_ ranks on recency
          neverSignedIn: dupGuardNeverLoggedIn_(u.lastLoginTime),
          orgUnitPath: u.orgUnitPath || ''
        });
      });
      pageToken = page.nextPageToken;
    } while (pageToken);
  } catch (e) {
    // A lookup failure must not crash provisioning; worst case the caller inserts
    // (the pre-guard behavior). Log loudly so it is visible.
    Logger.warn('CAPID lookup failed in duplicate-create guard', {
      capsn: wanted,
      errorMessage: e.message
    });
  }
  return out;
}

/**
 * Google returns lastLoginTime as the Unix epoch for accounts that never signed in
 * (rather than omitting it). Treat epoch and missing as "never". Named distinctly
 * from the scanner's copy so this module loads standalone.
 * @returns {boolean}
 */
function dupGuardNeverLoggedIn_(lastLoginTime) {
  if (!lastLoginTime) return true;
  return new Date(lastLoginTime).getTime() <= 0;
}

/** Lowercased email localpart with any trailing ".N" collision suffix removed. */
function dupGuardCanonicalLocalpart_(email) {
  const local = String(email || '').split('@')[0].toLowerCase();
  const parts = local.split('.');
  if (parts.length > 1 && /^\d+$/.test(parts[parts.length - 1])) parts.pop();
  return parts.join('.');
}

/**
 * Picks the authoritative account among several that share a CAPID.
 *
 * PREFERENCE ORDER — LOGIN HISTORY FIRST. The real directory told us plainly that
 * the canonical first.last address is NOT a reliable signal of which account a
 * member actually uses: in the duplicate population found on the cadets tenant, the
 * account people log into is overwhelmingly the OLDER, oddly-named one (a .N
 * collision suffix, or a hyphen the derived address drops), while the newer
 * canonically-named twin has never been signed into once. An earlier version of
 * this function preferred the canonical name and would therefore have retired the
 * account the member actually uses. Login history now outranks everything:
 *
 *   1. MOST RECENT login              (the member is demonstrably using it)
 *   2. active over suspended
 *   3. localpart equals canonical first.last
 *   4. no numeric collision suffix
 *   5. newest created
 *
 * Rule 1 compares the actual lastLogin TIMESTAMP, not a has-ever-signed-in boolean.
 * A boolean is not enough: the cadets tenant has a pair where BOTH accounts have
 * login history — one used days ago, the other months ago — and a boolean tied them,
 * letting "newest created" pick the stale account and mark the actively-used one for
 * retirement. Never-signed-in sorts as 0, so it still loses to any real login.
 *
 * Pure.
 *
 * @param {Array<{email,suspended,created,lastLogin,neverSignedIn}>} accounts
 * @param {string} canonicalLocal firstname.lastname (no domain), or '' if unknown
 * @returns {Object|null} the chosen account
 */
function chooseAuthoritativeAccount_(accounts, canonicalLocal) {
  if (!accounts || !accounts.length) return null;
  const canon = String(canonicalLocal || '').toLowerCase();

  function score(a) {
    const local = String(a.email || '').split('@')[0].toLowerCase();
    // Google returns the Unix epoch for never-signed-in accounts, which lands on 0
    // here anyway; the explicit flag guards a missing/!odd value.
    const loginAt = a.neverSignedIn
      ? 0
      : Math.max(0, new Date(a.lastLogin || 0).getTime());
    return {
      loginAt: loginAt,
      active: a.suspended ? 0 : 1,
      exact: (canon && local === canon) ? 1 : 0,
      noSuffix: /\.\d+$/.test(local) ? 0 : 1,
      created: new Date(a.created || 0).getTime()
    };
  }

  return accounts.slice().sort(function (x, y) {
    const sx = score(x), sy = score(y);
    return (sy.loginAt - sx.loginAt) ||   // most recently used first
           (sy.active - sx.active) ||
           (sy.exact - sx.exact) ||
           (sy.noSuffix - sx.noSuffix) ||
           (sy.created - sx.created);     // newest first
  })[0];
}

/**
 * Builds the CAPID -> primaryEmail map that provisioning uses to decide whether a
 * member already has an account (workspaceEmailByCapid in UpdateMembers.gs).
 *
 * WHY THIS EXISTS RATHER THAN CHANGING getActiveUsers()
 *   getActiveUsers() lists only NON-suspended users and reads the CAPID only from
 *   externalIds[type='organization']. Both blind spots let provisioning conclude
 *   "this member has no account" and insert a duplicate. It cannot simply be
 *   widened: suspendExpiredMembers() and ManageLicenses.gs depend on its
 *   active-only contract. So provisioning gets its own map builder, and
 *   getActiveUsers() is left exactly as it was.
 *
 * This map sees SUSPENDED accounts and every CAPID carrier, and when a CAPID has
 * more than one account it resolves to the one the member actually uses
 * (chooseAuthoritativeAccount_ — login history first). That last part is what makes
 * an existing duplicate stable: provisioning maintains the in-use account and never
 * touches the dead twin, so a retired twin stays retired instead of being
 * resurrected by the next run.
 *
 * @returns {Object} capid (string) -> primaryEmail
 */
function buildProvisioningEmailByCapid_() {
  const byCapid = {};
  let pageToken = null;
  let scanned = 0;

  do {
    const page = AdminDirectory.Users.list({
      customer: 'my_customer',
      maxResults: 500,
      projection: 'full',   // needed for externalIds + creationTime + lastLoginTime
      pageToken: pageToken
    });
    (page.users || []).forEach(function (u) {
      scanned++;
      const info = {
        email: u.primaryEmail,
        suspended: !!u.suspended,
        created: u.creationTime || null,
        lastLogin: u.lastLoginTime || null,   // chooseAuthoritativeAccount_ ranks on recency
        neverSignedIn: dupGuardNeverLoggedIn_(u.lastLoginTime)
      };
      provisioningCapidsFromUser_(u).forEach(function (capid) {
        (byCapid[capid] = byCapid[capid] || []).push(info);
      });
    });
    pageToken = page.nextPageToken;
  } while (pageToken);

  const map = {};
  let multi = 0;
  Object.keys(byCapid).forEach(function (capid) {
    const accounts = byCapid[capid];
    if (accounts.length > 1) multi++;
    // No member name is available at map-build time, so there is no canonical hint;
    // login history / active / newest is what decides.
    const chosen = chooseAuthoritativeAccount_(accounts, '');
    if (chosen) map[capid] = chosen.email;
  });

  Logger.info('Provisioning CAPID map built', {
    scannedUsers: scanned,
    mappedCapids: Object.keys(map).length,
    capidsWithMultipleAccounts: multi
  });
  return map;
}

/**
 * Retires one orphan account: retypes its organization externalId to the RETIRED
 * marker and suspends it, so reactivation / mapping cannot resurrect it. Reversible.
 * Writes only when dryRun === false.
 *
 * @param {string} email
 * @param {boolean} dryRun
 * @returns {{email, before: string[], after: string[], applied: boolean}}
 */
function retireOrphanAccount_(email, dryRun) {
  const u = AdminDirectory.Users.get(email, { projection: 'full' });
  const before = (u.externalIds || []).map(function (id) {
    return (id.type || '') + (id.customType ? '/' + id.customType : '') + '=' + id.value;
  });

  const after = (u.externalIds || []).map(function (id) {
    if (id.type === 'organization') {
      return { type: RETIRED_CAPID_EXTERNALID_TYPE, customType: RETIRED_CAPID_CUSTOM_TYPE, value: id.value };
    }
    return { type: id.type, customType: id.customType, value: id.value };
  });

  if (dryRun === false) {
    executeWithRetry(function () {
      return AdminDirectory.Users.update({ externalIds: after, suspended: true }, email);
    });
  }

  return {
    email: email,
    before: before,
    after: after.map(function (id) {
      return (id.type || '') + (id.customType ? '/' + id.customType : '') + '=' + id.value;
    }),
    applied: dryRun === false
  };
}

/**
 * CLEANUP ENTRYPOINT. Finds every CAPID with >1 account (via the read-only
 * scanner), decides which account is authoritative, and RETIRES the extras that
 * have never signed in. Any orphan that HAS signed in is left for a human (it may
 * hold mail). DEFAULTS TO A DRY RUN — pass false to actually retire.
 *
 *   suspendOrphanDuplicates()        // dry run: report only
 *   suspendOrphanDuplicates(false)   // retire never-signed-in orphans for real
 *
 * @param {boolean} [dryRun=true]
 * @returns {Object} summary
 */
function suspendOrphanDuplicates(dryRun) {
  const isDry = dryRun !== false;   // default to dry run
  Logger.info(isDry ? 'Orphan-duplicate cleanup PREVIEW (dry run)' : 'Orphan-duplicate cleanup — LIVE');

  // The scan is the single source of truth for KEEP vs retire. This function
  // deliberately does NOT re-decide with extra inputs (an earlier version re-ranked
  // using a canonical first.last pulled from CAPWATCH, which the scan's preview does
  // not use — so the account an admin reviewed as KEEP could differ from the one
  // cleanup actually kept). Preview and action must be the same decision.
  const scan = scanDuplicateAccountsByCapid();

  const summary = {
    dryRun: isDry,
    groups: scan.duplicateCapidCount,
    retired: [],        // orphans retired (or would be, in dry run)
    skippedSignedIn: [], // orphans left alone because they have login history
    kept: [],           // authoritative accounts
    errors: []
  };

  scan.groups.forEach(function (g) {
    if (!g.authoritativeEmail) return;
    summary.kept.push({ capid: g.capid, email: g.authoritativeEmail });

    g.accounts.forEach(function (a) {
      if (a.email === g.authoritativeEmail) return;   // keep the authoritative one

      if (!a.neverSignedIn) {
        summary.skippedSignedIn.push({ capid: g.capid, email: a.email, lastLogin: a.lastLogin });
        Logger.warn('Orphan has login history — left for manual review', {
          capsn: g.capid, email: a.email, authoritative: g.authoritativeEmail
        });
        return;
      }

      try {
        const res = retireOrphanAccount_(a.email, isDry);
        summary.retired.push({
          capid: g.capid, email: a.email, authoritative: g.authoritativeEmail,
          externalIdsBefore: res.before, externalIdsAfter: res.after, applied: res.applied
        });
        Logger.info(isDry ? 'WOULD retire orphan' : 'Retired orphan', {
          capsn: g.capid, orphan: a.email, authoritative: g.authoritativeEmail
        });
      } catch (e) {
        summary.errors.push({ capid: g.capid, email: a.email, errorMessage: e.message });
        Logger.error('Failed to retire orphan', {
          capsn: g.capid, email: a.email, errorMessage: e.message
        });
      }
    });
  });

  Logger.info('Orphan-duplicate cleanup complete', {
    dryRun: isDry,
    duplicateGroups: summary.groups,
    orphansRetired: summary.retired.length,
    orphansSkippedSignedIn: summary.skippedSignedIn.length,
    errors: summary.errors.length
  });
  if (isDry) {
    Logger.info('DRY RUN — nothing changed. Re-run suspendOrphanDuplicates(false) to apply.');
  }

  return summary;
}
