/**
 * Duplicate Account Guard & Cleanup
 *
 * Version: 1.0.0
 * Date: 2026-07-19
 * Author: Maj Isaac Wilson IV, California Wing
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
 * Picks the authoritative account among several that share a CAPID. Preference,
 * in order: (1) localpart exactly equals the canonical first.last, (2) active over
 * suspended, (3) no numeric collision suffix, (4) newest created. Pure.
 *
 * @param {Array<{email,suspended,created}>} accounts
 * @param {string} canonicalLocal firstname.lastname (no domain), or '' if unknown
 * @returns {Object|null} the chosen account
 */
function chooseAuthoritativeAccount_(accounts, canonicalLocal) {
  if (!accounts || !accounts.length) return null;
  const canon = String(canonicalLocal || '').toLowerCase();

  function score(a) {
    const local = String(a.email || '').split('@')[0].toLowerCase();
    const exact = (canon && local === canon) ? 1 : 0;
    const active = a.suspended ? 0 : 1;
    const noSuffix = /\.\d+$/.test(local) ? 0 : 1;
    return { exact: exact, active: active, noSuffix: noSuffix, created: new Date(a.created || 0).getTime() };
  }

  return accounts.slice().sort(function (x, y) {
    const sx = score(x), sy = score(y);
    return (sy.exact - sx.exact) ||
           (sy.active - sx.active) ||
           (sy.noSuffix - sx.noSuffix) ||
           (sy.created - sx.created);   // newest first
  })[0];
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

  const scan = scanDuplicateAccountsByCapid();

  // Canonical first.last per CAPID from CAPWATCH, to pick the authoritative twin.
  // Best-effort: if member data is unavailable, fall back to active-then-newest.
  let membersByCapid = {};
  try {
    membersByCapid = getMembers();
  } catch (e) {
    Logger.warn('Could not load CAPWATCH members; authoritative pick falls back to active/newest', {
      errorMessage: e.message
    });
  }

  const summary = {
    dryRun: isDry,
    groups: scan.duplicateCapidCount,
    retired: [],        // orphans retired (or would be, in dry run)
    skippedSignedIn: [], // orphans left alone because they have login history
    kept: [],           // authoritative accounts
    errors: []
  };

  scan.groups.forEach(function (g) {
    const member = membersByCapid[String(g.capid)];
    const canonicalLocal = member
      ? (String(member.firstName || '') + '.' + String(member.lastName || ''))
          .toLowerCase().replace(/\s+/g, '')
      : '';

    const authoritative = chooseAuthoritativeAccount_(g.accounts, canonicalLocal);
    if (!authoritative) return;
    summary.kept.push({ capid: g.capid, email: authoritative.email });

    g.accounts.forEach(function (a) {
      if (a.email === authoritative.email) return;   // keep the authoritative one

      if (!a.neverSignedIn) {
        summary.skippedSignedIn.push({ capid: g.capid, email: a.email, lastLogin: a.lastLogin });
        Logger.warn('Orphan has login history — left for manual review', {
          capsn: g.capid, email: a.email, authoritative: authoritative.email
        });
        return;
      }

      try {
        const res = retireOrphanAccount_(a.email, isDry);
        summary.retired.push({
          capid: g.capid, email: a.email, authoritative: authoritative.email,
          externalIdsBefore: res.before, externalIdsAfter: res.after, applied: res.applied
        });
        Logger.info(isDry ? 'WOULD retire orphan' : 'Retired orphan', {
          capsn: g.capid, orphan: a.email, authoritative: authoritative.email
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
