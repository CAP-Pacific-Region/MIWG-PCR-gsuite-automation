/**
 * Duplicate Workspace Account Scanner  (READ-ONLY diagnostic)
 *
 * Version: 1.2.0
 * Date: 2026-07-19
 *
 * Changes: 1.2.0 — accounts already retired by suspendOrphanDuplicates (whose only
 *   carrier for a CAPID is the duplicate_retired_capid marker) are EXCLUDED from the
 *   duplicate grouping and counted separately. Without this a completed cleanup would
 *   still report the same group count and look like it had done nothing, and a re-run
 *   would reprocess accounts it had already retired.
 *   1.1.0 — reports which FIELD carries each CAPID
 *   (capidCarriersForUser_) plus a count of accounts invisible to an
 *   `externalId=<capid>` query — the number that decides whether the duplicate-create
 *   guard can actually see this population. Each account is now labelled KEEP /
 *   retire so cleanup can be previewed, using the same login-history-first ranking as
 *   the cleanup itself. Summary is logged as its own entry and the detail in chunks,
 *   because Apps Script truncates a single oversized log message.
 *
 * PURPOSE
 *   Find CAP members who hold MORE THAN ONE Workspace account. The provisioning
 *   path in UpdateMembers.gs decides "update vs. create" by whether an account
 *   already exists at a CAPID's *mapped* email (getActiveUsers -> workspaceEmailByCapid).
 *   When that map misses an existing account — because the account is SUSPENDED
 *   (getActiveUsers filters isSuspended=false) or carries its CAPID somewhere the
 *   map doesn't read — provisioning derives a fresh `first.last` email and CREATES
 *   a second account instead of reusing the first. This scanner enumerates the
 *   damage.
 *
 * SAFETY
 *   Read-only. Calls only AdminDirectory.Users.list with projection:'full'.
 *   Uses the admin.directory.user.readonly scope already present in
 *   src/appsscript.json. Writes NOTHING. Safe to run on any tenant.
 *
 * HOW TO RUN
 *   Paste into the Apps Script editor (or run from the deployed project), select
 *   scanDuplicateAccountsByCapid, Run, and read the Execution log. A single
 *   consolidated report is logged at the end; the full structured result is also
 *   returned for programmatic use.
 *
 * WHAT IT READS PER ACCOUNT
 *   CAPID is collected from EVERY carrier so no legacy account is missed:
 *     - externalIds[] of ANY type whose value looks like a CAPID (the code writes
 *       type:'organization'; the Admin console "Employee ID" field is backed by
 *       this same externalIds array)
 *     - a top-level employeeId, if the account happens to carry one
 *     - organizations[].* CAPID-shaped values, defensively
 *   An account can legitimately match on more than one carrier; it is counted once.
 */

/** A CAPID on these tenants is a 5-7 digit number. */
var DUP_SCAN_CAPID_RE = /^\d{5,7}$/;

/**
 * Carrier label written for a CAPID parked on an account already retired by
 * suspendOrphanDuplicates. An account whose ONLY carrier for a CAPID is this marker
 * is no longer part of that member's duplicate set, so it is excluded from the
 * grouping below — otherwise a completed cleanup would still report the same group
 * count and look like it had done nothing, and a re-run would reprocess accounts it
 * had already retired.
 *
 * MUST match RETIRED_CAPID_EXTERNALID_TYPE / RETIRED_CAPID_CUSTOM_TYPE in
 * DuplicateAccountGuard.gs, in the label format capidCarriersForUser_ produces.
 */
var DUP_SCAN_RETIRED_CARRIER = 'externalId:custom/duplicate_retired_capid';

/**
 * Pulls every CAPID-shaped identifier off a Directory User resource, from all
 * the places a CAPID has historically been stored. Returns a de-duplicated array
 * (usually length 1; length 0 means "no CAPID we can see").
 *
 * @param {Object} user Directory User (projection:'full')
 * @returns {string[]} unique CAPID strings found on this account
 */
function extractCapidsFromUser_(user) {
  const found = {};

  (user.externalIds || []).forEach(function (id) {
    const v = String(id && id.value != null ? id.value : '').trim();
    if (DUP_SCAN_CAPID_RE.test(v)) found[v] = true;
  });

  // Some directories surface an employee id at the top level; harmless if absent.
  const emp = String(user.employeeId != null ? user.employeeId : '').trim();
  if (DUP_SCAN_CAPID_RE.test(emp)) found[emp] = true;

  (user.organizations || []).forEach(function (org) {
    ['costCenter', 'description'].forEach(function (k) {
      const v = String(org && org[k] != null ? org[k] : '').trim();
      if (DUP_SCAN_CAPID_RE.test(v)) found[v] = true;
    });
  });

  return Object.keys(found);
}

/**
 * Names every field on this account that carries the given CAPID, e.g.
 * ["externalId:organization", "employeeId"]. This is the field that decides whether
 * the provisioning map and the duplicate-create guard can SEE an account: the map
 * reads externalIds[type='organization'] and the guard queries `externalId=<capid>`,
 * so an account carrying its CAPID only in a top-level employeeId is invisible to
 * both. Pure.
 *
 * @param {Object} user Directory User (projection:'full')
 * @param {string} capid the CAPID to locate
 * @returns {string[]} carrier labels, in the order they were found
 */
function capidCarriersForUser_(user, capid) {
  const wanted = String(capid).trim();
  const out = [];

  (user.externalIds || []).forEach(function (id) {
    if (String(id && id.value != null ? id.value : '').trim() !== wanted) return;
    out.push('externalId:' + (id.type || '?') + (id.customType ? '/' + id.customType : ''));
  });

  if (String(user.employeeId != null ? user.employeeId : '').trim() === wanted) {
    out.push('employeeId');
  }

  (user.organizations || []).forEach(function (org) {
    ['costCenter', 'description'].forEach(function (k) {
      if (String(org && org[k] != null ? org[k] : '').trim() === wanted) {
        out.push('organizations.' + k);
      }
    });
  });

  return out;
}

/**
 * Splits an email localpart into lowercased name tokens on '.', dropping a
 * trailing numeric collision suffix (e.g. "sam.roe.2" -> ["sam","roe"]).
 * @returns {{tokens: string[], suffix: (string|null)}}
 */
function localpartTokens_(email) {
  const local = String(email || '').split('@')[0].toLowerCase();
  const parts = local.split('.');
  let suffix = null;
  if (parts.length > 1 && /^\d+$/.test(parts[parts.length - 1])) {
    suffix = parts.pop();
  }
  return { tokens: parts, suffix: suffix };
}

/**
 * Classifies the relationship between two account localparts in the same CAPID
 * group: 'reversed' (last.first vs first.last), 'collision' (differ only by a
 * trailing .N), or 'other'.
 */
function classifyLocalpartPair_(emailA, emailB) {
  const a = localpartTokens_(emailA);
  const b = localpartTokens_(emailB);

  const aKey = a.tokens.join('.');
  const bKey = b.tokens.join('.');
  const aRev = a.tokens.slice().reverse().join('.');

  if (aKey === bKey && (a.suffix || b.suffix)) return 'collision';
  if (a.tokens.length === b.tokens.length && a.tokens.length >= 2 && aRev === bKey) return 'reversed';
  return 'other';
}

/**
 * Google returns lastLoginTime as the Unix epoch (1970-01-01T00:00:00.000Z) for
 * accounts that have NEVER signed in, rather than omitting it. Treat epoch (and
 * missing) as "never".
 * @returns {boolean}
 */
function neverLoggedIn_(lastLoginTime) {
  if (!lastLoginTime) return true;
  return new Date(lastLoginTime).getTime() <= 0;
}

/**
 * MAIN — scan the whole directory and report every CAPID with more than one
 * account. Read-only.
 *
 * @returns {Object} { scannedUsers, usersWithCapid, usersWithoutCapid,
 *                     duplicateCapidCount, duplicateAccountCount, groups: [...] }
 */
function scanDuplicateAccountsByCapid() {
  const start = new Date();
  Logger.info('Duplicate-account scan starting (read-only)');

  const byCapid = {};        // capid -> [accountInfo, ...]
  let scannedUsers = 0;
  let usersWithoutCapid = 0;
  let retiredAccounts = 0;   // already retired by suspendOrphanDuplicates
  let pageToken = null;

  do {
    const page = AdminDirectory.Users.list({
      customer: 'my_customer',
      maxResults: 500,
      projection: 'full',   // NOTE: do not use a `fields` mask naming employeeId — it 400s
      pageToken: pageToken
    });

    (page.users || []).forEach(function (u) {
      scannedUsers++;
      const capids = extractCapidsFromUser_(u);
      if (capids.length === 0) {
        usersWithoutCapid++;
        return;
      }
      // A single account can expose the same CAPID via >1 carrier; group once.
      // Carriers are per-CAPID, so the entry is built inside the loop.
      capids.forEach(function (capid) {
        const carriers = capidCarriersForUser_(u, capid);

        // Already retired for this CAPID — not part of the live duplicate set.
        if (carriers.length > 0 && carriers.every(function (c) {
          return c === DUP_SCAN_RETIRED_CARRIER;
        })) {
          retiredAccounts++;
          return;
        }

        (byCapid[capid] = byCapid[capid] || []).push({
          email: u.primaryEmail,
          created: u.creationTime || null,
          lastLogin: u.lastLoginTime || null,
          neverSignedIn: neverLoggedIn_(u.lastLoginTime),
          suspended: !!u.suspended,
          archived: !!u.archived,
          orgUnitPath: u.orgUnitPath || '',
          aliases: (u.aliases || []).length,
          carriers: carriers,
          // The duplicate-create guard finds accounts with an `externalId=<capid>`
          // query. An account carrying its CAPID ONLY in a top-level employeeId is
          // invisible to that query — the guard could not have prevented its
          // duplicate, and cannot keep it stable now. This flag counts them.
          findableByExternalIdQuery: carriers.some(function (c) {
            return c.indexOf('externalId:') === 0;
          })
        });
      });
    });

    pageToken = page.nextPageToken;
  } while (pageToken);

  // Build the duplicate groups.
  const groups = [];
  let duplicateAccountCount = 0;

  Object.keys(byCapid).forEach(function (capid) {
    // A single account matched on two carriers would appear twice; de-dupe by email.
    const seen = {};
    const accounts = byCapid[capid].filter(function (a) {
      if (seen[a.email]) return false;
      seen[a.email] = true;
      return true;
    });
    if (accounts.length < 2) return;

    accounts.sort(function (x, y) {
      return new Date(x.created || 0) - new Date(y.created || 0);
    });

    // Pairwise localpart relationship vs the newest account (usual authoritative one).
    const newest = accounts[accounts.length - 1];
    accounts.forEach(function (a) {
      a.localpartVsNewest = a.email === newest.email
        ? 'newest'
        : classifyLocalpartPair_(a.email, newest.email);
    });

    const anyReversed = accounts.some(function (a) { return a.localpartVsNewest === 'reversed'; });
    const anyCollision = accounts.some(function (a) { return a.localpartVsNewest === 'collision'; });
    const allNeverSignedIn = accounts.every(function (a) { return a.neverSignedIn; });

    // Preview exactly what cleanup would keep vs retire. Login history wins — the
    // canonically-named twin is frequently the DEAD one, so name shape is not a
    // safe signal. See chooseAuthoritativeAccount_ in DuplicateAccountGuard.gs.
    const authoritative = chooseAuthoritativeAccount_(accounts, '');
    accounts.forEach(function (a) {
      a.role = (authoritative && a.email === authoritative.email) ? 'KEEP' : 'retire';
    });

    duplicateAccountCount += accounts.length;
    groups.push({
      capid: capid,
      accountCount: accounts.length,
      shape: anyReversed ? 'reversed' : (anyCollision ? 'collision' : 'other'),
      allNeverSignedIn: allNeverSignedIn,
      authoritativeEmail: authoritative ? authoritative.email : null,
      accounts: accounts
    });
  });

  // Most-accounts first, then reversed/collision shapes surface at the top.
  groups.sort(function (a, b) {
    return (b.accountCount - a.accountCount) ||
           (a.shape > b.shape ? 1 : a.shape < b.shape ? -1 : 0);
  });

  const usersWithCapid = scannedUsers - usersWithoutCapid;

  // How many accounts in the duplicate groups can the guard's `externalId=<capid>`
  // query actually see? Any that cannot are accounts the guard could not have
  // prevented and cannot keep stable — the number that decides whether the fix
  // covers this population.
  let guardBlindAccounts = 0;
  const carrierCounts = {};
  groups.forEach(function (g) {
    g.accounts.forEach(function (a) {
      if (!a.findableByExternalIdQuery) guardBlindAccounts++;
      (a.carriers.length ? a.carriers : ['(none)']).forEach(function (c) {
        carrierCounts[c] = (carrierCounts[c] || 0) + 1;
      });
    });
  });

  // Logged as its OWN entry: Apps Script truncates a single oversized message, and
  // this summary must survive that.
  Logger.info('===== DUPLICATE WORKSPACE ACCOUNT SCAN — SUMMARY =====' +
    '\nScanned users .................... ' + scannedUsers +
    '\n  with a readable CAPID .......... ' + usersWithCapid +
    '\n  with NO readable CAPID ......... ' + usersWithoutCapid +
    '\n  already RETIRED (excluded) ..... ' + retiredAccounts +
    '\nCAPIDs with >1 account ........... ' + groups.length +
    '\nAccounts in those groups ......... ' + duplicateAccountCount +
    '\n  INVISIBLE to externalId query .. ' + guardBlindAccounts +
    '   <-- guard cannot see these' +
    '\nCAPID carriers in those groups ... ' + JSON.stringify(carrierCounts));

  // Detail, chunked so a long report is not truncated away.
  const CHUNK = 8;
  for (let i = 0; i < groups.length; i += CHUNK) {
    const lines = ['===== DUPLICATE DETAIL ' + (i + 1) + '-' +
                   Math.min(i + CHUNK, groups.length) + ' of ' + groups.length + ' ====='];
    groups.slice(i, i + CHUNK).forEach(function (g) {
      lines.push('CAPID ' + g.capid + '  (' + g.accountCount + ' accounts, shape=' + g.shape +
                 (g.allNeverSignedIn ? ', none ever signed in' : '') + ')');
      g.accounts.forEach(function (a) {
        lines.push('    [' + a.role + '] ' + a.email +
          '\n        created=' + (a.created || '?') +
          '  ' + (a.suspended ? 'SUSPENDED' : 'active') +
          (a.archived ? '/ARCHIVED' : '') +
          '  ' + (a.neverSignedIn ? 'never-signed-in' : 'last-login=' + a.lastLogin) +
          '  ou=' + a.orgUnitPath +
          '\n        localpart=' + a.localpartVsNewest +
          '  carriers=' + (a.carriers.join(',') || '(none)') +
          (a.findableByExternalIdQuery ? '' : '  <-- NOT findable by externalId query'));
      });
      lines.push('');
    });
    Logger.info(lines.join('\n'));
  }

  const result = {
    scannedUsers: scannedUsers,
    usersWithCapid: usersWithCapid,
    usersWithoutCapid: usersWithoutCapid,
    duplicateCapidCount: groups.length,
    duplicateAccountCount: duplicateAccountCount,
    retiredAccounts: retiredAccounts,
    guardBlindAccounts: guardBlindAccounts,
    carrierCounts: carrierCounts,
    groups: groups,
    durationMs: new Date() - start
  };

  Logger.info('Duplicate-account scan complete', {
    duplicateCapidCount: result.duplicateCapidCount,
    guardBlindAccounts: result.guardBlindAccounts,
    duplicateAccountCount: result.duplicateAccountCount,
    durationMs: result.durationMs
  });

  return result;
}
