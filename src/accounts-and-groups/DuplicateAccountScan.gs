/**
 * Duplicate Workspace Account Scanner  (READ-ONLY diagnostic)
 *
 * Version: 1.5.0
 * Date: 2026-07-19
 *
 * Changes: 1.5.0 — (a) the KEEP/retire preview passes CONFIG.EMAIL_DOMAIN to
 *   chooseAuthoritativeAccount_ (see DuplicateAccountGuard.gs 1.3.0), so an account on
 *   the tenant's configured domain outranks a legacy-domain twin by policy instead of
 *   the created-date tiebreak. (b) Added scanDerivedAddressDrift(): for every account
 *   whose CAPID matches a CAPWATCH member, compares the address it HAS against what
 *   provisioning would DERIVE today, classified (collision-suffix / punctuation /
 *   reversed / initial / other). Settles where the odd addresses came from: a
 *   'punctuation' drift means the CAPWATCH name lacks the punctuation the address
 *   carries — i.e. baseEmail derived faithfully and the odd address predates
 *   provisioning. These mismatches are stable (the CAPID map updates the account in
 *   place), so they are documentation, not defects.
 *   1.4.0 — accounts in administrative / NHQ-staff org units
 *   (DUP_SCAN_EXEMPT_OU_SUBSTRINGS: "zz-administrative", "nhq employees") are blanket-exempt
 *   from needing a CAPID. Those containers hold role, workstation and staff identities that
 *   are not CAP members and are never provisioned from CAPWATCH, so they were noise in the
 *   UNKNOWN bucket. Matched case-insensitively as a substring because the same unit is
 *   spelled differently per tenant (/ZZ-Administrative vs /zz-Administrative). This is the
 *   ONE signal that outranks a CAPWATCH roster match — the exemption is deliberately
 *   blanket — but a roster match is still PRINTED on the account so the fact stays visible.
 *   1.3.0 — added scanAccountsWithoutCapid(), which LISTS the accounts the
 *   duplicate scan can only count. An account with no readable CAPID is invisible to
 *   BOTH the provisioning map and the duplicate-create guard's `externalId=` lookup, so
 *   if it belongs to a real member, provisioning cannot match them and will create a
 *   SECOND account — the remaining hole in the guard. Each is cross-referenced against
 *   the CAPWATCH roster by derived first.last localpart, so a member account that merely
 *   LOST its CAPID tag is called out for re-tagging. Role/service accounts (IT,
 *   automation, admin, …) are classified but STILL LISTED — never silently hidden, since
 *   a silent exemption is how a real member's untagged account would go unnoticed.
 *   1.2.0 — accounts already retired by suspendOrphanDuplicates (whose only
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
    const authoritative = chooseAuthoritativeAccount_(accounts, '', CONFIG.EMAIL_DOMAIN);
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

/* ===========================================================================
 * ACCOUNTS WITH NO READABLE CAPID
 * ======================================================================== */

/**
 * Localparts that normally belong to role / service accounts rather than members.
 * Matching one only CLASSIFIES an account — it is still listed in the report. A
 * silent exemption is precisely how a real member's untagged account would slip by.
 */
var DUP_SCAN_ROLE_LOCALPARTS = [
  'it', 'automation', 'admin', 'administrator', 'postmaster', 'abuse', 'noreply',
  'no-reply', 'donotreply', 'support', 'help', 'helpdesk', 'info', 'webmaster',
  'security', 'billing', 'notifications', 'calendar', 'test'
];

/**
 * Org units whose accounts are BLANKET-exempt from needing a CAPID: administrative
 * and NHQ-staff containers hold role, workstation and staff identities that are not
 * CAP members and are never provisioned from CAPWATCH.
 *
 * Matched case-insensitively as a substring of orgUnitPath, because the same unit is
 * spelled differently per tenant (`/ZZ-Administrative` on seniors, `/zz-Administrative`
 * on cadets).
 *
 * Unlike every other signal here, this one OUTRANKS a CAPWATCH roster match — an OU
 * exemption is deliberately blanket. The roster match is still printed on the account
 * when there is one, so the fact stays visible rather than being silently swallowed.
 */
var DUP_SCAN_EXEMPT_OU_SUBSTRINGS = ['zz-administrative', 'nhq employees'];

/**
 * Classifies an account as role/service-like, with the reasons why. Pure.
 * @param {Object} user Directory User (projection:'full')
 * @returns {{isRole: boolean, exemptOu: boolean, reasons: string[]}}
 */
function looksLikeRoleAccount_(user) {
  const reasons = [];
  const local = String((user && user.primaryEmail) || '').split('@')[0].toLowerCase();
  const ou = String((user && user.orgUnitPath) || '').toLowerCase();

  let exemptOu = false;
  DUP_SCAN_EXEMPT_OU_SUBSTRINGS.forEach(function (frag) {
    if (ou.indexOf(frag) > -1) exemptOu = true;
  });
  if (exemptOu) reasons.push('exempt-ou');

  if (DUP_SCAN_ROLE_LOCALPARTS.indexOf(local) > -1) reasons.push('role-localpart');
  if (user && user.isAdmin) reasons.push('super-admin');
  if (user && user.isDelegatedAdmin) reasons.push('delegated-admin');

  return { isRole: reasons.length > 0, exemptOu: exemptOu, reasons: reasons };
}

/**
 * The localpart provisioning would derive for a member — mirrors `baseEmail` in
 * addOrUpdateUser (lowercase, whitespace stripped, hyphens KEPT). Used to spot an
 * account that IS a member's, just missing its CAPID tag. Pure.
 * @returns {string} e.g. "jane.doe"
 */
function derivedLocalpartForMember_(member) {
  return (String((member && member.firstName) || '') + '.' +
          String((member && member.lastName) || ''))
    .toLowerCase().replace(/\s+/g, '');
}

/**
 * READ-ONLY. Lists every account carrying no readable CAPID, classified so the
 * genuinely concerning ones surface.
 *
 * WHY IT MATTERS: the provisioning map and the duplicate-create guard both locate a
 * member's existing account BY CAPID. An account with none is invisible to both — so
 * if it belongs to a real member, provisioning will not match them and will create a
 * SECOND account. These are the remaining hole in the duplicate guard.
 *
 * Buckets:
 *   LIKELY MEMBER — localpart matches a CAPWATCH member's derived first.last. Almost
 *                   certainly a member account that lost (or never got) its CAPID tag.
 *   UNKNOWN       — neither role-like nor roster-matched. Needs a human.
 *   ROLE/SERVICE  — IT, automation, admin, etc. Listed, not hidden.
 *
 * @returns {Object} summary + the three buckets
 */
function scanAccountsWithoutCapid() {
  const start = new Date();
  Logger.info('No-CAPID account scan starting (read-only)');

  // CAPWATCH roster keyed by the localpart provisioning would derive. Best effort:
  // without it the cross-reference is skipped, not fatal.
  const memberByLocalpart = {};
  let rosterSize = 0;
  try {
    const members = getMembers();
    Object.keys(members).forEach(function (capid) {
      const lp = derivedLocalpartForMember_(members[capid]);
      if (lp && lp !== '.') { memberByLocalpart[lp] = capid; rosterSize++; }
    });
  } catch (e) {
    Logger.warn('CAPWATCH roster unavailable — skipping the roster cross-reference', {
      errorMessage: e.message
    });
  }

  const likelyMember = [], unknown = [], role = [];
  let scannedUsers = 0;
  let pageToken = null;

  do {
    const page = AdminDirectory.Users.list({
      customer: 'my_customer',
      maxResults: 500,
      projection: 'full',
      pageToken: pageToken
    });

    (page.users || []).forEach(function (u) {
      scannedUsers++;
      if (extractCapidsFromUser_(u).length > 0) return;   // has a CAPID; not our problem

      const local = String(u.primaryEmail || '').split('@')[0].toLowerCase();
      const cls = looksLikeRoleAccount_(u);
      const rosterCapid = memberByLocalpart[local] || null;

      const info = {
        email: u.primaryEmail,
        fullName: (u.name && u.name.fullName) || '',
        created: u.creationTime || null,
        lastLogin: u.lastLoginTime || null,
        neverSignedIn: neverLoggedIn_(u.lastLoginTime),
        suspended: !!u.suspended,
        orgUnitPath: u.orgUnitPath || '',
        aliases: (u.aliases || []).length,
        classification: cls.reasons,
        rosterCapid: rosterCapid
      };

      // An OU exemption is blanket and wins outright (admin / NHQ-staff containers hold
      // non-member identities). Otherwise a roster match outranks a role hint: a real
      // member on a role-shaped localpart is still an account provisioning cannot see.
      if (cls.exemptOu) role.push(info);
      else if (rosterCapid) likelyMember.push(info);
      else if (cls.isRole) role.push(info);
      else unknown.push(info);
    });

    pageToken = page.nextPageToken;
  } while (pageToken);

  const total = likelyMember.length + unknown.length + role.length;

  // Summary as its own entry — Apps Script truncates one oversized message.
  Logger.info('===== ACCOUNTS WITH NO READABLE CAPID =====' +
    '\nScanned users .................... ' + scannedUsers +
    '\nNo readable CAPID ............... ' + total +
    '\n  LIKELY MEMBER (needs re-tag) .. ' + likelyMember.length +
    '   <-- provisioning cannot see these' +
    '\n  UNKNOWN (needs review) ........ ' + unknown.length +
    '\n  role / service ................ ' + role.length +
    '\nCAPWATCH roster entries ......... ' + rosterSize +
    (rosterSize ? '' : '   (roster unavailable — matches skipped)'));

  function render_(title, list, note) {
    if (!list.length) return;
    const lines = ['===== ' + title + ' (' + list.length + ') ====='];
    if (note) lines.push(note, '');
    list.sort(function (a, b) { return String(a.email).localeCompare(String(b.email)); });
    list.forEach(function (a) {
      lines.push('    ' + a.email + (a.fullName ? '   "' + a.fullName + '"' : '') +
        '\n        created=' + (a.created || '?') +
        '  ' + (a.suspended ? 'SUSPENDED' : 'active') +
        '  ' + (a.neverSignedIn ? 'never-signed-in' : 'last-login=' + a.lastLogin) +
        '  ou=' + a.orgUnitPath +
        (a.aliases ? '  aliases=' + a.aliases : '') +
        (a.rosterCapid ? '\n        ROSTER MATCH -> CAPID ' + a.rosterCapid : '') +
        (a.classification.length ? '\n        classified: ' + a.classification.join(',') : ''));
    });
    Logger.info(lines.join('\n'));
  }

  render_('LIKELY MEMBER ACCOUNTS MISSING A CAPID', likelyMember,
    'These localparts match a CAPWATCH member. Provisioning cannot match them by CAPID,\n' +
    'so it will create a SECOND account for them. Fix: set the organization externalId to\n' +
    'the CAPID shown (Admin console "Employee ID"), then re-run the duplicate scan.');
  render_('UNKNOWN — NO CAPID, NOT ROLE-LIKE, NOT ON THE ROSTER', unknown,
    'Could be departed members, shared mailboxes, or test accounts. Needs a human.');
  render_('ROLE / SERVICE ACCOUNTS (expected to have no CAPID)', role,
    'Listed for completeness — classified, not hidden.');

  const result = {
    scannedUsers: scannedUsers,
    totalWithoutCapid: total,
    likelyMember: likelyMember,
    unknown: unknown,
    role: role,
    rosterSize: rosterSize,
    durationMs: new Date() - start
  };
  Logger.info('No-CAPID account scan complete', {
    totalWithoutCapid: total,
    likelyMember: likelyMember.length,
    unknown: unknown.length,
    role: role.length,
    durationMs: result.durationMs
  });
  return result;
}

/* ===========================================================================
 * DERIVED-ADDRESS DRIFT
 * ======================================================================== */

/**
 * Classifies how an account's actual localpart differs from what provisioning
 * would derive from the member's CAPWATCH name today. Pure.
 *
 *   'match'            — identical; provisioning would derive exactly this address
 *   'collision-suffix' — identical once the account's trailing .N is dropped
 *   'punctuation'      — identical once hyphens/apostrophes etc. are stripped from
 *                        BOTH (baseEmail strips only whitespace, so this class means
 *                        the CAPWATCH name and the address disagree on punctuation)
 *   'reversed'         — same tokens, opposite order (last.first vs first.last)
 *   'initial'          — one side abbreviates the first token to its initial
 *   'other'            — none of the above; needs eyes
 *
 * @param {string} accountLocal the localpart the account actually has
 * @param {string} derivedLocal derivedLocalpartForMember_(rosterMember)
 * @returns {string}
 */
function classifyLocalpartDrift_(accountLocal, derivedLocal) {
  const raw = String(accountLocal || '').toLowerCase();
  const d = String(derivedLocal || '').toLowerCase();
  if (!raw || !d || d === '.') return 'other';
  if (raw === d) return 'match';

  const a = raw.replace(/\.\d+$/, '');            // drop a trailing collision suffix
  if (a === d) return 'collision-suffix';

  const strip = function (s) { return s.replace(/[^a-z0-9.]/g, ''); };
  if (strip(a) === strip(d)) return 'punctuation';

  const at = a.split('.'), dt = d.split('.');
  if (at.length === dt.length && at.length >= 2 &&
      at.slice().reverse().join('.') === d) return 'reversed';

  if (at.length === dt.length && at.length >= 2 &&
      at.slice(1).join('.') === dt.slice(1).join('.') &&
      ((at[0].length === 1 && dt[0].charAt(0) === at[0]) ||
       (dt[0].length === 1 && at[0].charAt(0) === dt[0]))) return 'initial';

  return 'other';
}

/**
 * READ-ONLY. For every account whose CAPID matches a CAPWATCH member, compares the
 * address the account HAS against the address provisioning would DERIVE from the
 * member's CAPWATCH name today, and reports every mismatch, classified.
 *
 * WHY: this settles where the odd addresses came from. If the in-use hyphenated /
 * suffixed addresses show up here with a hyphen-free derived counterpart, then
 * CAPWATCH's name fields are hyphen-free and baseEmail has been deriving faithfully
 * all along — the odd addresses predate provisioning (original population) and there
 * is no derivation bug to fix. The mismatches themselves are FINE and stable: the
 * provisioning map resolves members by CAPID, so these accounts are updated in
 * place at their existing address and no twin is created.
 *
 * Accounts whose CAPID is only the retired marker are skipped (intentionally parked).
 *
 * @returns {Object} summary + mismatches
 */
function scanDerivedAddressDrift() {
  const start = new Date();
  Logger.info('Derived-address drift scan starting (read-only)');

  let members;
  try {
    members = getMembers();
  } catch (e) {
    Logger.error('CAPWATCH roster unavailable — drift scan needs it to derive addresses', {
      errorMessage: e.message
    });
    return { error: 'roster unavailable' };
  }

  const mismatches = [];
  const counts = { match: 0 };
  let scannedUsers = 0, compared = 0;
  let pageToken = null;

  do {
    const page = AdminDirectory.Users.list({
      customer: 'my_customer',
      maxResults: 500,
      projection: 'full',
      pageToken: pageToken
    });

    (page.users || []).forEach(function (u) {
      scannedUsers++;
      const email = String(u.primaryEmail || '');
      const local = email.split('@')[0].toLowerCase();

      extractCapidsFromUser_(u).forEach(function (capid) {
        // Retired-marker-only carriage = deliberately parked; not provisioning's account.
        const carriers = capidCarriersForUser_(u, capid);
        if (carriers.length > 0 && carriers.every(function (c) {
          return c === DUP_SCAN_RETIRED_CARRIER;
        })) return;

        const m = members[capid];
        if (!m) return;                      // not on the roster (lapsed etc.) — nothing to derive
        compared++;

        const derived = derivedLocalpartForMember_(m);
        const cls = classifyLocalpartDrift_(local, derived);
        counts[cls] = (counts[cls] || 0) + 1;
        if (cls === 'match') return;

        mismatches.push({
          email: email,
          derivedLocalpart: derived,
          drift: cls,
          suspended: !!u.suspended,
          neverSignedIn: neverLoggedIn_(u.lastLoginTime),
          lastLogin: u.lastLoginTime || null
        });
      });
    });

    pageToken = page.nextPageToken;
  } while (pageToken);

  Logger.info('===== DERIVED-ADDRESS DRIFT =====' +
    '\nScanned users .................... ' + scannedUsers +
    '\nCompared against the roster ..... ' + compared +
    '\n  match (derives exactly) ....... ' + (counts.match || 0) +
    '\n  collision-suffix .............. ' + (counts['collision-suffix'] || 0) +
    '\n  punctuation ................... ' + (counts.punctuation || 0) +
    '   <-- CAPWATCH name lacks the punctuation the address has' +
    '\n  reversed ...................... ' + (counts.reversed || 0) +
    '\n  initial ....................... ' + (counts.initial || 0) +
    '\n  other ......................... ' + (counts.other || 0));

  const CHUNK = 15;
  mismatches.sort(function (a, b) {
    return (a.drift > b.drift ? 1 : a.drift < b.drift ? -1 : 0) ||
           String(a.email).localeCompare(String(b.email));
  });
  for (let i = 0; i < mismatches.length; i += CHUNK) {
    const lines = ['===== DRIFT DETAIL ' + (i + 1) + '-' +
                   Math.min(i + CHUNK, mismatches.length) + ' of ' + mismatches.length + ' ====='];
    mismatches.slice(i, i + CHUNK).forEach(function (mm) {
      lines.push('    [' + mm.drift + '] ' + mm.email +
        '\n        CAPWATCH would derive: ' + mm.derivedLocalpart +
        '  ' + (mm.suspended ? 'SUSPENDED' : 'active') +
        '  ' + (mm.neverSignedIn ? 'never-signed-in' : 'last-login=' + mm.lastLogin));
    });
    Logger.info(lines.join('\n'));
  }

  const result = {
    scannedUsers: scannedUsers,
    compared: compared,
    counts: counts,
    mismatches: mismatches,
    durationMs: new Date() - start
  };
  Logger.info('Derived-address drift scan complete', {
    compared: compared,
    mismatches: mismatches.length,
    durationMs: result.durationMs
  });
  return result;
}
