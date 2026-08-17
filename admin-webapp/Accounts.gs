/***********************************************
 * File: Accounts.gs
 * Description: The directory half of a member lookup — which Workspace accounts
 * carry a CAPID, which one is authoritative, and what state each is in.
 * Read-only; the writes live in Actions.gs.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-08-16
 * Changes: 1.0.0 — initial version.
 ***********************************************/

/**
 * WHY THIS IS A PORT OF DuplicateAccountGuard.gs AND NOT SOMETHING SIMPLER
 *
 * "The member's account" is not a derivable address on these tenants. A member
 * may hold two accounts with one CAPID (see the duplicate population found on
 * the cadets tenant), the one they actually use is often NOT the canonically
 * named first.last, and an out-of-band account may be named nothing like it.
 * Deriving first.last@domain here would give a help desk the wrong account
 * confidently — resetting a password on an account nobody signs into, while the
 * member keeps failing to log into the one they use.
 *
 * So the lookup and the ranking are the same ones provisioning uses:
 * findExistingAccountsByCapid_() (which sees suspended accounts and every
 * externalId type) and chooseAuthoritativeAccount_() (login history first).
 * Both are copied from src/accounts-and-groups/DuplicateAccountGuard.gs and both
 * are pinned against the originals by test/AdminWebApp.test.js.
 *
 * The app SHOWS every account it finds rather than only the chosen one. When a
 * CAPID has two, that fact is the answer to the help-desk call more often than
 * anything the app can do about it.
 */

/** externalId type/customType marking a CAPID parked on a retired duplicate. */
const ADM_RETIRED_CAPID_TYPE = 'custom';
const ADM_RETIRED_CAPID_CUSTOM_TYPE = 'duplicate_retired_capid';

/**
 * CAPIDs that mark a Directory User as a LIVE account for a member, ignoring the
 * retired-twin marker. Ported from provisioningCapidsFromUser_().
 *
 * @param {Object} user - Directory User (projection: 'full')
 * @returns {Array<string>}
 */
function admCapidsFromUser_(user) {
  const found = {};
  ((user && user.externalIds) || []).forEach(function (id) {
    if (id && id.type === ADM_RETIRED_CAPID_TYPE && id.customType === ADM_RETIRED_CAPID_CUSTOM_TYPE) return;
    const v = String(id && id.value != null ? id.value : '').trim();
    if (ADM_CAPID_RE.test(v)) found[v] = true;
  });
  const emp = String(user && user.employeeId != null ? user.employeeId : '').trim();
  if (ADM_CAPID_RE.test(emp)) found[emp] = true;
  return Object.keys(found);
}

/**
 * Google returns the Unix epoch for accounts that never signed in, rather than
 * omitting the field. Treat epoch and missing alike. Ported from
 * dupGuardNeverLoggedIn_().
 *
 * @param {string} lastLoginTime
 * @returns {boolean}
 */
function admNeverLoggedIn_(lastLoginTime) {
  if (!lastLoginTime) return true;
  return new Date(lastLoginTime).getTime() <= 0;
}

/**
 * The account fields this app reports and acts on, from a Directory User.
 *
 * @param {Object} u - Directory User (projection: 'full')
 * @returns {Object}
 */
function admAccountSummary_(u) {
  return {
    email: u.primaryEmail,
    fullName: (u.name && u.name.fullName) || '',
    suspended: !!u.suspended,
    suspensionReason: u.suspensionReason || '',
    archived: !!u.archived,
    admin: u.isAdmin === true || u.isDelegatedAdmin === true,
    superAdmin: u.isAdmin === true,
    enrolledIn2Sv: u.isEnrolledIn2Sv === true,
    enforcedIn2Sv: u.isEnforcedIn2Sv === true,
    created: u.creationTime || null,
    lastLogin: u.lastLoginTime || null,
    neverSignedIn: admNeverLoggedIn_(u.lastLoginTime),
    changePasswordAtNextLogin: u.changePasswordAtNextLogin === true,
    orgUnitPath: u.orgUnitPath || '',
    recoveryEmail: u.recoveryEmail || '',
    aliases: u.aliases || [],
    capids: admCapidsFromUser_(u)
  };
}

/**
 * Live directory lookup for accounts carrying a CAPID, INCLUDING suspended ones
 * and every externalId type. Ported from findExistingAccountsByCapid_().
 *
 * @param {string|number} capid
 * @returns {Array<Object>} account summaries
 */
function admAccountsByCapid_(capid) {
  const wanted = String(capid).trim();
  const out = [];
  if (!ADM_CAPID_RE.test(wanted)) return out;

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
        if (admCapidsFromUser_(u).indexOf(wanted) === -1) return;
        out.push(admAccountSummary_(u));
      });
      pageToken = page.nextPageToken;
    } while (pageToken);
  } catch (e) {
    // Surfaced rather than swallowed: unlike provisioning, which can fall back to
    // inserting, an empty result here would tell an admin the member has no
    // account and send them off to create a duplicate.
    Logger.error('CAPID account lookup failed', { capsn: wanted, errorMessage: e.message });
    throw new Error('The directory could not be searched for CAPID ' + wanted +
      '. Try again in a moment.');
  }
  return out;
}

/** Lowercased email localpart with any trailing ".N" collision suffix removed. */
function admCanonicalLocalpart_(email) {
  const local = String(email || '').split('@')[0].toLowerCase();
  const parts = local.split('.');
  if (parts.length > 1 && /^\d+$/.test(parts[parts.length - 1])) parts.pop();
  return parts.join('.');
}

/**
 * Picks the authoritative account among several sharing a CAPID. Ported verbatim
 * from chooseAuthoritativeAccount_() in DuplicateAccountGuard.gs, whose header
 * explains why login history outranks a canonical-looking address. Pure.
 *
 * @param {Array<Object>} accounts
 * @param {string} canonicalLocal - firstname.lastname, or '' if unknown
 * @param {string} [preferredDomain] - '@x.org' or 'x.org'; '' disables that rule
 * @returns {Object|null}
 */
function admChooseAuthoritativeAccount_(accounts, canonicalLocal, preferredDomain) {
  if (!accounts || !accounts.length) return null;
  const canon = String(canonicalLocal || '').toLowerCase();

  const bareDom = String(preferredDomain || '').trim().toLowerCase().replace(/^@+/, '');
  const domSuffix = bareDom ? '@' + bareDom : '';

  function score(a) {
    const email = String(a.email || '').toLowerCase();
    const local = email.split('@')[0];
    const loginAt = a.neverSignedIn ? 0 : Math.max(0, new Date(a.lastLogin || 0).getTime());
    return {
      loginAt: loginAt,
      active: a.suspended ? 0 : 1,
      onDomain: (domSuffix && email.slice(-domSuffix.length) === domSuffix) ? 1 : 0,
      exact: (canon && local === canon) ? 1 : 0,
      noSuffix: /\.\d+$/.test(local) ? 0 : 1,
      created: new Date(a.created || 0).getTime()
    };
  }

  return accounts.slice().sort(function (x, y) {
    const sx = score(x), sy = score(y);
    return (sy.loginAt - sx.loginAt) ||
           (sy.active - sx.active) ||
           (sy.onDomain - sx.onDomain) ||
           (sy.exact - sx.exact) ||
           (sy.noSuffix - sx.noSuffix) ||
           (sy.created - sx.created);
  })[0];
}

/**
 * Resolves whatever an admin typed into a member and their accounts.
 *
 * Three shapes, decided by the input and not by a mode the admin has to pick:
 *   - 5-7 digits              -> a CAPID
 *   - anything with an '@'    -> an account address (the CAPID comes off the account)
 *   - anything else           -> a name fragment, which returns CANDIDATES only
 *
 * The name case deliberately stops at a list. Acting on "the first Smith" is
 * exactly the mistake this app must not make easy, so the admin picks a CAPID
 * and the lookup runs again from there.
 *
 * @param {string} query
 * @returns {{kind: string, capid: string, member: Object|null,
 *   accounts: Array<Object>, chosen: Object|null, candidates: Array<Object>}}
 */
function admResolveQuery_(query) {
  const q = String(query || '').trim();
  const empty = {
    kind: 'none', capid: '', member: null, accounts: [], chosen: null, candidates: []
  };
  if (!q) return empty;

  if (ADM_CAPID_RE.test(q)) return admResolveByCapid_(q);

  if (q.indexOf('@') !== -1) {
    const user = admDirectoryUser_(q);
    if (!user) {
      return Object.assign({}, empty, { kind: 'no-such-account' });
    }
    const capids = admCapidsFromUser_(user);
    if (capids.length === 1) return admResolveByCapid_(capids[0]);

    // An account with no CAPID, or with two, is a real finding and is shown as
    // itself rather than resolved to a member the app would be guessing at.
    return {
      kind: capids.length ? 'ambiguous-capid' : 'account-only',
      capid: '', member: null,
      accounts: [admAccountSummary_(user)],
      chosen: admAccountSummary_(user),
      candidates: []
    };
  }

  return Object.assign({}, empty, {
    kind: 'candidates',
    candidates: admSearchMembersByName_(q, ADM_SEARCH_LIMIT)
  });
}

/**
 * The full picture for one CAPID: the CAPWATCH record, every account carrying
 * it, and which of those is authoritative.
 *
 * @param {string} capid
 * @returns {Object} same shape as admResolveQuery_()
 */
function admResolveByCapid_(capid) {
  const wanted = String(capid).trim();
  const member = admBuildMemberRecord_(wanted);
  const accounts = admAccountsByCapid_(wanted);
  const canonicalLocal = member
    ? (member.firstName + '.' + member.lastName).toLowerCase().replace(/\s+/g, '')
    : '';

  return {
    kind: 'member',
    capid: wanted,
    member: member,
    accounts: accounts,
    chosen: admChooseAuthoritativeAccount_(accounts, canonicalLocal, ADMIN_CONFIG.EMAIL_DOMAIN),
    candidates: []
  };
}

/**
 * Which of the managed groups an address belongs to.
 *
 * hasMember() per group rather than one listing per member: the managed list is
 * short by design, and hasMember answers for indirect membership too, which is
 * the honest answer to "is this member in the 2SV group?" when the group nests.
 *
 * @param {string} email
 * @returns {Array<{group: string, member: boolean, error: string}>}
 */
function admManagedGroupMembership_(email) {
  return admManagedGroups_().map(function (group) {
    try {
      const result = AdminDirectory.Members.hasMember(group, email);
      return { group: group, member: result && result.isMember === true, error: '' };
    } catch (err) {
      // Reported per group, not thrown: one mistyped address in
      // WEBAPP_MANAGED_GROUPS must not take the whole member card down.
      Logger.warn('Group membership check failed', {
        group: group, user: email, errorMessage: err.message
      });
      return { group: group, member: false, error: 'Could not be checked (' + err.message + ')' };
    }
  });
}

/**
 * Every group this app may modify: the 2SV setup group, plus any extras.
 * De-duplicated and lowercased, because this list is compared against a
 * client-supplied address in admGroupFromRequest_().
 *
 * @returns {Array<string>}
 */
function admManagedGroups_() {
  const all = [ADMIN_CONFIG.TWO_SV_GROUP].concat(ADMIN_CONFIG.MANAGED_GROUPS);
  const seen = {};
  return all
    .map(function (g) { return String(g || '').trim().toLowerCase(); })
    .filter(function (g) {
      if (!g || seen[g]) return false;
      seen[g] = true;
      return true;
    });
}
