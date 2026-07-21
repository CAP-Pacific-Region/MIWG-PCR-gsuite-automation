/***********************************************
 * File: Directory.gs
 * Description: CAPID -> Workspace account resolution and secondary-alias address
 * arithmetic for the admin web app. Read-only except addAliasToAccount_/
 * removeAliasFromAccount_, which are the only two calls in this project that
 * change an account.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-07-20
 * Changes: 1.0.0 — initial version.
 ***********************************************/

/**
 * DUPLICATED LOGIC — keep in step with the main project.
 *
 * The four pure helpers below are ports of:
 *   normalizeSecondaryDomain_ / deriveSecondaryAlias_  src/accounts-and-groups/SecondaryDomainAliases.gs
 *   provisioningCapidsFromUser_ / findExistingAccountsByCapid_ /
 *   chooseAuthoritativeAccount_                        src/accounts-and-groups/DuplicateAccountGuard.gs
 *
 * They are copied rather than shared because this is a separate script project
 * (see Config.gs for why). If the address derivation ever diverges, this app and
 * the nightly trigger will disagree about what a member's alias is, and the
 * symptom will be a row that flips status every night. test/AliasWebApp.test.js
 * pins both copies against the same expectations.
 */

/** '@cawg.cap.gov' or 'cawg.cap.gov' -> '@cawg.cap.gov'; '' for anything unusable. */
function webappNormalizeSecondaryDomain_(configured) {
  const bare = String(configured || '').trim().toLowerCase().replace(/^@+/, '');
  if (!bare || bare.indexOf('@') !== -1 || bare.indexOf('.') === -1) return '';
  return '@' + bare;
}

/** jane.doe@cawgcap.org + @cawg.cap.gov -> jane.doe@cawg.cap.gov */
function webappDeriveSecondaryAlias_(primaryEmail, secondaryDomain) {
  const at = String(primaryEmail || '').indexOf('@');
  if (at <= 0 || !secondaryDomain) return '';
  return String(primaryEmail).toLowerCase().slice(0, at) + secondaryDomain;
}

/** CAPIDs marking a Directory User as a live account, ignoring the retired marker. */
function webappProvisioningCapidsFromUser_(user) {
  const found = {};
  (user.externalIds || []).forEach(function (id) {
    // The 'duplicate_retired_capid' marker parks a CAPID on a retired twin; it is
    // not an account the member uses, so it must not resolve here.
    if (id && id.type === 'custom' && id.customType === 'duplicate_retired_capid') return;
    const v = String(id && id.value != null ? id.value : '').trim();
    if (WEBAPP_CAPID_RE.test(v)) found[v] = true;
  });
  const emp = String(user.employeeId != null ? user.employeeId : '').trim();
  if (WEBAPP_CAPID_RE.test(emp)) found[emp] = true;
  return Object.keys(found);
}

/**
 * Ranks candidate accounts for one CAPID, most-likely-the-real-one first.
 * Ported from chooseAuthoritativeAccount_ but returns the whole sorted list: this
 * app SHOWS the duplicates to a human instead of picking silently, because the
 * canonical-looking name is often the dead twin (see the duplicate-accounts work).
 */
function webappRankAccounts_(accounts) {
  function score(a) {
    const local = String(a.email || '').split('@')[0].toLowerCase();
    const loginAt = a.neverSignedIn ? 0 : Math.max(0, new Date(a.lastLogin || 0).getTime());
    return {
      loginAt: loginAt,
      active: a.suspended ? 0 : 1,
      noSuffix: /\.\d+$/.test(local) ? 0 : 1,
      created: new Date(a.created || 0).getTime()
    };
  }
  return (accounts || []).slice().sort(function (x, y) {
    const sx = score(x), sy = score(y);
    return (sy.loginAt - sx.loginAt) ||
           (sy.active - sx.active) ||
           (sy.noSuffix - sx.noSuffix) ||
           (sy.created - sx.created);
  });
}

/**
 * Every account in the tenant carrying this CAPID, suspended ones included.
 *
 * @param {string|number} capid
 * @returns {Array<Object>} ranked; [] when the CAPID is malformed or unknown
 */
function webappFindAccountsByCapid_(capid) {
  const wanted = String(capid == null ? '' : capid).trim();
  if (!WEBAPP_CAPID_RE.test(wanted)) return [];

  const out = [];
  let pageToken = null;
  do {
    const page = AdminDirectory.Users.list({
      customer: 'my_customer',
      query: 'externalId=' + wanted,   // matches any externalId value, suspended included
      projection: 'full',
      maxResults: 200,
      pageToken: pageToken
    });
    (page.users || []).forEach(function (u) {
      // externalId= is a loose match; confirm the CAPID really is on this account.
      if (webappProvisioningCapidsFromUser_(u).indexOf(wanted) === -1) return;
      out.push({
        email: String(u.primaryEmail || '').toLowerCase(),
        name: (u.name && u.name.fullName) ? u.name.fullName : '',
        suspended: !!u.suspended,
        archived: !!u.archived,
        created: u.creationTime || null,
        lastLogin: u.lastLoginTime || null,
        neverSignedIn: !u.lastLoginTime || new Date(u.lastLoginTime).getTime() <= 0,
        orgUnitPath: u.orgUnitPath || '',
        aliases: (u.aliases || []).map(function (a) { return String(a).toLowerCase(); })
      });
    });
    pageToken = page.nextPageToken;
  } while (pageToken);

  return webappRankAccounts_(out);
}

/**
 * Re-reads one account. Used to verify a client-supplied email still belongs to
 * the CAPID the client claimed — the browser is not trusted to pair them.
 *
 * @returns {Object|null} null when the account does not exist
 */
function webappGetAccount_(primaryEmail) {
  try {
    const u = AdminDirectory.Users.get(String(primaryEmail).toLowerCase(), {
      projection: 'full'
    });
    return {
      email: String(u.primaryEmail || '').toLowerCase(),
      name: (u.name && u.name.fullName) ? u.name.fullName : '',
      suspended: !!u.suspended,
      capids: webappProvisioningCapidsFromUser_(u),
      aliases: (u.aliases || []).map(function (a) { return String(a).toLowerCase(); })
    };
  } catch (err) {
    return null;
  }
}

/** True if the secondary domain is present AND verified on this tenant. */
function webappIsDomainVerified_(secondaryDomain) {
  const bare = String(secondaryDomain || '').replace(/^@/, '').toLowerCase();
  if (!bare) return false;
  try {
    const result = AdminDirectory.Domains.list('my_customer');
    return (result.domains || []).some(function (d) {
      return String(d.domainName).toLowerCase() === bare && d.verified === true;
    });
  } catch (err) {
    Logger.error('Could not list tenant domains', { errorMessage: err.message });
    return false;
  }
}

// ============================================================================
// THE ONLY TWO WRITES IN THIS PROJECT THAT TOUCH AN ACCOUNT
//
// Both re-assert authorization even though every caller in AliasAdminApi.gs has
// already done so. That is not defensive clutter: google.script.run dispatches by
// function NAME from the browser, and while Apps Script treats a trailing
// underscore as private, the safety of this project should not rest on a naming
// convention holding. webappRemoveAliasFromAccount_ takes a caller-supplied
// account object, so if it were ever reachable it would be directly exploitable.
// The extra check is one group-membership lookup.
// ============================================================================

/**
 * Adds one directory alias.
 *
 * @returns {{status: string, conflict: boolean}} status uses the SAME vocabulary
 *   the nightly module writes to column C, so a row this app touches and a row the
 *   trigger touches are indistinguishable — and a CONFLICT written here latches
 *   under the trigger's shouldSkipLatched_ exactly as one written there does.
 */
function webappAddAliasToAccount_(primaryEmail, aliasEmail) {
  requireAuthorized_();
  try {
    AdminDirectory.Users.Aliases.insert({ alias: aliasEmail }, primaryEmail);
    Logger.info('Secondary-domain alias added via web app', { user: primaryEmail, alias: aliasEmail });
    return { status: 'ADDED — ' + aliasEmail, conflict: false };
  } catch (err) {
    const conflict = err.details && err.details.code === 409;
    Logger.error('Failed to add secondary-domain alias via web app', {
      user: primaryEmail, attemptedAlias: aliasEmail, errorMessage: err.message
    });
    return {
      status: conflict
        ? WEBAPP_CONFLICT_PREFIX + ' — ' + aliasEmail + ' is already in use by another account or group'
        : 'ERROR — ' + err.message,
      conflict: !!conflict
    };
  }
}

/**
 * Removes one directory alias — the single most dangerous call in this project,
 * so it re-checks everything rather than trusting its arguments.
 *
 * THREE GUARDS, none of which are redundant:
 *   1. The address must sit on the configured SECONDARY domain. Without this, a
 *      crafted or buggy call could strip a member's @cawgcap.org alias, which this
 *      app has no business touching and no way to restore knowledge of.
 *   2. It must not be the account's primaryEmail. Google rejects that anyway, but
 *      failing here makes the reason legible instead of an opaque 400.
 *   3. It must actually be an alias on THIS account. Deleting is otherwise a
 *      404-or-worse against whoever does hold it.
 *
 * @returns {{removed: boolean, status: string}}
 */
function webappRemoveAliasFromAccount_(account, aliasEmail, secondaryDomain) {
  requireAuthorized_();
  const alias = String(aliasEmail || '').toLowerCase();
  const suffix = String(secondaryDomain || '').toLowerCase();

  if (!suffix || !alias || alias.slice(-suffix.length) !== suffix) {
    Logger.error('Refused to remove an alias outside the secondary domain', {
      user: account.email, alias: alias, secondaryDomain: suffix
    });
    return { removed: false, status: 'REFUSED — ' + alias + ' is not on the secondary domain' };
  }
  if (alias === account.email) {
    return { removed: false, status: 'REFUSED — that is the account\'s primary address, not an alias' };
  }
  if (account.aliases.indexOf(alias) === -1) {
    // Not an error: the row may predate the alias ever being created, or the
    // alias may have been removed by hand already. The row still comes off the tab.
    return { removed: false, status: 'NOT PRESENT — ' + alias + ' was not an alias on this account' };
  }

  try {
    AdminDirectory.Users.Aliases.remove(account.email, alias);
    Logger.info('Secondary-domain alias revoked via web app', { user: account.email, alias: alias });
    return { removed: true, status: 'REVOKED — ' + alias };
  } catch (err) {
    Logger.error('Failed to revoke secondary-domain alias', {
      user: account.email, alias: alias, errorMessage: err.message
    });
    return { removed: false, status: 'ERROR — ' + err.message };
  }
}
