/***********************************************
 * File: Auth.gs
 * Description: Establishes who is calling and what they are allowed to do to
 * whom. Every entry point in AdminApi.gs passes through requireAdmin_(), and
 * every action that touches an account passes through admAssertMayActOn_().
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-08-16
 * Changes: 1.0.0 — initial version.
 ***********************************************/

/**
 * THE TRUST MODEL, because it is not obvious from the code alone.
 *
 * The app is deployed `executeAs: USER_DEPLOYING`, so every directory write runs
 * with the DEPLOYER's rights, not the caller's. That is the whole point — a help
 * desk admin who could already do all of this in the Admin console would not need
 * a web app; the ones who use this cannot reach the Apps Script editor at all.
 *
 * It also means the deployment grants nothing by itself and this file IS the
 * access control. Two gates, and both matter:
 *
 *   1. WHO MAY CALL — requireAdmin_(). A super admin, or a holder of one of the
 *      directory roles in WEBAPP_ADMIN_ROLES (default: Google's built-in Help
 *      Desk Administrator). Not a group of names someone maintains by hand: the
 *      Admin console already records this, and a second copy would drift.
 *
 *   2. WHOM THEY MAY ACT ON — admAssertMayActOn_(). A caller who is not a super
 *      admin may not reset, re-credential, or re-group ANOTHER ADMIN's account.
 *      This mirrors Google's own rule — the Admin console refuses a help-desk
 *      admin the same reset — and it is the difference between a delegated tool
 *      and a privilege-escalation ladder: without it, a help-desk volunteer could
 *      reset a super admin's password here and sign in as them.
 *
 * `access: DOMAIN` is load-bearing rather than a preference: under
 * ANYONE_ANONYMOUS getActiveUser() returns an empty address and every visitor
 * would be indistinguishable. admResolveActor_() fails closed if it is ever blank.
 *
 * Note this is one of the places Session.getActiveUser() is legitimate in this
 * codebase. In the main project it throws for want of a userinfo.email scope; in
 * a DOMAIN-restricted web app the platform supplies the caller's address without
 * one.
 */

/** Per-execution memo for directory reads that several panels want. */
const ADM_USER_MEMO = {};

/**
 * The signed-in caller's email address, lowercased.
 *
 * @returns {string} '' when no identity is available — always treat as untrusted.
 */
function admResolveActor_() {
  try {
    return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  } catch (err) {
    // A thrown identity lookup must read as "no identity", never as an error some
    // caller might swallow into a pass.
    Logger.error('Could not resolve the active user', { errorMessage: err.message });
    return '';
  }
}

/**
 * One directory user, memoized for this request.
 *
 * @param {string} email
 * @returns {Object|null} the Directory User resource, or null if unreadable
 */
function admDirectoryUser_(email) {
  const key = String(email || '').trim().toLowerCase();
  if (!key) return null;
  if (Object.prototype.hasOwnProperty.call(ADM_USER_MEMO, key)) return ADM_USER_MEMO[key];

  let user = null;
  try {
    user = AdminDirectory.Users.get(key, { projection: 'full' });
  } catch (err) {
    // 404 is the ordinary "no such account" and is not worth an ERROR; anything
    // else is a real problem, but both end the same way for the caller.
    Logger.warn('Directory lookup failed', { user: key, errorMessage: err.message });
  }
  ADM_USER_MEMO[key] = user;
  return user;
}

/**
 * The directory role names assigned to a user.
 *
 * Two API calls per role and no way around it: RoleAssignments gives role IDs,
 * and only Roles.get turns an ID into the name a human configured. Admins are a
 * handful of people, each holding one or two roles, so the cost is bounded — and
 * it is cached per request because the page asks once on load and again on each
 * action.
 *
 * @param {Object} user - a Directory User resource
 * @returns {Array<string>} role names, e.g. ['_HELP_DESK_ADMIN_ROLE']
 */
function admRoleNamesForUser_(user) {
  if (!user || !user.id) return [];

  const names = [];
  let pageToken = null;
  try {
    do {
      const page = AdminDirectory.RoleAssignments.list('my_customer', {
        userKey: user.id,
        maxResults: 200,
        pageToken: pageToken
      });
      (page.items || []).forEach(function (assignment) {
        const roleId = String(assignment.roleId || '').trim();
        if (!roleId) return;
        try {
          const role = AdminDirectory.Roles.get('my_customer', roleId);
          const name = String((role && role.roleName) || '').trim();
          if (name && names.indexOf(name) === -1) names.push(name);
        } catch (err) {
          Logger.warn('Could not read a role assigned to this user', {
            user: user.primaryEmail, roleId: roleId, errorMessage: err.message
          });
        }
      });
      pageToken = page.nextPageToken;
    } while (pageToken);
  } catch (err) {
    // Fails CLOSED: the caller ends up with no roles, so unless they are a super
    // admin or in the optional group they are refused. A role lookup that cannot
    // be completed is not a pass.
    Logger.error('Role assignment lookup failed; treating this caller as unroled', {
      user: user.primaryEmail, errorMessage: err.message
    });
    return [];
  }
  return names;
}

/**
 * What this caller is, and whether they may use the app at all.
 *
 * @param {string} email - the authenticated caller, never a client-supplied value
 * @returns {{email: string, superAdmin: boolean, roles: Array<string>,
 *   allowed: boolean, via: string}}
 */
function admActorPrivileges_(email) {
  const denied = {
    email: email || '', superAdmin: false, roles: [], allowed: false, via: 'none'
  };
  if (!email) return denied;

  if (!admIsTenantAddress_(email)) {
    // A DOMAIN deployment already restricts callers to this Workspace, so this is
    // belt-and-braces — but it is cheap, and "only ever an address we own" is
    // worth asserting here rather than trusting a manifest setting three files away.
    Logger.warn('Rejected a caller from outside this tenant', { user: email });
    return denied;
  }

  const user = admDirectoryUser_(email);
  if (!user) return denied;

  if (user.isAdmin === true) {
    return { email: email, superAdmin: true, roles: [], allowed: true, via: 'super-admin' };
  }

  const roles = admRoleNamesForUser_(user);
  const wanted = ADMIN_CONFIG.ALLOWED_ROLES.map(function (r) { return r.toLowerCase(); });
  const matched = roles.filter(function (r) { return wanted.indexOf(r.toLowerCase()) !== -1; });
  if (matched.length) {
    return { email: email, superAdmin: false, roles: roles, allowed: true, via: 'role:' + matched[0] };
  }

  // The optional escape hatch, for someone who must use the app without holding
  // an admin role. Additive only.
  if (ADMIN_CONFIG.ADMIN_GROUP) {
    try {
      const result = AdminDirectory.Members.hasMember(ADMIN_CONFIG.ADMIN_GROUP, email);
      if (result && result.isMember === true) {
        return { email: email, superAdmin: false, roles: roles, allowed: true, via: 'group' };
      }
    } catch (err) {
      // A 404 here means the configured address is wrong — a bug that would
      // otherwise present as "the app is broken for one person" with no reason in
      // the log. Say which group failed.
      Logger.error('Admin-group membership check failed; not granting access by it', {
        group: ADMIN_CONFIG.ADMIN_GROUP, user: email, errorMessage: err.message
      });
    }
  }

  return { email: email, superAdmin: false, roles: roles, allowed: false, via: 'none' };
}

/**
 * Gate for every server function the client can reach. Returns the caller's
 * privileges, which every action then consults for what it may do.
 *
 * Throwing (rather than returning a flag) is the point: a caller that forgets to
 * check a boolean silently proceeds. Keep the list of entry points in AdminApi.gs
 * short and audit it as a unit.
 *
 * @returns {Object} the privileges object from admActorPrivileges_()
 */
function requireAdmin_() {
  // Config before identity — see admMissingConfig_() in Config.gs.
  const missing = admMissingConfig_();
  if (missing.length) {
    Logger.error('Refusing every caller: required Script Properties are not set', {
      missing: missing
    });
    throw new Error('This page is not set up for ' + ADMIN_CONFIG.ORG_LABEL + ' yet. ' +
      admSupportSentence_());
  }

  const actor = admResolveActor_();
  const privileges = admActorPrivileges_(actor);
  if (!privileges.allowed) {
    Logger.warn('Rejected an unauthorized admin request', {
      user: actor || '(no identity)',
      allowedRoles: ADMIN_CONFIG.ALLOWED_ROLES
    });
    throw new Error('Your account does not hold an administrator role that may use ' +
      'this page. ' + admSupportSentence_());
  }
  return privileges;
}

/**
 * Refuses an action aimed at an account the caller may not touch.
 *
 * The rule is Google's own: only a super admin may act on an admin. Everything
 * else this app does to a member — reset their password, mail them credentials,
 * change their group membership — is exactly what a Help Desk Administrator is
 * for, and is left alone.
 *
 * @param {Object} actor - from requireAdmin_()
 * @param {Object} account - a Directory User resource for the target
 * @throws {Error} with a sentence written for the admin reading the page
 */
function admAssertMayActOn_(actor, account) {
  if (!account) throw new Error('That account could not be read from the directory.');

  const targetIsAdmin = account.isAdmin === true || account.isDelegatedAdmin === true;
  if (targetIsAdmin && !actor.superAdmin) {
    Logger.warn('Refused an action against an administrator account', {
      actor: actor.email, target: account.primaryEmail, via: actor.via
    });
    throw new Error(account.primaryEmail + ' is an administrator account. Only a super ' +
      'admin may act on one — the Admin console applies the same rule.');
  }
}
