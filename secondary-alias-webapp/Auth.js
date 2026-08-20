/***********************************************
 * File: Auth.gs
 * Description: Identifies the caller and decides whether they may manage secondary
 * aliases. Every entry point in AliasAdminApi.gs passes through requireAuthorized_().
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-07-20
 * Changes: 1.0.0 — initial version.
 ***********************************************/

/**
 * THE TRUST MODEL, because it is not obvious from the code alone.
 *
 * The app is deployed `executeAs: USER_DEPLOYING`, so every Directory API call
 * runs with the deployer's super-admin rights — NOT the visitor's. That is
 * deliberate: it is what lets a unit IT officer who is not a Workspace admin add
 * an alias. It also means the deployment itself grants nothing and this file is
 * the entire access control. Treat it accordingly.
 *
 * `access: DOMAIN` is what makes an identity exist at all. Under
 * ANYONE_ANONYMOUS, getActiveUser() returns an empty email and every visitor
 * would be indistinguishable — so the manifest setting is load-bearing, not a
 * preference. resolveActor_() fails closed if it ever comes back blank.
 *
 * Note this is the one place Session.getActiveUser() is legitimate in this
 * codebase. In the main project it throws for want of a userinfo.email scope; in
 * a DOMAIN-restricted web app the platform supplies the same-domain caller's
 * address without one.
 */

/**
 * The signed-in caller's email address, lowercased.
 *
 * @returns {string} '' when no identity is available — always treat as untrusted.
 */
function resolveActor_() {
  try {
    return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  } catch (err) {
    // Should not happen on a DOMAIN deployment, but a thrown identity lookup must
    // read as "no identity", never as an error that some caller might swallow.
    Logger.error('Could not resolve the active user', { errorMessage: err.message });
    return '';
  }
}

/**
 * True if `email` is a member of the configured admin group, directly or through
 * a nested group (hasMember resolves nesting).
 *
 * Fails CLOSED on every uncertainty: no group configured, no caller, or an API
 * error all return false. An outage must not open the door.
 *
 * @param {string} email
 * @returns {boolean}
 */
function isAliasAdmin_(email) {
  const group = WEBAPP_CONFIG.ADMIN_GROUP;
  if (!group) {
    Logger.error('WEBAPP_ALIAS_ADMIN_GROUP is not set; denying everyone.', {
      hint: 'Set it in Project Settings > Script Properties to the group address that grants access.'
    });
    return false;
  }
  if (!email) return false;

  try {
    const result = AdminDirectory.Members.hasMember(group, email);
    return result && result.isMember === true;
  } catch (err) {
    // A 404 here means the group address is wrong — a configuration bug that
    // would otherwise present as "the app is broken for everyone" with no reason
    // in the log. Say which group failed.
    Logger.error('Admin-group membership check failed; denying access', {
      group: group,
      user: email,
      errorMessage: err.message
    });
    return false;
  }
}

/**
 * Gate for every server function the client can reach. Returns the actor's email
 * so callers can attribute the change; throws otherwise.
 *
 * Throwing (rather than returning a flag) is the point: a caller that forgets to
 * check a boolean silently proceeds, whereas a caller that forgets to call this
 * at all is the bug this cannot prevent — so keep the list of entry points in
 * AliasAdminApi.gs short and audit it as a unit.
 *
 * @returns {string} the authorized caller's email
 */
function requireAuthorized_() {
  const actor = resolveActor_();
  if (!isAliasAdmin_(actor)) {
    Logger.warn('Rejected an unauthorized alias-admin request', {
      user: actor || '(no identity)',
      group: WEBAPP_CONFIG.ADMIN_GROUP || '(unset)'
    });
    // Deliberately vague to the client: naming the group would tell an
    // unauthorized visitor exactly what to ask to be added to.
    throw new Error('You are not authorized to manage secondary aliases.');
  }
  return actor;
}
