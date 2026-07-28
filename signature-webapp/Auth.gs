/***********************************************
 * File: Auth.gs
 * Description: Identifies the caller. Every entry point in SignatureApi.gs passes
 * through requireMember_(), and the address it returns is the ONLY account the
 * rest of the app will ever read or write.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-07-28
 * Changes: 1.0.0 — initial version.
 ***********************************************/

/**
 * THE TRUST MODEL, because it is not obvious from the code alone.
 *
 * The app is deployed `executeAs: USER_DEPLOYING`, so the CAPWATCH read and the
 * service-account impersonation run with the deployer's rights, not the visitor's.
 * That is what makes self-service possible at all: a member cannot read the
 * CAPWATCH extract and cannot mint a Gmail settings token, and neither ability
 * should be handed to them just so they can format their own name.
 *
 * It also means the deployment grants nothing by itself, and the whole of the
 * access control is: (1) this file establishes WHO is calling, and (2) every
 * server function derives the target account from that identity alone. The client
 * never names an account. There is no admin-acts-for-a-member mode — the
 * bulk pushAllSignatures() in src/ already covers that ground — so no request can
 * ever reach a mailbox other than the caller's own.
 *
 * `access: DOMAIN` is load-bearing, not a preference: under ANYONE_ANONYMOUS
 * getActiveUser() returns an empty address and every visitor would be
 * indistinguishable. resolveActor_() fails closed if it ever comes back blank.
 *
 * Note this is one of the two places Session.getActiveUser() is legitimate in this
 * codebase (webapp/Auth.gs is the other). In the main project it throws for want
 * of a userinfo.email scope; in a DOMAIN-restricted web app the platform supplies
 * the same-domain caller's address without one.
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
 * True if this caller may use the app.
 *
 * With SIGNATURE_WEBAPP_ALLOWED_GROUP unset, any authenticated domain user may —
 * see the property's note in Config.gs for why that default is the right one here.
 * With it set, membership is required and every uncertainty (no caller, API error)
 * denies: a group check that cannot be completed is not a pass.
 *
 * @param {string} email
 * @returns {boolean}
 */
function mayUseSignatureApp_(email) {
  if (!email) return false;
  if (!isOnATenantDomain_(email)) {
    // A DOMAIN deployment already restricts callers to this Workspace, so this is
    // belt-and-braces — but the address is about to become the impersonation
    // subject, and "only ever an address we own" is worth asserting where it is
    // cheap rather than trusting a manifest setting three files away.
    Logger.warn('Rejected a caller from outside this tenant', { user: email });
    return false;
  }

  const group = SIG_CONFIG.ALLOWED_GROUP;
  if (!group) return true;

  try {
    const result = AdminDirectory.Members.hasMember(group, email);
    return result && result.isMember === true;
  } catch (err) {
    // A 404 here means the group address is wrong — a configuration bug that would
    // otherwise present as "the app is broken for everyone" with no reason in the
    // log. Say which group failed.
    Logger.error('Allowed-group membership check failed; denying access', {
      group: group,
      user: email,
      errorMessage: err.message
    });
    return false;
  }
}

/**
 * True if an address sits on a domain this tenant hands out — its primary mail
 * domain, or the secondary domain if one is configured.
 *
 * Ported from isOrgOwnedSendAs_() in src/accounts-and-groups/UpdateMembers.gs.
 * Compares the full domain, never a suffix: a lookalike @cawg.cap.gov.example.com
 * must not pass. Accepts TENANT_SECONDARY_EMAIL_DOMAIN with or without a leading
 * '@'; it is set bare on at least one tenant.
 *
 * @param {string} email
 * @returns {boolean}
 */
function isOnATenantDomain_(email) {
  const addr = String(email || '').trim().toLowerCase();
  const at = addr.lastIndexOf('@');
  if (at < 0) return false;
  const domain = addr.slice(at + 1);
  if (!domain) return false;

  return [SIG_CONFIG.EMAIL_DOMAIN, SIG_CONFIG.SECONDARY_EMAIL_DOMAIN]
    .map(function (d) { return String(d || '').trim().toLowerCase().replace(/^@+/, ''); })
    .filter(Boolean)
    .indexOf(domain) !== -1;
}

/**
 * Gate for every server function the client can reach. Returns the caller's own
 * address, which is the only account any of them may act on.
 *
 * Throwing (rather than returning a flag) is the point: a caller that forgets to
 * check a boolean silently proceeds. Keep the list of entry points in
 * SignatureApi.gs short and audit it as a unit.
 *
 * @returns {string} the authorized caller's email
 */
function requireMember_() {
  // Config before identity: an unconfigured tenant must not read as "you are not
  // authorized". See sigMissingConfig_() in Config.gs.
  const missing = sigMissingConfig_();
  if (missing.length) {
    Logger.error('Refusing every caller: required Script Properties are not set', {
      missing: missing
    });
    throw new Error('This page is not set up for ' + SIG_CONFIG.ORG_LABEL + ' yet. ' +
      sigSupportSentence_());
  }

  const actor = resolveActor_();
  if (!mayUseSignatureApp_(actor)) {
    Logger.warn('Rejected an unauthorized signature request', {
      user: actor || '(no identity)',
      allowedGroup: SIG_CONFIG.ALLOWED_GROUP || '(unrestricted)'
    });
    throw new Error('You are not signed in to a ' + SIG_CONFIG.ORG_LABEL +
      ' account that may use this page.');
  }
  return actor;
}
