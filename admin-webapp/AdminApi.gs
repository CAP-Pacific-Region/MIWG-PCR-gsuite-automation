/***********************************************
 * File: AdminApi.gs
 * Description: Web app entry point and the complete list of server functions the
 * browser can reach. Every one of them starts with requireAdmin_().
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-08-16
 * Changes: 1.0.0 — initial version.
 ***********************************************/

/**
 * THIS FILE IS THE ATTACK SURFACE. Anything callable from google.script.run is
 * callable by any signed-in domain user who opens a console, whether or not the
 * page ever offers them a button — so the list below is deliberately short, and
 * every entry begins the same way:
 *
 *   const actor = requireAdmin_();
 *
 * Nothing else in the project may be added to this file without that line. The
 * authorization is not in the UI, and a function that skips it is not protected
 * by the fact that no button calls it.
 */

/**
 * Serves the UI. A caller the app cannot help gets a plain refusal rather than a
 * shell that fails on first click.
 */
function doGet() {
  const missing = admMissingConfig_();
  if (missing.length) {
    // Named in the log, not on the page: a help-desk admin cannot act on a
    // property name, and whoever can is reading the execution log anyway.
    Logger.error('Serving the not-configured page: required Script Properties are not set', {
      missing: missing
    });
    return HtmlService.createHtmlOutput(
      '<p style="font:14px/1.5 system-ui,sans-serif;padding:2rem">' +
      'This page is not set up yet.<br>' + admEscape_(admSupportSentence_()) + '</p>'
    ).setTitle('Not set up yet');
  }

  const actor = admResolveActor_();
  const privileges = admActorPrivileges_(actor);
  if (!privileges.allowed) {
    Logger.warn('Refused to serve the admin page', { user: actor || '(no identity)' });
    return HtmlService.createHtmlOutput(
      '<p style="font:14px/1.5 system-ui,sans-serif;padding:2rem">' +
      'Your account does not hold an administrator role that may use this page.<br>' +
      admEscape_(admSupportSentence_()) + '</p>'
    ).setTitle('Not authorized');
  }

  const template = HtmlService.createTemplateFromFile('Index');
  template.actor = actor;
  template.orgLabel = ADMIN_CONFIG.ORG_LABEL;

  return template.evaluate()
    .setTitle(ADMIN_CONFIG.ORG_LABEL + ' Account Admin')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * What the page needs before it can render anything: who the caller is, and
 * which groups they may change.
 *
 * @returns {Object}
 */
function apiGetState() {
  const actor = requireAdmin_();
  return {
    actor: actor.email,
    via: actor.via,
    superAdmin: actor.superAdmin,
    orgLabel: ADMIN_CONFIG.ORG_LABEL,
    supportEmail: ADMIN_CONFIG.SUPPORT_EMAIL,
    twoSvGroup: ADMIN_CONFIG.TWO_SV_GROUP,
    // So the page can offer the cadet tools site as a link rather than as text
    // an admin has to retype.
    peerAdminUrl: ADMIN_CONFIG.PEER_ADMIN_URL,
    managedGroups: admManagedGroups_(),
    auditVisible: !!ADMIN_CONFIG.AUDIT_SPREADSHEET_ID
  };
}

/**
 * The member card: everything known about whoever the admin typed.
 *
 * READ-ONLY, and it is the only call the page makes on its own — every action is
 * a button. That separation is what lets an admin look someone up without
 * wondering whether they have just changed something.
 *
 * @param {string} query - a CAPID, an email address, or a name fragment
 * @returns {Object}
 */
function apiLookup(query) {
  const actor = requireAdmin_();
  const resolved = admResolveQuery_(query);

  // Logged without an outcome of its own: a lookup is not an action, but "who
  // looked this member up" is exactly the question asked after an incident.
  Logger.info('Admin web app lookup', {
    actor: actor.email, query: String(query || '').trim(), kind: resolved.kind
  });

  const out = {
    kind: resolved.kind,
    capid: resolved.capid,
    candidates: resolved.candidates,
    member: null,
    accounts: [],
    chosen: resolved.chosen ? resolved.chosen.email : '',
    welcome: null,
    // Set when the member belongs to a tenant this app does not serve; the page
    // shows it instead of "no account — that's a provisioning job", which would
    // send a help desk off to create an account that must not exist here.
    elsewhere: '',
    canAct: true
  };

  if (resolved.member) {
    out.member = {
      capid: resolved.member.capsn,
      name: admMemberDisplayName_(resolved.member),
      type: resolved.member.type,
      status: resolved.member.status,
      expiration: resolved.member.expiration,
      orgName: resolved.member.orgName,
      charter: resolved.member.charter,
      email: resolved.member.email || '',
      secondaryEmail: resolved.member.secondaryEmail || '',
      // Shown because "why did the welcome email refuse?" is usually answered by
      // a DoNotContact flag nobody can see from the Admin console.
      primaryEmailDNC: resolved.member.primaryEmailDNC,
      secondaryEmailDNC: resolved.member.secondaryEmailDNC,
      // A member this tenant does not provision — every cadet in the wing, on a
      // seniors tenant. CAPWATCH is scoped to the wing, so they are findable
      // here and actionable nowhere on this page.
      offTenant: !admTenantProvisionsType_(resolved.member.type)
    };
    if (out.member.offTenant) out.elsewhere = admElsewhereSentence_(resolved.member.type);
  }

  out.accounts = resolved.accounts.map(function (a) {
    return {
      email: a.email,
      fullName: a.fullName,
      suspended: a.suspended,
      suspensionReason: a.suspensionReason,
      archived: a.archived,
      admin: a.admin,
      superAdmin: a.superAdmin,
      enrolledIn2Sv: a.enrolledIn2Sv,
      enforcedIn2Sv: a.enforcedIn2Sv,
      created: a.created,
      lastLogin: a.neverSignedIn ? '' : a.lastLogin,
      neverSignedIn: a.neverSignedIn,
      changePasswordAtNextLogin: a.changePasswordAtNextLogin,
      orgUnitPath: a.orgUnitPath,
      recoveryEmail: a.recoveryEmail,
      // Resolved PER ACCOUNT rather than only for the authoritative one. A
      // duplicate pair can differ in both 2SV state and group membership, and
      // showing one account's membership under another's radio button is how an
      // admin ends up adding the wrong one. Costs accounts × managed-groups
      // hasMember calls, and both numbers are 1 or 2 in practice.
      groups: admManagedGroupMembership_(a.email),
      isChosen: !!resolved.chosen && a.email === resolved.chosen.email,
      // Told to the page so it can grey the buttons rather than let an admin
      // click into a refusal. The server refuses regardless — see
      // admAssertMayActOn_().
      actionable: !a.admin || actor.superAdmin
    };
  });

  if (resolved.capid) {
    out.welcome = admWelcomeLedgerStatus_(resolved.capid);
  }
  return out;
}

/**
 * READ-ONLY. What a welcome-email resend would do to this CAPID.
 *
 * @param {string} capid
 * @param {boolean} force
 * @returns {Object}
 */
function apiPreviewWelcomeResend(capid, force) {
  requireAdmin_();
  return admPreviewWelcomeResend_(String(capid || '').trim(), force === true);
}

/**
 * Resets the password and sends the welcome email. Destructive.
 *
 * @param {string} capid
 * @param {boolean} force
 * @returns {Object}
 */
function apiResendWelcome(capid, force) {
  const actor = requireAdmin_();
  return admResendWelcome_(actor, String(capid || '').trim(), force === true);
}

/**
 * Issues a new temporary password for one account. Destructive.
 *
 * @param {string} email
 * @param {boolean} alsoEmail - also mail it to the member's off-tenant address
 * @returns {Object} including the password, which is shown to the admin and
 *   stored nowhere
 */
function apiResetPassword(email, alsoEmail) {
  const actor = requireAdmin_();
  return admResetPassword_(actor, email, alsoEmail === true);
}

/**
 * Adds or removes one account from one managed group.
 *
 * @param {string} group - must be on the managed list
 * @param {string} email
 * @param {boolean} add
 * @returns {Object}
 */
function apiSetGroupMembership(group, email, add) {
  const actor = requireAdmin_();
  return admSetGroupMembership_(actor, group, email, add === true);
}
