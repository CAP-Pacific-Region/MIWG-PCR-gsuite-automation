/***********************************************
 * File: Actions.gs
 * Description: The four things this app can actually do to an account, and every
 * guard that decides whether it may. Each one is called from exactly one entry
 * point in AdminApi.gs and writes exactly one audit row.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-08-16
 * Changes: 1.0.0 — initial version.
 ***********************************************/

/**
 * THE SHAPE EVERY ACTION SHARES
 *
 *   1. The caller is already authorized (requireAdmin_ ran in AdminApi.gs).
 *   2. The TARGET is re-resolved from the directory here, server-side. The
 *      browser names an account, but nothing the browser says about that account
 *      is believed — not its CAPID, not whether it is suspended, not whether its
 *      holder is an admin.
 *   3. admAssertMayActOn_() refuses admin targets for non-super-admin callers.
 *   4. A pure eligibility function decides yes or no, so preview and apply can
 *      never disagree about the answer.
 *   5. One audit row, whatever happened.
 *
 * Step 4 is why previews exist at all. Everything here is destructive in a way
 * that cannot be undone by running it again: a password reset locks out whoever
 * held the old one, and the temporary password replacing it is unrecoverable
 * once the page is closed.
 */

// ============================================================================
// WELCOME EMAIL RESEND
// ============================================================================

/**
 * Decides whether a member may be sent (or re-sent) a welcome email, and to
 * whom. Pure — no directory reads, no sends.
 *
 * PORTED from welcomeResendEligibility_() in
 * src/accounts-and-groups/WelcomeEmailResend.gs. test/AdminWebApp.test.js runs
 * both copies over the same cases and fails if they ever DECIDE differently —
 * verdict, reason slug or recipients. Only two sentences differ, both refusals
 * that told an operator in the script editor to "pass {force: true}"; on a web
 * page that is a checkbox, and the test pins the new wording by name so a third
 * divergence cannot arrive quietly. The two guards that matter most, in the
 * original's words:
 *
 *   - NEVER reset an account with login history. The member is using it; a reset
 *     locks them out of a working account to fix a cosmetic gap.
 *   - NEVER send to the account being reset. Credentials would land in the
 *     mailbox the new password is needed to open.
 *
 * @param {Object|null} member - from admBuildMemberRecord_(), or null
 * @param {Object|null} account - {email, suspended, archived, neverSignedIn}, or null
 * @param {Object} [opts]
 * @param {string[]} [opts.tenantDomains]
 * @param {boolean} [opts.force] - overrides the login-history and suspended
 *   guards ONLY. Never overrides "no account", "no CAPWATCH record" or "no
 *   off-tenant recipient" — those make the send impossible, not merely unwise.
 * @returns {{ok: boolean, reason: string, detail: string, recipients: string[]}}
 */
function admWelcomeEligibility_(member, account, opts) {
  const options = opts || {};
  const force = options.force === true;

  const excluded = (options.tenantDomains || [])
    .map(function (d) { return String(d || '').trim().toLowerCase().replace(/^@/, ''); })
    .filter(Boolean);

  const fail = function (reason, detail) {
    return { ok: false, reason: reason, detail: detail, recipients: [] };
  };

  if (!member) {
    return fail('no-capwatch-record',
      'No CAPWATCH member with that CAPID (expired, wrong wing, or a typo).');
  }
  if (!account || !account.email) {
    return fail('no-account',
      'No Workspace account carries that CAPID. This member needs provisioning, ' +
      'not a resend — run updateAllMembers() (senior accounts are gated on Level I).');
  }
  if (account.archived) {
    return fail('archived',
      'The account is archived. Restore it before sending credentials.');
  }

  const accountEmail = String(account.email).toLowerCase();
  const recipients = [member.email, member.secondaryEmail]
    .filter(Boolean)
    .map(function (e) { return String(e).trim(); })
    .filter(function (e) {
      if (e.toLowerCase() === accountEmail) return false;
      const at = e.lastIndexOf('@');
      if (at < 0) return false;
      return excluded.indexOf(e.slice(at + 1).toLowerCase()) === -1;
    })
    .filter(function (e, i, all) { return all.indexOf(e) === i; });

  if (recipients.length === 0) {
    return fail('no-off-tenant-recipient',
      'CAPWATCH has no contact address outside this tenant, so the credentials ' +
      'would be mailed to the very mailbox they unlock. Have the member add a ' +
      'personal address in eServices first.');
  }

  if (account.suspended && !force) {
    return fail('suspended',
      'The account is suspended — credentials would not work. Resolve the ' +
      'suspension first, or tick "force" if you are about to lift it.');
  }

  if (!account.neverSignedIn && !force) {
    return fail('already-signed-in',
      'The member has signed into this account, so they already have a working ' +
      'password. A resend RESETS it and would lock them out. Tick "force" only ' +
      'if they have genuinely lost access.');
  }

  return {
    ok: true,
    reason: force && (account.suspended || !account.neverSignedIn) ? 'forced' : 'eligible',
    detail: 'Eligible — the account exists, has never been signed into, and the ' +
      'member has an off-tenant address to receive credentials.',
    recipients: recipients
  };
}

/**
 * Every domain this tenant issues addresses on. An address on any of them is
 * unreachable by a member with no working password here, so none of them can
 * receive credentials. Ported from welcomeResendTenantDomains_().
 *
 * @returns {string[]}
 */
function admTenantDomains_() {
  return [ADMIN_CONFIG.DOMAIN, ADMIN_CONFIG.EMAIL_DOMAIN, ADMIN_CONFIG.SECONDARY_EMAIL_DOMAIN];
}

/**
 * READ-ONLY. What a resend would do, without doing it.
 *
 * @param {string} capid
 * @param {boolean} force
 * @returns {Object} the eligibility verdict, plus the account it would act on
 */
function admPreviewWelcomeResend_(capid, force) {
  const resolved = admResolveByCapid_(capid);
  const verdict = admWelcomeEligibility_(resolved.member, resolved.chosen, {
    tenantDomains: admTenantDomains_(),
    force: force === true
  });
  return {
    verdict: verdict,
    account: resolved.chosen ? resolved.chosen.email : '',
    accountCount: resolved.accounts.length,
    cc: ADMIN_CONFIG.SUPPORT_EMAIL || ''
  };
}

/**
 * Resets a member's password and mails them the welcome email carrying the new
 * one. For accounts created out-of-band, which never received one.
 *
 * DESTRUCTIVE: the account's existing password stops working immediately.
 *
 * @param {Object} actor - from requireAdmin_()
 * @param {string} capid
 * @param {boolean} force
 * @returns {{ok: boolean, reason: string, message: string, recipients: string[]}}
 */
function admResendWelcome_(actor, capid, force) {
  const resolved = admResolveByCapid_(capid);
  const verdict = admWelcomeEligibility_(resolved.member, resolved.chosen, {
    tenantDomains: admTenantDomains_(),
    force: force === true
  });

  if (!verdict.ok) {
    admAudit_(actor, 'resend-welcome', {
      capid: capid,
      target: resolved.chosen ? resolved.chosen.email : '',
      outcome: 'refused',
      detail: verdict.reason + ': ' + verdict.detail
    });
    return { ok: false, reason: verdict.reason, message: verdict.detail, recipients: [] };
  }

  // The target is an account, so the escalation rule applies even though the
  // caller reached it by CAPID rather than by naming an address.
  admAssertMayActOn_(actor, admDirectoryUser_(resolved.chosen.email));

  const member = resolved.member;
  const email = resolved.chosen.email;
  const password = admGenerateTempPassword_();

  try {
    admSetPassword_(email, password);
  } catch (err) {
    admAudit_(actor, 'resend-welcome', {
      capid: capid, target: email, outcome: 'failed',
      detail: 'Password reset failed, nothing was sent: ' + err.message
    });
    return {
      ok: false, reason: 'password-reset-failed',
      message: 'The password could not be reset, so nothing was sent: ' + err.message,
      recipients: []
    };
  }

  // The password is live from here on. A send failure now leaves the member
  // locked out of an account whose password just changed — so it is reported to
  // the admin standing at the phone, with the address to hand-deliver to, rather
  // than logged and forgotten.
  try {
    admSendWelcomeEmail_(member, email, password, verdict.recipients);
  } catch (err) {
    admAudit_(actor, 'resend-welcome', {
      capid: capid, target: email, outcome: 'failed',
      detail: 'PASSWORD WAS RESET but the email failed to send: ' + err.message
    });
    return {
      ok: false, reason: 'send-failed',
      message: 'The password was reset but the email did NOT send (' + err.message +
        '). The member cannot sign in until you give them credentials by hand — ' +
        'reset the password again below and read it to them.',
      recipients: verdict.recipients
    };
  }

  // Swallowed by design: the email is already gone, and a bookkeeping problem
  // must never turn a successful send into a failed one.
  try {
    admWelcomeLedgerRecordSent_(capid);
  } catch (err) {
    Logger.warn('Welcome email SENT, but the audit ledger could not be updated', {
      capsn: String(capid), email: email, errorMessage: err.message
    });
  }

  admAudit_(actor, 'resend-welcome', {
    capid: capid, target: email, outcome: 'ok',
    detail: 'Password reset and welcome email sent to ' + verdict.recipients.join(', ') +
      (verdict.reason === 'forced' ? ' (forced past a guard)' : '')
  });

  return {
    ok: true, reason: verdict.reason,
    message: 'Sent. ' + email + ' has a new temporary password, and credentials went to ' +
      verdict.recipients.join(', ') + '.',
    recipients: verdict.recipients
  };
}

// ============================================================================
// PASSWORD RESET
// ============================================================================

/**
 * Contact addresses that can safely receive credentials for an account: not the
 * account itself, and nothing on a domain this tenant issues.
 *
 * A near-duplicate of the filter inside admWelcomeEligibility_() rather than an
 * extraction from it, because that function is held byte-comparable to its
 * original in src/ and refactoring it would break the parity test that keeps the
 * welcome-email policy honest. This copy is small, and it is pinned by its own
 * test case.
 *
 * @param {Object|null} member
 * @param {string} accountEmail
 * @returns {string[]}
 */
function admOffTenantRecipients_(member, accountEmail) {
  if (!member) return [];
  const excluded = admTenantDomains_()
    .map(function (d) { return String(d || '').trim().toLowerCase().replace(/^@/, ''); })
    .filter(Boolean);
  const account = String(accountEmail || '').toLowerCase();

  return [member.email, member.secondaryEmail]
    .filter(Boolean)
    .map(function (e) { return String(e).trim(); })
    .filter(function (e) {
      if (e.toLowerCase() === account) return false;
      const at = e.lastIndexOf('@');
      if (at < 0) return false;
      return excluded.indexOf(e.slice(at + 1).toLowerCase()) === -1;
    })
    .filter(function (e, i, all) { return all.indexOf(e) === i; });
}

/**
 * Resets one account's password to a fresh temporary one.
 *
 * This is the plain reset a help desk needs on the phone, and it deliberately
 * carries FEWER guards than the welcome resend: an admin talking to a member who
 * has lost access is entitled to reset a working account, which is the case the
 * resend refuses. The password is returned so the admin can read it out, and is
 * additionally mailed to an off-tenant address when asked for.
 *
 * The escalation rule still applies, and a suspended account is still flagged —
 * a reset there is legal but pointless, and an admin should be told before they
 * read a password to someone who cannot use it.
 *
 * @param {Object} actor - from requireAdmin_()
 * @param {string} targetEmail - an account address, re-resolved here
 * @param {boolean} alsoEmail - mail the credentials to the member's off-tenant address
 * @returns {{ok: boolean, password: string, message: string, mailedTo: string[]}}
 */
function admResetPassword_(actor, targetEmail, alsoEmail) {
  const requested = String(targetEmail || '').trim().toLowerCase();
  const user = admDirectoryUser_(requested);

  if (!user) {
    admAudit_(actor, 'reset-password', {
      target: requested, outcome: 'refused', detail: 'No such account'
    });
    throw new Error('No account called ' + requested + ' exists in this directory.');
  }
  if (!admIsTenantAddress_(user.primaryEmail)) {
    // Cannot normally happen — the directory only holds our own domains — but the
    // address came from the browser, and this is the one place to say so.
    admAudit_(actor, 'reset-password', {
      target: user.primaryEmail, outcome: 'refused', detail: 'Not an address on a tenant domain'
    });
    throw new Error(user.primaryEmail + ' is not on a domain this app manages.');
  }
  admAssertMayActOn_(actor, user);

  const capids = admCapidsFromUser_(user);
  const member = capids.length === 1 ? admBuildMemberRecord_(capids[0]) : null;

  let recipients = [];
  if (alsoEmail) {
    recipients = admOffTenantRecipients_(member, user.primaryEmail);
    if (!recipients.length) {
      admAudit_(actor, 'reset-password', {
        target: user.primaryEmail, capid: capids[0] || '', outcome: 'refused',
        detail: 'Asked to mail credentials, but there is no off-tenant address to send to'
      });
      throw new Error('CAPWATCH has no contact address outside this tenant for this member, ' +
        'so the credentials would be mailed to the very mailbox they unlock. Reset without ' +
        'emailing and read the password out instead.');
    }
  }

  const password = admGenerateTempPassword_();
  try {
    admSetPassword_(user.primaryEmail, password);
  } catch (err) {
    admAudit_(actor, 'reset-password', {
      target: user.primaryEmail, capid: capids[0] || '', outcome: 'failed',
      detail: err.message
    });
    throw new Error('The password could not be reset: ' + err.message);
  }

  let mailNote = '';
  if (recipients.length) {
    try {
      admSendPasswordNotice_(member, user.primaryEmail, password, recipients);
    } catch (err) {
      // The reset already happened, so this cannot fail the whole action — but the
      // admin is holding the password and needs to know it did not go anywhere.
      Logger.error('Password was reset but the notice failed to send', {
        target: user.primaryEmail, errorMessage: err.message
      });
      mailNote = ' The email did NOT send (' + err.message + ') — read the password out instead.';
      recipients = [];
    }
  }

  admAudit_(actor, 'reset-password', {
    target: user.primaryEmail, capid: capids[0] || '', outcome: 'ok',
    // The password itself is never recorded anywhere. See AuditLog.gs.
    detail: 'Temporary password issued, change-at-next-login set' +
      (recipients.length ? '; mailed to ' + recipients.join(', ') : '; shown on screen only')
  });

  return {
    ok: true,
    password: password,
    mailedTo: recipients,
    message: user.primaryEmail + ' now has a temporary password and must change it at next ' +
      'sign-in.' + (recipients.length ? ' Credentials were mailed to ' + recipients.join(', ') + '.' : '') +
      mailNote +
      (user.suspended ? ' NOTE: this account is SUSPENDED, so the new password will not work ' +
        'until the suspension is lifted.' : '')
  };
}

/**
 * Mails a temporary password to a member's off-tenant address.
 *
 * Plain and short on purpose: the welcome email introduces an account to someone
 * who has never had one, and sending that to a member of ten years who has
 * merely forgotten a password reads as a system error on our end. `recipients`
 * has already been vetted by admOffTenantRecipients_().
 *
 * @param {Object|null} member
 * @param {string} email - the account the password is for
 * @param {string} password
 * @param {Array<string>} recipients
 * @returns {void}
 */
function admSendPasswordNotice_(member, email, password, recipients) {
  const name = member ? (member.rank ? member.rank + ' ' + member.lastName : member.firstName) : '';
  const org = ADMIN_CONFIG.ORG_LABEL;
  const support = ADMIN_CONFIG.SUPPORT_EMAIL;

  const body =
    '<p style="font:14px/1.6 system-ui,sans-serif">' +
    (name ? admEscape_(name) + ',' : 'Hello,') +
    '</p>' +
    '<p style="font:14px/1.6 system-ui,sans-serif">The password on your ' + admEscape_(org) +
    ' Workspace account has been reset at your request.</p>' +
    '<p style="font:14px/1.6 system-ui,sans-serif">' +
    '<strong>Account:</strong> ' + admEscape_(email) + '<br>' +
    '<strong>Temporary password:</strong> ' + admEscape_(password) + '</p>' +
    '<p style="font:14px/1.6 system-ui,sans-serif">You will be asked to choose a new password ' +
    'the first time you sign in at ' +
    '<a href="https://accounts.google.com/AccountChooser?hd=' + admEscape_(ADMIN_CONFIG.DOMAIN) + '">' +
    'accounts.google.com</a>. If you did not ask for this, tell us straight away' +
    (support ? ' at ' + admEscape_(support) : '') + '.</p>' +
    '<p style="font:14px/1.6 system-ui,sans-serif">' + admEscape_(org) + ' Information Technology</p>';

  const options = {
    to: recipients.join(','),
    subject: 'Your ' + org + ' Workspace password has been reset',
    htmlBody: body
  };
  if (support) options.cc = support;
  MailApp.sendEmail(options);
}

// ============================================================================
// GROUP MEMBERSHIP (2SV setup, and anything else on the managed list)
// ============================================================================

/**
 * Validates a group address the BROWSER supplied against the managed list.
 *
 * This is the boundary that makes group management safe to delegate: without it,
 * a help-desk admin could point the app at any group in the domain — including
 * one that grants Drive access or admin-adjacent privileges — simply by editing
 * a request. The list comes from Script Properties, so widening it is an Admin
 * console decision.
 *
 * @param {string} requested
 * @returns {string} the canonical group address
 * @throws {Error} when the group is not managed
 */
function admGroupFromRequest_(requested) {
  const wanted = String(requested || '').trim().toLowerCase();
  const managed = admManagedGroups_();

  if (!managed.length) {
    throw new Error('No groups are configured for this app. Set WEBAPP_2SV_SETUP_GROUP ' +
      'in Script Properties first.');
  }
  if (managed.indexOf(wanted) === -1) {
    Logger.warn('Refused a membership change against an unmanaged group', { group: wanted });
    throw new Error('This app may only change membership of the groups it is configured for.');
  }
  return wanted;
}

/**
 * Adds or removes one member from one managed group.
 *
 * Idempotent in both directions: adding someone already in the group, or
 * removing someone who is not, reports success rather than an error. A help desk
 * repeating an action because a page was stale should not be handed a failure
 * for a state that is already what they wanted.
 *
 * @param {Object} actor - from requireAdmin_()
 * @param {string} groupEmail - checked against the managed list
 * @param {string} targetEmail - an account address, re-resolved here
 * @param {boolean} add - true to add, false to remove
 * @returns {{ok: boolean, member: boolean, message: string}}
 */
function admSetGroupMembership_(actor, groupEmail, targetEmail, add) {
  const group = admGroupFromRequest_(groupEmail);
  const requested = String(targetEmail || '').trim().toLowerCase();
  const user = admDirectoryUser_(requested);
  const action = add ? 'group-add' : 'group-remove';

  if (!user) {
    admAudit_(actor, action, { target: requested, outcome: 'refused', detail: 'No such account' });
    throw new Error('No account called ' + requested + ' exists in this directory.');
  }
  admAssertMayActOn_(actor, user);

  const email = user.primaryEmail;
  const capids = admCapidsFromUser_(user);

  try {
    if (add) {
      AdminDirectory.Members.insert({ email: email, role: 'MEMBER' }, group);
    } else {
      AdminDirectory.Members.remove(group, email);
    }
  } catch (err) {
    const message = String(err.message || '');
    // Google reports "already exists" on a duplicate insert and 404/"Resource Not
    // Found" on removing a non-member. Both mean the group is already in the state
    // the admin asked for.
    const alreadyDone = add
      ? /duplicate|already exists|Member already/i.test(message)
      : /not found|notFound|Resource Not Found/i.test(message);

    if (!alreadyDone) {
      admAudit_(actor, action, {
        target: email, capid: capids[0] || '', outcome: 'failed',
        detail: group + ': ' + message
      });
      throw new Error('The group could not be changed: ' + message);
    }

    admAudit_(actor, action, {
      target: email, capid: capids[0] || '', outcome: 'ok',
      detail: group + ': already ' + (add ? 'a member' : 'not a member') + '; nothing to do'
    });
    return {
      ok: true, member: add,
      message: email + ' was already ' + (add ? 'in ' : 'out of ') + group + '.'
    };
  }

  admAudit_(actor, action, {
    target: email, capid: capids[0] || '', outcome: 'ok',
    detail: (add ? 'Added to ' : 'Removed from ') + group
  });
  return {
    ok: true, member: add,
    message: email + (add ? ' was added to ' : ' was removed from ') + group + '.'
  };
}
