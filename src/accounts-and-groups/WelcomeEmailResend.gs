/**
 * -------------------------------------------------------------------------
 * Version: 1.0.0
 * Date: 2026-07-26
 * Authors: Michigan Wing (MIWG) — Extended and Maintained by Lt Col Noel Luneau
 * Contributors: Maj Isaac Wilson IV, California Wing (1.0.0)
 * Changes: 1.0.0 — new module. Sends the welcome email to a member who already
 *   has an account, for the case provisioning cannot cover: an account created
 *   OUT-OF-BAND (Admin console / GAM / another admin) never passes through the
 *   insert branch of addOrUpdateUser(), which holds the only sendWelcomeEmail()
 *   call in the codebase — so that member never receives credentials and the
 *   next sync, seeing an existing account, simply updates it. Nothing detects
 *   or repairs this. previewWelcomeEmailResend() reports; resendWelcomeEmail()
 *   resets the password and sends.
 * -------------------------------------------------------------------------
 *
 * WHY A RESEND EXISTS AT ALL
 *
 * The welcome email is inseparable from account creation: it carries the temp
 * password generated at insert time (UpdateMembers.gs, generateTempPassword_),
 * which is never stored anywhere. There is therefore no way to "re-send" the
 * original message — the password in it is unrecoverable. A resend is really a
 * password RESET plus a fresh send, and that is what this module does.
 *
 * That makes it a destructive operation on a live account, so the decision of
 * whether a given member may be resent is separated out into a pure function
 * (welcomeResendEligibility_) and guarded hard. Two guards matter most:
 *
 *   - NEVER reset an account with login history. The member is using it; a
 *     reset locks them out of a working account to fix a cosmetic gap. Blocked
 *     unless the caller explicitly forces it.
 *   - NEVER send to the account being reset. If a member's only CAPWATCH
 *     address is their own tenant address, the credentials land in the mailbox
 *     the new password is needed to open. Blocked outright — this is the silent
 *     failure that makes a "sent" welcome email useless.
 *
 * Both are decided by welcomeResendEligibility_ and covered in
 * test/WelcomeEmailResend.test.js.
 */

/**
 * Decides whether a member may be sent (or re-sent) a welcome email, and to
 * whom. Pure — no directory reads, no sends — so the policy is testable
 * without a tenant. All rules are evaluated here rather than at the call site
 * so that the preview and the live path can never disagree about the answer.
 *
 * @param {Object|null} member - Member object from getMembers(), or null if the
 *   CAPID is not in CAPWATCH
 * @param {Object|null} account - The chosen account
 *   ({email, suspended, archived, neverSignedIn}), or null if the member has none
 * @param {Object} [opts]
 * @param {string[]} [opts.tenantDomains] - The tenant's own domains. An address
 *   on one of these cannot receive credentials for an account on the same
 *   tenant, so it never counts as a recipient.
 * @param {boolean} [opts.force] - Override the login-history and suspended
 *   guards. Does NOT override "no account", "no CAPWATCH record" or
 *   "no off-tenant recipient" — those make the send impossible, not just unwise.
 * @returns {{ok: boolean, reason: string, detail: string, recipients: string[]}}
 *   `reason` is a stable slug for logging; `detail` is the sentence for a human.
 */
function welcomeResendEligibility_(member, account, opts) {
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

  // The address being reset can never be told about the reset. Anything on a
  // tenant domain is assumed unreachable for the same reason — the member has
  // no working credentials for any account on this tenant.
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
      'suspension first, or pass {force: true} if you are about to lift it.');
  }

  if (!account.neverSignedIn && !force) {
    return fail('already-signed-in',
      'The member has signed into this account, so they already have a working ' +
      'password. A resend RESETS it and would lock them out. Pass {force: true} ' +
      'only if they have genuinely lost access.');
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
 * unreachable by a member who has no working password here, so none of them can
 * receive credentials. DOMAIN and EMAIL_DOMAIN are the same domain in the '@'-less
 * and '@'-prefixed forms on every tenant configured so far; both are listed rather
 * than assumed equal, and the normaliser in welcomeResendEligibility_ folds them.
 *
 * @returns {string[]}
 */
function welcomeResendTenantDomains_() {
  return [CONFIG.DOMAIN, CONFIG.EMAIL_DOMAIN, CONFIG.SECONDARY_EMAIL_DOMAIN];
}

/**
 * Resolves a CAPID to the member record and the account that carries it, using
 * the same live-directory lookup and same authoritative-account ranking as the
 * duplicate-create guard. Sharing that path is deliberate: a resend must land on
 * the account provisioning maintains, not on a derived first.last address (which
 * may be a dead twin, or may not exist at all for an out-of-band account).
 *
 * @param {string|number} capid
 * @returns {{member: Object|null, account: Object|null, accounts: Array}}
 */
function welcomeResendResolve_(capid) {
  const wanted = String(capid).trim();

  const members = getMembers();
  const member = members[wanted] || null;

  const accounts = findExistingAccountsByCapid_(wanted);
  const canonicalLocal = member
    ? (member.firstName + '.' + member.lastName).toLowerCase().replace(/\s+/g, '')
    : '';
  const account = chooseAuthoritativeAccount_(accounts, canonicalLocal, CONFIG.EMAIL_DOMAIN);

  return { member: member, account: account, accounts: accounts };
}

/**
 * READ-ONLY. Reports whether a member would be sent a welcome email, to which
 * addresses, and on which account — changing nothing. Run this first.
 *
 * Output goes to the Execution log. Addresses are printed because the operator
 * needs to confirm them before credentials are mailed anywhere.
 *
 * @param {string|number} capid
 * @param {Object} [opts] - Same options as resendWelcomeEmail(); pass the same
 *   ones here so the preview answers the question you are actually about to ask.
 * @returns {Object} The eligibility result, for programmatic callers
 */
function previewWelcomeEmailResend(capid, opts) {
  const resolved = welcomeResendResolve_(capid);
  const verdict = welcomeResendEligibility_(resolved.member, resolved.account, {
    tenantDomains: welcomeResendTenantDomains_(),
    force: opts && opts.force === true
  });

  console.log('Welcome email resend — PREVIEW (nothing was changed or sent)');
  console.log('  CAPID:    ' + capid);
  console.log('  Member:   ' + (resolved.member
    ? resolved.member.firstName + ' ' + resolved.member.lastName +
      ' (' + (resolved.member.type || 'unknown type') + ', ' +
      (resolved.member.charter || 'no charter') + ')'
    : 'NOT FOUND in CAPWATCH'));

  if (resolved.accounts.length > 1) {
    console.log('  NOTE:     ' + resolved.accounts.length + ' accounts carry this CAPID — ' +
      'the authoritative one is used. Run scanDuplicateAccountsByCapid() if that is unexpected.');
    resolved.accounts.forEach(function (a) {
      console.log('              ' + a.email +
        (a.suspended ? ' [suspended]' : '') +
        (a.neverSignedIn ? ' [never signed in]' : ' [has login history]'));
    });
  }

  console.log('  Account:  ' + (resolved.account
    ? resolved.account.email +
      (resolved.account.suspended ? ' [suspended]' : '') +
      (resolved.account.neverSignedIn ? ' [never signed in]' : ' [HAS LOGIN HISTORY]')
    : 'none'));
  console.log('  Verdict:  ' + (verdict.ok ? 'WOULD SEND' : 'BLOCKED (' + verdict.reason + ')'));
  console.log('  Reason:   ' + verdict.detail);
  if (verdict.ok) {
    console.log('  Would mail to: ' + verdict.recipients.join(', '));
    console.log('  Would CC:      ' + ITSUPPORT_EMAIL);
    console.log('  This WILL reset the account password when you run resendWelcomeEmail(' + capid + ').');
  }

  return verdict;
}

/**
 * Resets a member's password and sends them the welcome email carrying the new
 * one. For accounts created outside provisioning, which never got one.
 *
 * DESTRUCTIVE: the account's existing password stops working immediately. Run
 * previewWelcomeEmailResend() first. The guards in welcomeResendEligibility_
 * refuse the cases where that would do harm rather than good.
 *
 * @param {string|number} capid - The member's CAPID
 * @param {Object} [opts]
 * @param {boolean} [opts.force=false] - Proceed despite login history or a
 *   suspension. Use only when the member has genuinely lost access.
 * @returns {{sent: boolean, reason: string, email: string|null, recipients: string[]}}
 */
function resendWelcomeEmail(capid, opts) {
  const options = opts || {};
  const resolved = welcomeResendResolve_(capid);
  const verdict = welcomeResendEligibility_(resolved.member, resolved.account, {
    tenantDomains: welcomeResendTenantDomains_(),
    force: options.force === true
  });

  if (!verdict.ok) {
    Logger.warn('Welcome email resend refused', {
      capsn: String(capid),
      reason: verdict.reason,
      detail: verdict.detail,
      email: resolved.account ? resolved.account.email : null
    });
    console.log('REFUSED (' + verdict.reason + '): ' + verdict.detail);
    return { sent: false, reason: verdict.reason, email: resolved.account ? resolved.account.email : null, recipients: [] };
  }

  const member = resolved.member;
  const email = resolved.account.email;

  // Same generator provisioning uses, so the password satisfies the same
  // complexity policy and is not derivable from public member data.
  const generatedPassword = generateTempPassword_();

  try {
    executeWithRetry(function () {
      return AdminDirectory.Users.update({
        password: generatedPassword,
        changePasswordAtNextLogin: true
      }, email);
    });
  } catch (e) {
    Logger.error('Welcome email resend — password reset failed, nothing sent', {
      capsn: String(capid),
      email: email,
      errorMessage: e.message,
      errorCode: e.details?.code
    });
    return { sent: false, reason: 'password-reset-failed', email: email, recipients: [] };
  }

  // The password is live from here on. A send failure now leaves the member
  // locked out of an account whose password just changed, so it is logged as an
  // ERROR with the address to hand-deliver to — not swallowed as a warning.
  try {
    sendWelcomeEmail(member, email, generatedPassword);
  } catch (e) {
    Logger.error('Welcome email resend — PASSWORD WAS RESET but the email failed to send. ' +
      'The member cannot sign in until they are given credentials by hand.', {
      capsn: String(capid),
      email: email,
      recipients: verdict.recipients,
      errorMessage: e.message,
      errorCode: e.details?.code
    });
    return { sent: false, reason: 'send-failed', email: email, recipients: verdict.recipients };
  }

  Logger.info('Welcome email resent', {
    capsn: String(capid),
    email: email,
    recipientCount: verdict.recipients.length,
    forced: verdict.reason === 'forced'
  });
  console.log('Sent. ' + email + ' has a new temporary password; credentials mailed to ' +
    verdict.recipients.join(', ') + ' (CC ' + ITSUPPORT_EMAIL + ').');

  return { sent: true, reason: verdict.reason, email: email, recipients: verdict.recipients };
}
