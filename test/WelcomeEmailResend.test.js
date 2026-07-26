/**
 * WelcomeEmailResend.gs — the guards on a welcome-email resend.
 *
 * A resend is a PASSWORD RESET plus a send (the original temp password is never
 * stored, so nothing else is possible). That makes every "no" in this policy a
 * protection against doing real damage to a live account, and each one is pinned
 * here. The two that matter most:
 *
 *   - an account with login history must NOT be reset (the member is using it)
 *   - credentials must NOT be mailed to the mailbox they unlock
 *
 * All names and addresses are synthetic.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'WelcomeEmailResend.gs');
const { section, check, done } = makeChecker();

const m = loadModule(MODULE, {
  Logger: makeLogger().logger,
  AdminDirectory: {},
  CONFIG: {},
  console: console
}, ['welcomeResendEligibility_']);

const TENANT = { tenantDomains: ['cawgcap.org', '@cawgcap.org', '@cawg.cap.gov'] };

// The shape this module was written for: an account created out-of-band, never
// signed into, with a personal address on file in CAPWATCH.
const member = () => ({
  firstName: 'Dana',
  lastName: 'Okonkwo',
  type: 'SENIOR',
  charter: 'PCR-XX-000',
  email: 'dana.okonkwo@example.com',
  secondaryEmail: null
});
const account = () => ({
  email: 'dana.okonkwo@cawgcap.org',
  suspended: false,
  archived: false,
  neverSignedIn: true
});

const verdict = (mem, acc, opts) =>
  m.welcomeResendEligibility_(mem, acc, Object.assign({}, TENANT, opts || {}));

// ---------------------------------------------------------------------------
section('1. The case this exists for — out-of-band account, never welcomed');
{
  const v = verdict(member(), account());
  check('eligible', v.ok, true);
  check('reason slug', v.reason, 'eligible');
  check('mails the personal address, not the tenant one', v.recipients, ['dana.okonkwo@example.com']);
}

// ---------------------------------------------------------------------------
section('2. Never reset an account the member is actually using');
{
  const used = Object.assign(account(), { neverSignedIn: false });
  const v = verdict(member(), used);
  check('blocked', v.ok, false);
  check('reason', v.reason, 'already-signed-in');
  check('no recipients leak out of a refusal', v.recipients, []);

  const forced = verdict(member(), used, { force: true });
  check('force overrides it (member has genuinely lost access)', forced.ok, true);
  check('and says so, so the log records that a guard was bypassed', forced.reason, 'forced');
}

// ---------------------------------------------------------------------------
section('3. Never mail credentials to the mailbox they unlock');
{
  // The silent failure: CAPWATCH lists only the member's CAP address, so the
  // welcome email lands in the account the new password is needed to open.
  const onlyTenantAddress = Object.assign(member(), {
    email: 'dana.okonkwo@cawgcap.org',
    secondaryEmail: null
  });
  const v = verdict(onlyTenantAddress, account());
  check('blocked', v.ok, false);
  check('reason', v.reason, 'no-off-tenant-recipient');

  // The secondary domain is the same tenant wearing a different name.
  check('an address on the secondary domain is equally unreachable',
    verdict(Object.assign(member(), {
      email: 'd.okonkwo@cawg.cap.gov',
      secondaryEmail: null
    }), account()).reason,
    'no-off-tenant-recipient');

  // A tenant address alongside a personal one is fine — drop the former, keep the latter.
  check('tenant address dropped, personal address kept',
    verdict(Object.assign(member(), {
      email: 'dana.okonkwo@cawgcap.org',
      secondaryEmail: 'dana@example.net'
    }), account()).recipients,
    ['dana@example.net']);

  check('force does NOT override this — the send would be useless, not merely unwise',
    verdict(onlyTenantAddress, account(), { force: true }).ok,
    false);

  check('no CAPWATCH address at all is the same refusal',
    verdict(Object.assign(member(), { email: null, secondaryEmail: null }), account()).reason,
    'no-off-tenant-recipient');
}

// ---------------------------------------------------------------------------
section('4. Recipient list hygiene');
{
  check('both personal addresses are used when both are on file',
    verdict(Object.assign(member(), {
      email: 'dana@example.com',
      secondaryEmail: 'dana.alt@example.net'
    }), account()).recipients,
    ['dana@example.com', 'dana.alt@example.net']);

  check('the same address listed twice is mailed once',
    verdict(Object.assign(member(), {
      email: 'dana@example.com',
      secondaryEmail: 'dana@example.com'
    }), account()).recipients,
    ['dana@example.com']);

  check('surrounding whitespace does not create a second recipient',
    verdict(Object.assign(member(), {
      email: '  dana@example.com  ',
      secondaryEmail: 'dana@example.com'
    }), account()).recipients,
    ['dana@example.com']);

  check('the account address is matched case-insensitively',
    verdict(Object.assign(member(), {
      email: 'Dana.Okonkwo@CAWGCAP.org',
      secondaryEmail: null
    }), account()).reason,
    'no-off-tenant-recipient');

  check('a malformed address with no @ is not a recipient',
    verdict(Object.assign(member(), {
      email: 'not-an-address',
      secondaryEmail: null
    }), account()).reason,
    'no-off-tenant-recipient');
}

// ---------------------------------------------------------------------------
section('5. Account state');
{
  check('suspended is blocked — credentials would not work',
    verdict(member(), Object.assign(account(), { suspended: true })).reason,
    'suspended');
  check('...but can be forced, for an admin about to lift the suspension',
    verdict(member(), Object.assign(account(), { suspended: true }), { force: true }).ok,
    true);

  check('archived is blocked outright',
    verdict(member(), Object.assign(account(), { archived: true })).reason,
    'archived');
  check('force does not resurrect an archived account',
    verdict(member(), Object.assign(account(), { archived: true }), { force: true }).ok,
    false);
}

// ---------------------------------------------------------------------------
section('6. Nothing to act on');
{
  // No account means this member needs PROVISIONING, not a resend — and for a
  // senior that may be the Level I gate doing its job. Saying so is the point.
  check('no account', verdict(member(), null).reason, 'no-account');
  check('no account, even forced', verdict(member(), null, { force: true }).ok, false);
  check('an account object with no email is no account',
    verdict(member(), { email: '' }).reason, 'no-account');

  check('CAPID not in CAPWATCH', verdict(null, account()).reason, 'no-capwatch-record');
  check('missing record outranks a missing account',
    verdict(null, null).reason, 'no-capwatch-record');
}

// ---------------------------------------------------------------------------
section('7. Refusals never carry recipients');
{
  const refusals = [
    verdict(null, account()),
    verdict(member(), null),
    verdict(member(), Object.assign(account(), { archived: true })),
    verdict(member(), Object.assign(account(), { suspended: true })),
    verdict(member(), Object.assign(account(), { neverSignedIn: false }))
  ];
  check('every refusal returns ok:false with an empty recipient list',
    refusals.every(r => r.ok === false && r.recipients.length === 0),
    true);
  check('every refusal explains itself to a human',
    refusals.every(r => typeof r.detail === 'string' && r.detail.length > 20),
    true);
}

done();
