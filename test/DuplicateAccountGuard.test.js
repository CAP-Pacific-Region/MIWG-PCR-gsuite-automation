/**
 * DuplicateAccountGuard.gs — the pure decision logic behind the duplicate-create
 * guard and the cleanup: which CAPID carriers count as a live account (retired
 * markers must NOT), and which of several accounts sharing a CAPID is authoritative.
 * The trigger case (reversed last.first orphan + canonical first.last) is pinned
 * directly, using synthetic names.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'DuplicateAccountGuard.gs');
const { section, check, done } = makeChecker();

const m = loadModule(MODULE, { Logger: makeLogger().logger, AdminDirectory: {} }, [
  'provisioningCapidsFromUser_', 'isRetiredCapidExternalId_', 'chooseAuthoritativeAccount_'
]);

// ---------------------------------------------------------------------------
section('1. provisioningCapidsFromUser_ — live carriers only, retired marker ignored');
{
  check('organization externalId counts',
    m.provisioningCapidsFromUser_({ externalIds: [{ type: 'organization', value: '123456' }] }),
    ['123456']);
  check('top-level employeeId counts',
    m.provisioningCapidsFromUser_({ employeeId: '654321' }),
    ['654321']);
  check('RETIRED marker is NOT a live carrier (prevents resurrecting a retired orphan)',
    m.provisioningCapidsFromUser_({ externalIds: [{ type: 'custom', customType: 'duplicate_retired_capid', value: '123456' }] }),
    []);
  check('retired marker alongside a live org id still yields the CAPID once',
    m.provisioningCapidsFromUser_({ externalIds: [
      { type: 'custom', customType: 'duplicate_retired_capid', value: '123456' },
      { type: 'organization', value: '123456' }
    ] }),
    ['123456']);
}

// ---------------------------------------------------------------------------
section('2. isRetiredCapidExternalId_');
{
  check('the marker', m.isRetiredCapidExternalId_({ type: 'custom', customType: 'duplicate_retired_capid', value: '1' }), true);
  check('a normal custom id is not the marker', m.isRetiredCapidExternalId_({ type: 'custom', customType: 'other', value: '1' }), false);
  check('an organization id is not the marker', m.isRetiredCapidExternalId_({ type: 'organization', value: '1' }), false);
}

// ---------------------------------------------------------------------------
section('3. chooseAuthoritativeAccount_ — LOGIN HISTORY outranks the canonical name');
{
  // The shape actually found on the cadets tenant: the older, oddly-named account
  // (a .N collision suffix, or a hyphen the derived address drops) is the one the
  // member logs into; the newer canonically-named twin has never been signed into.
  // Preferring the canonical name here would retire the account in active use.
  const realWorld = [
    { email: 'sam.roe.2@example.org', suspended: false, created: '2025-11-24T00:00:00Z', lastLogin: '2026-07-15T00:00:00Z', neverSignedIn: false },
    { email: 'sam.roe@example.org', suspended: false, created: '2026-01-23T00:00:00Z', neverSignedIn: true }
  ];
  check('the signed-into account wins over the canonical-but-dead twin',
    m.chooseAuthoritativeAccount_(realWorld, 'sam.roe').email,
    'sam.roe.2@example.org');

  check('hyphen-drift variant: in-use hyphenated account beats the dead canonical one',
    m.chooseAuthoritativeAccount_([
      { email: 'lu-ann.fernandez@example.org', suspended: false, created: '2025-11-24T00:00:00Z', lastLogin: '2026-07-08T00:00:00Z', neverSignedIn: false },
      { email: 'luann.fernandez@example.org', suspended: false, created: '2026-01-23T00:00:00Z', neverSignedIn: true }
    ], 'luann.fernandez').email,
    'lu-ann.fernandez@example.org');

  check('login history beats even active-vs-suspended',
    m.chooseAuthoritativeAccount_([
      { email: 'used.but.suspended@x.org', suspended: true, created: '2025-01-01T00:00:00Z', lastLogin: '2026-06-01T00:00:00Z', neverSignedIn: false },
      { email: 'fresh.active@x.org', suspended: false, created: '2026-01-01T00:00:00Z', neverSignedIn: true }
    ], '').email,
    'used.but.suspended@x.org');

  // BOTH accounts have login history — the case a has-ever-signed-in boolean tied,
  // letting "newest created" pick the STALE account and mark the actively-used one
  // for retirement. Recency must decide.
  check('both signed in: the RECENTLY used account wins over a stale one',
    m.chooseAuthoritativeAccount_([
      { email: 'stale.newer@x.org', suspended: false, created: '2026-01-23T00:00:00Z', lastLogin: '2026-01-26T00:00:00Z', neverSignedIn: false },
      { email: 'active.older@x.org', suspended: false, created: '2025-11-24T00:00:00Z', lastLogin: '2026-07-15T00:00:00Z', neverSignedIn: false }
    ], '').email,
    'active.older@x.org');

  check('both signed in: recency beats the canonical name too',
    m.chooseAuthoritativeAccount_([
      { email: 'sam.roe@x.org', suspended: false, created: '2026-01-23T00:00:00Z', lastLogin: '2026-01-26T00:00:00Z', neverSignedIn: false },
      { email: 'roe.sam@x.org', suspended: false, created: '2025-11-24T00:00:00Z', lastLogin: '2026-07-15T00:00:00Z', neverSignedIn: false }
    ], 'sam.roe').email,
    'roe.sam@x.org');

  // With login history equal, the older tie-breakers still apply.
  check('both never signed in: canonical first.last wins',
    m.chooseAuthoritativeAccount_([
      { email: 'roe.sam@x.org', suspended: false, created: '2025-11-24T00:00:00Z', neverSignedIn: true },
      { email: 'sam.roe@x.org', suspended: false, created: '2026-01-23T00:00:00Z', neverSignedIn: true }
    ], 'sam.roe').email,
    'sam.roe@x.org');

  check('both never signed in: active beats suspended',
    m.chooseAuthoritativeAccount_([
      { email: 'a.b@x.org', suspended: true, created: '2026-05-01T00:00:00Z', neverSignedIn: true },
      { email: 'b.a@x.org', suspended: false, created: '2024-01-01T00:00:00Z', neverSignedIn: true }
    ], '').email,
    'b.a@x.org');

  check('both never signed in, no canonical hint: unsuffixed beats a .N collision',
    m.chooseAuthoritativeAccount_([
      { email: 'jane.doe.2@x.org', suspended: false, created: '2026-05-01T00:00:00Z', neverSignedIn: true },
      { email: 'jane.doe@x.org', suspended: false, created: '2024-01-01T00:00:00Z', neverSignedIn: true }
    ], '').email,
    'jane.doe@x.org');

  check('all else equal, newest wins',
    m.chooseAuthoritativeAccount_([
      { email: 'x.y@x.org', suspended: false, created: '2024-01-01T00:00:00Z', neverSignedIn: true },
      { email: 'z.w@x.org', suspended: false, created: '2026-01-01T00:00:00Z', neverSignedIn: true }
    ], '').email,
    'z.w@x.org');

  check('empty input', m.chooseAuthoritativeAccount_([], 'a.b'), null);
}

done();
