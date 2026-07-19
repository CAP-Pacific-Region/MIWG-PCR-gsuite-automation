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
section('3. chooseAuthoritativeAccount_ — preference order');
{
  // Trigger case: two never-signed-in active accounts. Canonical = sam.roe.
  const triggerCase = [
    { email: 'roe.sam@example.org', suspended: false, created: '2025-11-24T00:00:00Z' },
    { email: 'sam.roe@example.org', suspended: false, created: '2026-01-23T00:00:00Z' }
  ];
  check('canonical first.last wins even though it is newer AND regardless of order',
    m.chooseAuthoritativeAccount_(triggerCase, 'sam.roe').email,
    'sam.roe@example.org');
  check('canonical wins even when it is the OLDER account',
    m.chooseAuthoritativeAccount_([
      { email: 'sam.roe@x.org', suspended: false, created: '2025-01-01T00:00:00Z' },
      { email: 'roe.sam@x.org', suspended: false, created: '2026-01-01T00:00:00Z' }
    ], 'sam.roe').email,
    'sam.roe@x.org');

  check('with no canonical hint, active beats suspended',
    m.chooseAuthoritativeAccount_([
      { email: 'a.b@x.org', suspended: true, created: '2026-05-01T00:00:00Z' },
      { email: 'b.a@x.org', suspended: false, created: '2024-01-01T00:00:00Z' }
    ], '').email,
    'b.a@x.org');

  check('among active, unsuffixed beats a .N collision',
    m.chooseAuthoritativeAccount_([
      { email: 'jane.doe.2@x.org', suspended: false, created: '2026-05-01T00:00:00Z' },
      { email: 'jane.doe@x.org', suspended: false, created: '2024-01-01T00:00:00Z' }
    ], '').email,
    'jane.doe@x.org');

  check('all else equal, newest wins',
    m.chooseAuthoritativeAccount_([
      { email: 'x.y@x.org', suspended: false, created: '2024-01-01T00:00:00Z' },
      { email: 'z.w@x.org', suspended: false, created: '2026-01-01T00:00:00Z' }
    ], '').email,
    'z.w@x.org');

  check('empty input', m.chooseAuthoritativeAccount_([], 'a.b'), null);
}

done();
