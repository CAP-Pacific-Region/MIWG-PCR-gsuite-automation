/**
 * DuplicateAccountScan.gs — the pure classification helpers the read-only scanner
 * relies on: pulling a CAPID off an account no matter which carrier holds it,
 * recognizing "never signed in" (Google returns the Unix epoch, not absent), and
 * telling a reversed-name duplicate (last.first vs first.last) apart from a .N
 * collision. The reversed-name trigger case is pinned directly (synthetic names).
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'DuplicateAccountScan.gs');
const { section, check, done } = makeChecker();

const m = loadModule(MODULE, { Logger: makeLogger().logger, AdminDirectory: {} }, [
  'extractCapidsFromUser_', 'localpartTokens_', 'classifyLocalpartPair_', 'neverLoggedIn_',
  'capidCarriersForUser_'
]);

// ---------------------------------------------------------------------------
section('1. extractCapidsFromUser_ — CAPID from any carrier, de-duplicated');
{
  check('organization externalId (what the code writes)',
    m.extractCapidsFromUser_({ externalIds: [{ type: 'organization', value: '123456' }] }),
    ['123456']);
  check('top-level employeeId (legacy/synced account)',
    m.extractCapidsFromUser_({ employeeId: '654321' }),
    ['654321']);
  check('same CAPID via two carriers counts once',
    m.extractCapidsFromUser_({ employeeId: '123456', externalIds: [{ type: 'organization', value: '123456' }] }),
    ['123456']);
  check('no CAPID anywhere',
    m.extractCapidsFromUser_({ externalIds: [{ type: 'custom', value: 'not-a-capid' }] }),
    []);
  check('non-CAPID-shaped externalId ignored',
    m.extractCapidsFromUser_({ externalIds: [{ type: 'organization', value: '42' }] }),
    []);
}

// ---------------------------------------------------------------------------
section('2. neverLoggedIn_ — epoch and missing both mean never');
{
  check('missing', m.neverLoggedIn_(null), true);
  check('unix epoch string', m.neverLoggedIn_('1970-01-01T00:00:00.000Z'), true);
  check('real login', m.neverLoggedIn_('2026-01-15T10:00:00.000Z'), false);
}

// ---------------------------------------------------------------------------
section('3. localpartTokens_ — split on ".", strip trailing numeric suffix');
{
  check('first.last', m.localpartTokens_('sam.roe@example.org'),
    { tokens: ['sam', 'roe'], suffix: null });
  check('collision .2', m.localpartTokens_('sam.roe.2@example.org'),
    { tokens: ['sam', 'roe'], suffix: '2' });
}

// ---------------------------------------------------------------------------
section('4. classifyLocalpartPair_ — the trigger case and its cousins');
{
  check('reversed: roe.sam vs sam.roe is reversed',
    m.classifyLocalpartPair_('roe.sam@example.org', 'sam.roe@example.org'),
    'reversed');
  check('collision differs only by trailing .N',
    m.classifyLocalpartPair_('sam.roe.2@example.org', 'sam.roe@example.org'),
    'collision');
  check('unrelated localparts are other',
    m.classifyLocalpartPair_('jane.doe@x.org', 'sam.roe@x.org'),
    'other');
}

// ---------------------------------------------------------------------------
section('5. capidCarriersForUser_ — which field holds the CAPID decides guard visibility');
{
  check('organization externalId (what the guard query can find)',
    m.capidCarriersForUser_({ externalIds: [{ type: 'organization', value: '123456' }] }, '123456'),
    ['externalId:organization']);
  check('employeeId ONLY — invisible to an externalId= query',
    m.capidCarriersForUser_({ employeeId: '123456' }, '123456'),
    ['employeeId']);
  check('both carriers reported',
    m.capidCarriersForUser_({ employeeId: '123456', externalIds: [{ type: 'organization', value: '123456' }] }, '123456'),
    ['externalId:organization', 'employeeId']);
  check('retired marker is reported with its customType',
    m.capidCarriersForUser_({ externalIds: [{ type: 'custom', customType: 'duplicate_retired_capid', value: '123456' }] }, '123456'),
    ['externalId:custom/duplicate_retired_capid']);
  check('a different CAPID on the account is not reported',
    m.capidCarriersForUser_({ externalIds: [{ type: 'organization', value: '999999' }] }, '123456'),
    []);
}

done();
