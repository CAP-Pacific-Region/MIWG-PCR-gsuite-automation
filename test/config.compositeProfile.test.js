/**
 * config.gs — the 'composite' profile (one tenant holding a wing's cadets AND seniors).
 *
 * The failure it exists to prevent: running a combined tenant on plain `seniors`,
 * where CADET is not an active member type. All identifiers are synthetic.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'config.gs');
const { section, check, done } = makeChecker();

function load(profile) {
  const props = {
    TENANT_PROFILE: profile, TENANT_WING: 'OR', TENANT_DOMAIN: 'example.org',
    TENANT_EMAIL_DOMAIN: '@example.org', TENANT_CAPWATCH_ORGID: '1'
  };
  return loadModule(MODULE, {
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null) }) },
    console: console
  }, ['PROFILE_', 'CONFIG']);
}

section('1. Composite provisions cadets and seniors');
{
  const c = load('composite');
  const t = c.CONFIG.MEMBER_TYPES.ACTIVE;
  check('CADET is active', t.includes('CADET'), true);
  check('SENIOR is active', t.includes('SENIOR'), true);
  check('INDEFINITE is active', t.includes('INDEFINITE'), true);
  check('cadets get full accounts (not cadet-lite)', c.PROFILE_.CADET_LITE, false);
}

section('2. A single tenant has no peer');
{
  const p = load('composite').PROFILE_;
  check('no transition role', p.TRANSITION_ROLE, '');
  check('no cross-tenant inbound', p.CROSS_TENANT.RUN_INBOUND, false);
  check('no cross-tenant parents', p.CROSS_TENANT.RUN_PARENTS, false);
}

section('2b. No Level I wait for new seniors');
{
  check('composite does not gate on Level I', load('composite').PROFILE_.REQUIRE_LEVEL_I_FOR_SENIORS, false);
  check('seniors still does', load('seniors').PROFILE_.REQUIRE_LEVEL_I_FOR_SENIORS, true);
}

section('3. Oregon labels and holding unit derive');
{
  const c = load('composite');
  check('abbreviation', c.CONFIG.WING_ABBREVIATION, 'ORWG');
  check('name', c.CONFIG.WING_NAME, 'Oregon Wing');
  check('holding unit is not California\'s', c.CONFIG.EXCLUDED_ORG_IDS, ['117']);
}

section('4. Plain seniors would have excluded cadets (why composite exists)');
{
  check('seniors omits CADET', load('seniors').CONFIG.MEMBER_TYPES.ACTIVE.includes('CADET'), false);
}

done();
