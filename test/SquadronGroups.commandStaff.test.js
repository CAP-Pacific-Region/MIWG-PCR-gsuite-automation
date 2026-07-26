/**
 * SquadronGroups.gs — which command-staff DLs a unit gets, and which ones the
 * cleanup can take away.
 *
 * CAP establishes a plain Deputy Commander only at senior units; cadet and
 * composite units have Deputy Commander for Cadets / for Seniors in its place. The
 * module used to hand every unit type a ca###.deputy-commander DL anyway, so most
 * of the wing carried a group no CAPWATCH duty can ever fill — created, listed in
 * the GAL, and silently accepting mail for nobody.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'squadron-groups', 'SquadronGroups.gs');
const { section, check, done } = makeChecker();

const CONFIG = { WING: 'CA', EMAIL_DOMAIN: '@example.org' };

const { logger } = makeLogger();
const m = loadModule(MODULE, {
  Logger: logger,
  CONFIG: CONFIG,
  // PROFILE_ absent on purpose: the fallback toggles have every list enabled, so
  // what these assertions see is the type rule alone, not a tenant's toggles.
  parseFile: () => []
}, [
  'getDistributionListsForSquadron',
  'getUnnecessaryDistributionLists',
  'classifyUnitCommandStaffKind_',
  'COMMAND_STAFF_SUFFIXES_BY_UNIT_KIND_'
]);

/** A squadron object with the type set explicitly, so no Organization.txt is read. */
function unit(type, unitNumber = '080') {
  return { orgid: '2500', charter: `PCR-CA-${unitNumber}`, unit: unitNumber, wing: 'CA', scope: 'UNIT', type: type };
}

const cadet = { type: 'CADET' };
const senior = { type: 'SENIOR' };

function suffixes(squadron, squadronMembers) {
  return m.getDistributionListsForSquadron(squadron, squadronMembers).map(l => l.suffix);
}

// ---------------------------------------------------------------------------
section('1. A plain Deputy Commander exists only at senior units');
{
  check('senior squadron', suffixes(unit('SENIOR')),
    ['all', 'commander', 'deputy-commander']);

  check('composite squadron: the two deputies, not the plain one',
    suffixes(unit('COMPOSITE')),
    ['all', 'cadets', 'seniors', 'parents', 'commander', 'deputy-commander-cadets', 'deputy-commander-seniors']);

  check('cadet squadron: same command staff as composite',
    suffixes(unit('CADET')),
    ['all', 'cadets', 'seniors', 'parents', 'commander', 'deputy-commander-cadets', 'deputy-commander-seniors']);

  check('no unit type gets both flavors',
    Object.keys(m.COMMAND_STAFF_SUFFIXES_BY_UNIT_KIND_).every(kind => {
      const s = m.COMMAND_STAFF_SUFFIXES_BY_UNIT_KIND_[kind];
      return !(s.includes('deputy-commander') && s.includes('deputy-commander-cadets'));
    }), true);
}

// ---------------------------------------------------------------------------
section('2. A flight is classified by who is actually in it');
{
  check('cadets present', suffixes(unit('FLIGHT'), [cadet, senior]),
    ['all', 'cadets', 'seniors', 'parents', 'commander', 'deputy-commander-cadets', 'deputy-commander-seniors']);

  check('seniors only', suffixes(unit('FLIGHT'), [senior]),
    ['all', 'commander', 'deputy-commander']);

  check('cadets only', suffixes(unit('FLIGHT'), [cadet]),
    ['all', 'cadets', 'seniors', 'parents', 'commander', 'deputy-commander-cadets', 'deputy-commander-seniors']);
}

// ---------------------------------------------------------------------------
section('3. Undetermined units get the one list that is always right');
{
  check('flight with no members: rosters, but Commander only',
    suffixes(unit('FLIGHT'), []),
    ['all', 'cadets', 'seniors', 'parents', 'commander']);

  check('unknown type',
    suffixes(unit('SOMETHING-NEW')),
    ['all', 'cadets', 'seniors', 'parents', 'commander']);

  check('kind classification', [
    m.classifyUnitCommandStaffKind_('SENIOR', null),
    m.classifyUnitCommandStaffKind_('COMPOSITE', null),
    m.classifyUnitCommandStaffKind_('CADET', null),
    m.classifyUnitCommandStaffKind_('FLIGHT', null),
    m.classifyUnitCommandStaffKind_('', null)
  ], ['SENIOR', 'CADET_OR_COMPOSITE', 'CADET_OR_COMPOSITE', 'UNDETERMINED', 'UNDETERMINED']);
}

// ---------------------------------------------------------------------------
section('4. Cleanup offers exactly the command-staff DLs the type forbids');
{
  const cadetDeletes = m.getUnnecessaryDistributionLists(unit('CADET'), 'ca080');
  check('cadet unit: the plain deputy-commander',
    cadetDeletes, ['ca080.deputy-commander@example.org']);

  const compositeDeletes = m.getUnnecessaryDistributionLists(unit('COMPOSITE'), 'ca080');
  check('composite unit: the plain deputy-commander',
    compositeDeletes, ['ca080.deputy-commander@example.org']);

  const seniorDeletes = m.getUnnecessaryDistributionLists(unit('SENIOR'), 'ca080');
  check('senior unit: the two cadet/senior deputies, plus its usual roster lists',
    seniorDeletes, [
      'ca080.deputy-commander-cadets@example.org',
      'ca080.deputy-commander-seniors@example.org',
      'ca080.cadets@example.org',
      'ca080.seniors@example.org',
      'ca080.parents@example.org'
    ]);

  check('a flight is left alone: its kind cannot be resolved without members',
    m.getUnnecessaryDistributionLists(unit('FLIGHT'), 'ca080'), []);

  check('the Commander DL is never a deletion candidate',
    []
      .concat(cadetDeletes, compositeDeletes, seniorDeletes)
      .some(e => e.endsWith('.commander@example.org')), false);
}

// ---------------------------------------------------------------------------
section('5. Special and non-unit orgs are unchanged');
{
  check('holding unit 000 loses every list',
    m.getUnnecessaryDistributionLists({ unit: '000', scope: 'UNIT', charter: 'PCR-CA-000', type: 'SENIOR' }, 'ca000'),
    [
      'ca000.allhands@example.org',
      'ca000.cadets@example.org',
      'ca000.seniors@example.org',
      'ca000.parents@example.org'
    ]);

  check('group echelon loses every list',
    m.getUnnecessaryDistributionLists({ unit: '010', scope: 'GROUP', charter: 'PCR-CA-010', type: 'GROUP' }, 'ca010'),
    [
      'ca010.allhands@example.org',
      'ca010.cadets@example.org',
      'ca010.seniors@example.org',
      'ca010.parents@example.org'
    ]);
}

done();
