/**
 * SquadronGroups.gs — where a resumed run picks up.
 *
 * The bug this exists to prevent already happened once, in the other direction:
 * updateAllSquadronGroups() stopped on time without recording a position, so every
 * run restarted at the top and died at the same place, and the last 9 of 68
 * squadrons on the CAWG cadet tenant were never reached on any run for weeks.
 * Nothing reported an error — a run that stops early still succeeds at what it did.
 *
 * Resuming introduces the opposite hazard. The parked position is an INDEX into a
 * list rebuilt from CAPWATCH each execution, so a unit chartering or folding shifts
 * every index after it. Resuming on a stale number would skip squadrons silently —
 * the exact failure mode being fixed, reintroduced by the fix. Hence the charter
 * check, and hence these assertions: the resolver must prefer redundant work over
 * any chance of a skip.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'squadron-groups', 'SquadronGroups.gs');
const { section, check, done } = makeChecker();

const { logger, calls: logs } = makeLogger();
const m = loadModule(MODULE, {
  Logger: logger,
  CONFIG: { WING: 'CA', EMAIL_DOMAIN: '@example.org' },
  parseFile: () => []
}, ['resolveSquadronResumePosition_']);

/** A UNIT-scope squadron list in iteration order, like the real loop walks. */
function roster(charters) {
  return charters.map(c => ({ charter: c, scope: 'UNIT' }));
}

const LIST = roster(['PCR-CA-006', 'PCR-CA-070', 'PCR-CA-146', 'PCR-CA-403', 'PCR-CA-478']);

function warnedAbout(pattern) {
  return logs.warn.some(l => pattern.test(l.msg));
}

// ---------------------------------------------------------------------------
section('1. A fresh run starts at the beginning');
{
  check('no parked state', m.resolveSquadronResumePosition_(LIST, null), 0);
  check('undefined is not a position', m.resolveSquadronResumePosition_(LIST, undefined), 0);
  check('index 0 is the beginning anyway',
    m.resolveSquadronResumePosition_(LIST, { squadronIndex: 0, charterAtIndex: 'PCR-CA-006' }), 0);
}

// ---------------------------------------------------------------------------
section('2. An intact list resumes exactly where it stopped');
{
  check('mid-list', m.resolveSquadronResumePosition_(LIST,
    { squadronIndex: 3, charterAtIndex: 'PCR-CA-403' }), 3);

  check('the last squadron — the one that was starved',
    m.resolveSquadronResumePosition_(LIST, { squadronIndex: 4, charterAtIndex: 'PCR-CA-478' }), 4);

  check('a position with no parked charter is trusted (older state file)',
    m.resolveSquadronResumePosition_(LIST, { squadronIndex: 2 }), 2);
}

// ---------------------------------------------------------------------------
section('3. A list that moved starts over rather than skipping');
{
  // A unit chartered ahead of the parked position: index 3 is now 146, not 403.
  const grown = roster(['PCR-CA-006', 'PCR-CA-008', 'PCR-CA-070', 'PCR-CA-146', 'PCR-CA-403', 'PCR-CA-478']);
  check('a new unit shifted everything after it',
    m.resolveSquadronResumePosition_(grown, { squadronIndex: 3, charterAtIndex: 'PCR-CA-403' }), 0);
  check('and it says why', warnedAbout(/list changed/i), true);

  // A unit folded: the list is shorter than the parked index.
  const shrunk = roster(['PCR-CA-006', 'PCR-CA-070']);
  check('parked past the end',
    m.resolveSquadronResumePosition_(shrunk, { squadronIndex: 4, charterAtIndex: 'PCR-CA-478' }), 0);
  check('and it says why', warnedAbout(/past the end/i), true);

  check('exactly at the end is also past it (nothing left to do there)',
    m.resolveSquadronResumePosition_(LIST, { squadronIndex: 5, charterAtIndex: '' }), 0);
}

// ---------------------------------------------------------------------------
section('4. Malformed state never becomes a wrong position');
{
  check('a missing index', m.resolveSquadronResumePosition_(LIST, {}), 0);
  check('a non-numeric index',
    m.resolveSquadronResumePosition_(LIST, { squadronIndex: 'three' }), 0);
  check('a negative index',
    m.resolveSquadronResumePosition_(LIST, { squadronIndex: -2 }), 0);
  check('NaN', m.resolveSquadronResumePosition_(LIST, { squadronIndex: NaN }), 0);
  check('Infinity', m.resolveSquadronResumePosition_(LIST, { squadronIndex: Infinity }), 0);
  check('a numeric string is still a usable position',
    m.resolveSquadronResumePosition_(LIST, { squadronIndex: '3', charterAtIndex: 'PCR-CA-403' }), 3);
}

// ---------------------------------------------------------------------------
section('5. Every rejection lands on 0, never on a guess');
{
  const hostile = [
    null, undefined, {}, { squadronIndex: null }, { squadronIndex: {} },
    { squadronIndex: 99, charterAtIndex: 'PCR-CA-478' },
    { squadronIndex: 2, charterAtIndex: 'PCR-CA-999' }
  ];
  const results = hostile.map(r => m.resolveSquadronResumePosition_(LIST, r));
  check('all resolve to a valid index', results.every(i => i >= 0 && i < LIST.length), true);
  check('and none of them invents a mid-list position',
    results.filter(i => i !== 0), []);
}

done();
