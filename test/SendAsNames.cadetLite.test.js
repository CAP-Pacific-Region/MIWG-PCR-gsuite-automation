/**
 * UpdateMembers.gs — which accounts the Send-As / displayName sync will name.
 *
 * WHAT WENT WRONG
 *
 * processSendAsNamesBatchLocked() built its roster with a bare getMembers().
 * That signature is:
 *
 *     getMembers(types = ACTIVE, includeDutyPositions = true, includeCadetLite = false)
 *
 * so cadet-lite members were filtered out. The loop, however, walks accounts
 * WORKSPACE already holds — and some cadet-lite members hold one. Each of those
 * matched no CAPWATCH record, was logged as "No CAPWATCH record for user", and
 * kept whatever name it was created with. Rank changes never landed, which is
 * most of what this job is for.
 *
 * Worse, that warning then covered two unrelated populations: members the roster
 * had filtered out, and accounts with no member behind them at all. The second
 * is the one worth chasing, and it was buried in the first.
 *
 * THE LESSON THESE ASSERTIONS ENCODE
 * The cadet-lite rule decides WHO GETS AN ACCOUNT. It has nothing to say about
 * what an account that already exists is called. A rule applied outside the
 * decision it was written for reads as a data problem — a missing record —
 * rather than as the policy it actually is.
 *
 * All CAPIDs and names here are synthetic.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'UpdateMembers.gs');
const { section, check, done } = makeChecker();

const EXCLUDED = ['CADET', 'C/Amn', 'C/A1C', 'C/SrA'];

/** Loads the module against a CONFIG, the way each tenant supplies its own. */
function load(config) {
  return loadModule(MODULE, {
    CONFIG: config,
    Logger: { info: () => {}, warn: () => {}, error: () => {} },
    Utilities: { sleep: () => {} },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) }
  }, ['isCadetLiteGrade_', 'shouldProcessMember']);
}

const cadetsTenant = load({
  CADET_LITE: true,
  CADET_LITE_EXCLUDED_GRADES: EXCLUDED,
  MEMBER_TYPES: { ACTIVE: ['CADET'] },
  EXCLUDED_ORG_IDS: []
});

// ---------------------------------------------------------------------------
section('1. Who the cadet-lite rule covers');
{
  const isLite = cadetsTenant.isCadetLiteGrade_;

  EXCLUDED.forEach(rank => {
    check(`${rank} is cadet-lite`, isLite(rank), true);
  });

  check('C/SSgt is not — this is the threshold', isLite('C/SSgt'), false);
  check('C/Maj is not', isLite('C/Maj'), false);
  check('a senior grade is not', isLite('Capt'), false);
}

// ---------------------------------------------------------------------------
section('2. It reads the same list the roster filter reads');
{
  // One definition, not two. If CONFIG.CADET_LITE_EXCLUDED_GRADES changes, both
  // the roster filter and the reporting helper have to move together — a second
  // hardcoded copy is exactly how "who is cadet-lite" drifts apart.
  const isLite = cadetsTenant.isCadetLiteGrade_;
  const shouldProcess = cadetsTenant.shouldProcessMember;

  // memberRow: [0]=capid ... [14]=rank ... [21]=type ... [24]=status
  const row = (rank) => {
    const r = new Array(25).fill('');
    r[0] = '100001';
    r[14] = rank;
    r[21] = 'CADET';
    r[24] = 'ACTIVE';
    return r;
  };

  EXCLUDED.concat(['C/SSgt', 'C/Maj']).forEach(rank => {
    const filteredOut = !shouldProcess(row(rank), ['CADET'], false);
    check(`${rank}: helper and roster filter agree`, isLite(rank), filteredOut);
  });
}

// ---------------------------------------------------------------------------
section('3. Asking for them overrides the filter');
{
  // The call the Send-As sync now makes. Without this, an account belonging to
  // one of these members has no record to be named from.
  const shouldProcess = cadetsTenant.shouldProcessMember;
  const r = new Array(25).fill('');
  r[0] = '100002';
  r[14] = 'C/Amn';
  r[21] = 'CADET';
  r[24] = 'ACTIVE';

  check('excluded by default', shouldProcess(r, ['CADET'], false), false);
  check('included when asked for', shouldProcess(r, ['CADET'], true), true);
}

// ---------------------------------------------------------------------------
section('4. A tenant without the rule has no cadet-lite at all');
{
  // The seniors and region tenants never withhold accounts this way, so nothing
  // should be reported as cadet-lite there however the grades are spelled.
  const seniors = load({
    CADET_LITE: false,
    CADET_LITE_EXCLUDED_GRADES: EXCLUDED,
    MEMBER_TYPES: { ACTIVE: ['SENIOR'] },
    EXCLUDED_ORG_IDS: []
  });

  check('the rule is off, so no grade is cadet-lite',
    EXCLUDED.map(seniors.isCadetLiteGrade_), [false, false, false, false]);
}

// ---------------------------------------------------------------------------
section('5. Reading a grade defensively');
{
  const isLite = cadetsTenant.isCadetLiteGrade_;

  check('padding does not defeat the match', isLite('  C/Amn  '), true);
  check('a blank grade is not cadet-lite', isLite(''), false);
  check('null is not', isLite(null), false);
  check('undefined is not', isLite(undefined), false);

  const noList = load({
    CADET_LITE: true,
    MEMBER_TYPES: { ACTIVE: ['CADET'] },
    EXCLUDED_ORG_IDS: []
  });
  check('a tenant with the rule on but no list reports nobody',
    noList.isCadetLiteGrade_('C/Amn'), false);
}

done();
