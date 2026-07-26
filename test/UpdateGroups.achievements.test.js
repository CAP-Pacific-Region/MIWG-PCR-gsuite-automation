/**
 * UpdateGroups.gs — achievement-driven DLs (the Education & Training level lists).
 *
 * MbrAchievements.txt identifies an achievement only by numeric AchvID; the name
 * lives in Achievements.txt. A Groups sheet Values column reading "Level II" —
 * the natural thing to write — therefore matched no row, and the run created a
 * real, permanently empty Google Group without logging an error. Names now
 * resolve, and a row that still matches nobody is left unmanaged and reported
 * rather than created empty.
 *
 * All CAPIDs and addresses here are synthetic.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeDrive, makeChecker, Utilities } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'UpdateGroups.gs');
const { section, check, done } = makeChecker();

const CONFIG = {
  WING: 'CA',
  EMAIL_DOMAIN: '@example.org',
  CAPWATCH_DATA_FOLDER_ID: 'folder-id'
};

const squadrons = {
  '188': { orgid: '188', name: 'CALIFORNIA WING HQ', charter: 'PCR-CA-001', unit: '001', nextLevel: '', scope: 'WING', wing: 'CA' },
  '900': { orgid: '900', name: 'GROUP 10', charter: 'PCR-CA-010', unit: '010', nextLevel: '188', scope: 'GROUP', wing: 'CA' },
  '2500': { orgid: '2500', name: 'EXAMPLE SR SQDN 80', charter: 'PCR-CA-080', unit: '080', nextLevel: '900', scope: 'UNIT', wing: 'CA' }
};

const members = {
  '100001': { capsn: '100001', orgid: '2500', group: '900', email: 'alex.two@example.org', type: 'SENIOR', dutyPositions: [], dutyPositionIds: [] },
  '100002': { capsn: '100002', orgid: '2500', group: '900', email: 'blair.training@example.org', type: 'SENIOR', dutyPositions: [], dutyPositionIds: [] },
  '100003': { capsn: '100003', orgid: '2500', group: '900', email: 'casey.pending@example.org', type: 'SENIOR', dutyPositions: [], dutyPositionIds: [] },
  '100004': { capsn: '100004', orgid: '2500', group: '900', email: 'devon.one@example.org', type: 'SENIOR', dutyPositions: [], dutyPositionIds: [] }
};

// Real header spelling; the name column has had more than one over the years,
// which is why the index looks for several.
const ACHIEVEMENTS_TXT = [
  'AchvID,MemberType,Achv,LastUpdate',
  '96,SENIOR,Level I,2026-01-01',
  '97,SENIOR,Level II,2026-01-01',
  '53,CADET,Curry Achievement,2026-01-01'
].join('\n');

// parseFile() strips the header row, so these are data rows only:
// [CAPID, AchvID, Status, ...]
const MBR_ACHIEVEMENTS = [
  ['100001', '97', 'ACTIVE'],
  ['100002', '97', 'TRAINING'],
  ['100003', '97', 'PENDING'],
  ['100004', '96', 'ACTIVE']
];

function load() {
  const { logger, calls } = makeLogger();
  const m = loadModule(MODULE, {
    Logger: logger,
    CONFIG: CONFIG,
    Utilities: Utilities,
    DriveApp: makeDrive({ 'Achievements.txt': ACHIEVEMENTS_TXT }),
    parseFile: name => (name === 'MbrAchievements' ? MBR_ACHIEVEMENTS : [])
  }, ['getGroupMembers', 'normalizeAchievementLabel_', 'resolveAchievementValuesToIds_']);
  return { m, calls };
}

function membersOf(generated, groupId) {
  return Object.keys(generated[groupId] || {}).sort();
}

// ---------------------------------------------------------------------------
section('1. A Values column written as a name resolves to the AchvID');
{
  const { m } = load();
  const generated = m.getGroupMembers('all-level-ii', 'achievements', 'Level II', members, squadrons);

  check('ACTIVE and TRAINING holders are in',
    membersOf(generated, 'ca.all-level-ii'),
    ['alex.two@example.org', 'blair.training@example.org']);
  check('a PENDING row is not a completion',
    membersOf(generated, 'ca.all-level-ii').includes('casey.pending@example.org'), false);
  check('a different level is a different group',
    membersOf(generated, 'ca.all-level-ii').includes('devon.one@example.org'), false);
  check('group-echelon child DL still built',
    membersOf(generated, 'ca010.all-level-ii'),
    ['alex.two@example.org', 'blair.training@example.org']);
}

// ---------------------------------------------------------------------------
section('2. Numeric AchvIDs keep working, and spelling variants reach the same place');
{
  const byId = load().m.getGroupMembers('all-level-ii', 'achievements', '97', members, squadrons);
  check('numeric', membersOf(byId, 'ca.all-level-ii'),
    ['alex.two@example.org', 'blair.training@example.org']);

  const byArabic = load().m.getGroupMembers('all-level-ii', 'achievements', 'Level 2', members, squadrons);
  check('Roman numeral folded to a digit', membersOf(byArabic, 'ca.all-level-ii'),
    ['alex.two@example.org', 'blair.training@example.org']);

  const byCase = load().m.getGroupMembers('all-level-ii', 'achievements', '  level ii  ', members, squadrons);
  check('case and padding', membersOf(byCase, 'ca.all-level-ii'),
    ['alex.two@example.org', 'blair.training@example.org']);
}

// ---------------------------------------------------------------------------
section('3. Several values in one row');
{
  const { m } = load();
  const generated = m.getGroupMembers('all-level-i-and-ii', 'achievements', 'Level I,Level II', members, squadrons);
  check('union of both achievements',
    membersOf(generated, 'ca.all-level-i-and-ii'),
    ['alex.two@example.org', 'blair.training@example.org', 'devon.one@example.org']);
}

// ---------------------------------------------------------------------------
section('4. A value that matches nothing is reported, and creates nothing');
{
  const { m, calls } = load();
  const generated = m.getGroupMembers('all-level-vii', 'achievements', 'Level VII', members, squadrons);

  check('no group generated', Object.keys(generated), []);
  check('unresolved value warned about',
    calls.warn.some(w => /could not be resolved/i.test(w.msg)), true);
  check('empty group warned about',
    calls.warn.some(w => /matched no members/i.test(w.msg)), true);
}

// ---------------------------------------------------------------------------
section('5. An achievement nobody holds is left unmanaged rather than emptied');
{
  const { m, calls } = load();
  const generated = m.getGroupMembers('all-curry', 'achievements', 'Curry Achievement', members, squadrons);

  check('resolves to an AchvID', Object.keys(m.resolveAchievementValuesToIds_(['Curry Achievement'], 'x')), ['53']);
  check('but produces no group', Object.keys(generated), []);
  check('reported as matching no members',
    calls.warn.some(w => /matched no members/i.test(w.msg)), true);
}

// ---------------------------------------------------------------------------
section('6. Label normalization');
{
  const { m } = load();
  check('punctuation and case', m.normalizeAchievementLabel_('Level-II'), 'level 2');
  check('collapses whitespace', m.normalizeAchievementLabel_('  Level   II '), 'level 2');
  check('leaves ordinary names alone', m.normalizeAchievementLabel_('Curry Achievement'), 'curry achievement');
  check('null-safe', m.normalizeAchievementLabel_(null), '');
}

done();
