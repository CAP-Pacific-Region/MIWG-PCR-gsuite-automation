/**
 * UpdateGroups.gs — professional development level DLs, read from the PL_* tables.
 *
 * Levels I–V are not in MbrAchievements. The post-2018 PD program has its own
 * CAPWATCH subsystem: PL_Paths names each path, PL_MemberPathCredit records who
 * holds credit for it, PL_Lookup says which StatusID counts as approved.
 * Achievements.txt still lists the RETIRED "Level II".."Level V" (AchvIDs 131-134),
 * which resolve cleanly and then match nobody — the trap this attribute exists to
 * get out of.
 *
 * Fixtures use the real column names and the real CAPWATCH PathIDs. CAPIDs and
 * addresses are synthetic.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeDrive, makeChecker, Utilities } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'UpdateGroups.gs');
const { section, check, done } = makeChecker();

const CONFIG = { WING: 'CA', EMAIL_DOMAIN: '@example.org', CAPWATCH_DATA_FOLDER_ID: 'folder-id' };

const squadrons = {
  '188': { orgid: '188', name: 'CALIFORNIA WING HQ', charter: 'PCR-CA-001', unit: '001', scope: 'WING', wing: 'CA' },
  '900': { orgid: '900', name: 'GROUP 10', charter: 'PCR-CA-010', unit: '010', nextLevel: '188', scope: 'GROUP', wing: 'CA' },
  '2500': { orgid: '2500', name: 'EXAMPLE SR SQDN 80', charter: 'PCR-CA-080', unit: '080', nextLevel: '900', scope: 'UNIT', wing: 'CA' }
};

function senior(capsn, email) {
  return { capsn, orgid: '2500', group: '900', email, type: 'SENIOR', dutyPositions: [], dutyPositionIds: [] };
}

const members = {
  '100001': senior('100001', 'alex.l2@example.org'),      // both Level 2 parts
  '100002': senior('100002', 'blair.part1@example.org'),  // Part 1 only
  '100003': senior('100003', 'casey.l3@example.org'),     // Level 3
  '100004': senior('100004', 'devon.pending@example.org'),// Level 2, not approved
  '100005': senior('100005', 'erin.l1@example.org')       // Level 1 only
};

// Real names and PathIDs from PL_Paths.txt.
const PL_PATHS = [
  'PathID,PathName',
  '"4","Level 1"',
  '"7","Level 2 Part 1"',
  '"8","Level 2 Part 2"',
  '"3","Level 3"',
  '"2","Level 4"',
  '"1","Level 5"',
  '"12","Squadron Commander Training"'
].join('\n');

const PL_LOOKUP = [
  'LookupID,LookupType,LookupValue',
  '"8","ApprovalStatus","APPROVED"',
  '"26","ApprovalStatus","PENDING"',
  '"27","ApprovalStatus","DISAPPROVED"',
  '"9","MemberLevel","LV1"'
].join('\n');

const PL_MEMBER_PATH_CREDIT = [
  'MemberPathCreditID,PathID,CAPID,StatusID,Completed,Expiration,ExtraCreditEarned',
  '"1","4","100001","8","08/09/2020","08/09/2020","{CreditEarned:false}"',
  '"2","7","100001","8","09/13/2020","09/13/2020","{CreditEarned:false}"',
  '"3","8","100001","8","10/08/2020","10/08/2020","{CreditEarned:false}"',
  '"4","7","100002","8","09/15/2020","09/15/2020","{CreditEarned:false}"',
  '"5","3","100003","8","11/23/2020","11/23/2020","{CreditEarned:false}"',
  '"6","7","100004","26","12/01/2020","12/01/2020","{CreditEarned:false}"',
  '"7","8","100004","26","12/02/2020","12/02/2020","{CreditEarned:false}"',
  '"8","4","100005","8","01/05/2021","01/05/2021","{CreditEarned:false}"'
].join('\n');

function load(files) {
  const { logger, calls } = makeLogger();
  const m = loadModule(MODULE, {
    Logger: logger,
    CONFIG: CONFIG,
    Utilities: Utilities,
    DriveApp: makeDrive(files || {
      'PL_Paths.txt': PL_PATHS,
      'PL_Lookup.txt': PL_LOOKUP,
      'PL_MemberPathCredit.txt': PL_MEMBER_PATH_CREDIT
    }),
    parseFile: () => []
  }, ['getGroupMembers', 'resolveProfessionalLevelRequirements_', 'getApprovedPathCreditsByCapid_']);
  return { m, calls };
}

function membersOf(generated, groupId) {
  return Object.keys(generated[groupId] || {}).sort();
}

// ---------------------------------------------------------------------------
section('1. Level 2 needs BOTH parts');
{
  const { m } = load();
  const generated = m.getGroupMembers('all-level-ii', 'professionalLevel', 'Level 2', members, squadrons);

  check('holder of Part 1 + Part 2 is in',
    membersOf(generated, 'ca.all-level-ii'), ['alex.l2@example.org']);
  check('Part 1 alone is progress, not the level',
    membersOf(generated, 'ca.all-level-ii').includes('blair.part1@example.org'), false);
  check('group-echelon rollup built',
    membersOf(generated, 'ca010.all-level-ii'), ['alex.l2@example.org']);
}

// ---------------------------------------------------------------------------
section('2. Roman numerals, digits and PathIDs all reach the same place');
{
  check('"Level II"', membersOf(
    load().m.getGroupMembers('all-level-ii', 'professionalLevel', 'Level II', members, squadrons),
    'ca.all-level-ii'), ['alex.l2@example.org']);

  check('"Level 3"', membersOf(
    load().m.getGroupMembers('all-level-iii', 'professionalLevel', 'Level 3', members, squadrons),
    'ca.all-level-iii'), ['casey.l3@example.org']);

  check('PathID 3', membersOf(
    load().m.getGroupMembers('all-level-iii', 'professionalLevel', '3', members, squadrons),
    'ca.all-level-iii'), ['casey.l3@example.org']);
}

// ---------------------------------------------------------------------------
section('3. Approval status is honored, and read from PL_Lookup');
{
  const { m } = load();
  const generated = m.getGroupMembers('all-level-ii', 'professionalLevel', 'Level 2', members, squadrons);
  check('PENDING credit does not count',
    membersOf(generated, 'ca.all-level-ii').includes('devon.pending@example.org'), false);

  const credits = m.getApprovedPathCreditsByCapid_();
  check('only approved rows indexed', Object.keys(credits).sort(),
    ['100001', '100002', '100003', '100005']);
  check('both Level 2 parts recorded for the holder',
    Object.keys(credits['100001']).sort(), ['4', '7', '8']);
}

// ---------------------------------------------------------------------------
section('4. Several values are an OR');
{
  const { m } = load();
  const generated = m.getGroupMembers('all-level-ii-plus', 'professionalLevel', 'Level 2,Level 3', members, squadrons);
  check('either level qualifies', membersOf(generated, 'ca.all-level-ii-plus'),
    ['alex.l2@example.org', 'casey.l3@example.org']);
}

// ---------------------------------------------------------------------------
section('5. Requirement resolution');
{
  const { m } = load();
  const parts = m.resolveProfessionalLevelRequirements_(['Level 2'], 'x');
  check('Level 2 resolves to both parts', parts[0].pathIds, ['7', '8']);

  const single = m.resolveProfessionalLevelRequirements_(['Level 5'], 'x');
  check('Level 5 is one path', single[0].pathIds, ['1']);

  const named = m.resolveProfessionalLevelRequirements_(['Squadron Commander Training'], 'x');
  check('any path name works, not just levels', named[0].pathIds, ['12']);
}

// ---------------------------------------------------------------------------
section('6. A level nobody holds creates nothing, and says why');
{
  const { m, calls } = load();
  const generated = m.getGroupMembers('all-level-v', 'professionalLevel', 'Level 5', members, squadrons);
  check('no group', Object.keys(generated), []);
  check('reported', calls.warn.some(w => /matched no members/i.test(w.msg)), true);
}

// ---------------------------------------------------------------------------
section('7. A missing PL table is reported, not guessed at');
{
  const { m, calls } = load({ 'PL_Paths.txt': PL_PATHS, 'PL_Lookup.txt': PL_LOOKUP });
  const generated = m.getGroupMembers('all-level-ii', 'professionalLevel', 'Level 2', members, squadrons);

  check('no group created', Object.keys(generated), []);
  check('names the missing file',
    calls.warn.some(w => /PL_MemberPathCredit/.test(JSON.stringify(w))), true);
}

// ---------------------------------------------------------------------------
section('8. An unresolvable value is named in the log');
{
  const { m, calls } = load();
  m.getGroupMembers('all-level-vi', 'professionalLevel', 'Level 6', members, squadrons);
  check('warned', calls.warn.some(w => /could not be resolved to a PL path/i.test(w.msg)), true);
}

done();
