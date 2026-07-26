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
 * A member belongs only to the group for their HIGHEST completed level — the levels
 * are rungs, not badges that accumulate — and professionalLevelInProgress covers the
 * people partway up one.
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
  '100001': senior('100001', 'alex.l2@example.org'),      // highest completed: Level 2
  '100002': senior('100002', 'blair.part1@example.org'),  // Level 2 Part 1 only
  '100003': senior('100003', 'casey.l3@example.org'),     // highest completed: Level 3
  '100004': senior('100004', 'devon.pending@example.org'),// Level 2 credits, not approved
  '100005': senior('100005', 'erin.l1@example.org'),      // highest completed: Level 1
  '100006': senior('100006', 'frankie.l5@example.org')    // all the way up: Level 5
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
  '"8","4","100005","8","01/05/2021","01/05/2021","{CreditEarned:false}"',
  // Frankie climbed the whole ladder: 1, both parts of 2, then 3, 4, 5.
  '"9","4","100006","8","02/01/2019","02/01/2019","{CreditEarned:false}"',
  '"10","7","100006","8","03/01/2019","03/01/2019","{CreditEarned:false}"',
  '"11","8","100006","8","04/01/2019","04/01/2019","{CreditEarned:false}"',
  '"12","3","100006","8","05/01/2020","05/01/2020","{CreditEarned:false}"',
  '"13","2","100006","8","06/01/2021","06/01/2021","{CreditEarned:false}"',
  '"14","1","100006","8","07/01/2022","07/01/2022","{CreditEarned:false}"'
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
  }, ['getGroupMembers', 'resolveProfessionalLevelSpec_', 'getApprovedPathCreditsByCapid_',
    'getProfessionalLevelLadder_', 'summarizeMemberLevelStanding_']);
  return { m, calls };
}

function membersOf(generated, groupId) {
  return Object.keys(generated[groupId] || {}).sort();
}

// ---------------------------------------------------------------------------
section('1. A member sits on their HIGHEST rung and no other');
{
  const { m } = load();

  const l2 = m.getGroupMembers('all-level-ii', 'professionalLevel', 'Level 2', members, squadrons);
  check('the Level 2 group holds only those whose highest is Level 2',
    membersOf(l2, 'ca.all-level-ii'), ['alex.l2@example.org']);
  check('a Level 5 holder is NOT in the Level 2 group',
    membersOf(l2, 'ca.all-level-ii').includes('frankie.l5@example.org'), false);

  const l5 = m.getGroupMembers('all-level-v', 'professionalLevel', 'Level 5', members, squadrons);
  check('they are in the Level 5 group instead',
    membersOf(l5, 'ca.all-level-v'), ['frankie.l5@example.org']);

  const l1 = m.getGroupMembers('all-level-i', 'professionalLevel', 'Level 1', members, squadrons);
  check('Level 1 holds only the member who stopped there',
    membersOf(l1, 'ca.all-level-i'), ['erin.l1@example.org']);

  check('group-echelon rollup still built',
    membersOf(l2, 'ca010.all-level-ii'), ['alex.l2@example.org']);
}

// ---------------------------------------------------------------------------
section('2. Level 2 counts as completed only with BOTH parts');
{
  const { m } = load();
  const l2 = m.getGroupMembers('all-level-ii', 'professionalLevel', 'Level 2', members, squadrons);
  check('Part 1 alone is not the level',
    membersOf(l2, 'ca.all-level-ii').includes('blair.part1@example.org'), false);

  const l1 = m.getGroupMembers('all-level-i', 'professionalLevel', 'Level 1', members, squadrons);
  check('and Part 1 alone does not promote them past Level 1',
    membersOf(l1, 'ca.all-level-i').includes('blair.part1@example.org'), false);
}

// ---------------------------------------------------------------------------
section('3. professionalLevelInProgress: Part 1 done, Part 2 not');
{
  const { m } = load();
  const partial = m.getGroupMembers(
    'all-level-ii-part-1-only', 'professionalLevelInProgress', 'Level 2', members, squadrons);

  check('the Part 1 holder is the whole list',
    membersOf(partial, 'ca.all-level-ii-part-1-only'), ['blair.part1@example.org']);
  check('someone who finished both parts is not in progress',
    membersOf(partial, 'ca.all-level-ii-part-1-only').includes('alex.l2@example.org'), false);
  check('nor is someone who never started',
    membersOf(partial, 'ca.all-level-ii-part-1-only').includes('erin.l1@example.org'), false);
  check('unapproved credit is not progress',
    membersOf(partial, 'ca.all-level-ii-part-1-only').includes('devon.pending@example.org'), false);
}

// ---------------------------------------------------------------------------
section('4. A single-path level has no in-progress state, and says so');
{
  const { m, calls } = load();
  const generated = m.getGroupMembers(
    'all-level-iii-partial', 'professionalLevelInProgress', 'Level 3', members, squadrons);

  check('no group', Object.keys(generated), []);
  check('warned that the level has no parts',
    calls.warn.some(w => /no parts, so it has no in-progress state/i.test(w.msg)), true);
}

// ---------------------------------------------------------------------------
section('5. Roman numerals, digits and PathIDs all reach the same rung');
{
  check('"Level II"', membersOf(
    load().m.getGroupMembers('all-level-ii', 'professionalLevel', 'Level II', members, squadrons),
    'ca.all-level-ii'), ['alex.l2@example.org']);

  check('"Level 3"', membersOf(
    load().m.getGroupMembers('all-level-iii', 'professionalLevel', 'Level 3', members, squadrons),
    'ca.all-level-iii'), ['casey.l3@example.org']);

  check('PathID 3 is the Level 3 rung', membersOf(
    load().m.getGroupMembers('all-level-iii', 'professionalLevel', '3', members, squadrons),
    'ca.all-level-iii'), ['casey.l3@example.org']);

  check('a PathID of one PART still means the whole level', membersOf(
    load().m.getGroupMembers('all-level-ii', 'professionalLevel', '7', members, squadrons),
    'ca.all-level-ii'), ['alex.l2@example.org']);
}

// ---------------------------------------------------------------------------
section('6. Approval status is honored, and read from PL_Lookup');
{
  const { m } = load();
  const credits = m.getApprovedPathCreditsByCapid_();
  check('only approved rows indexed', Object.keys(credits).sort(),
    ['100001', '100002', '100003', '100005', '100006']);
  check('both Level 2 parts recorded for the holder',
    Object.keys(credits['100001']).sort(), ['4', '7', '8']);
}

// ---------------------------------------------------------------------------
section('7. Several values are an OR across rungs');
{
  const { m } = load();
  const generated = m.getGroupMembers(
    'all-level-ii-or-iii', 'professionalLevel', 'Level 2,Level 3', members, squadrons);
  check('either rung qualifies', membersOf(generated, 'ca.all-level-ii-or-iii'),
    ['alex.l2@example.org', 'casey.l3@example.org']);
  check('the Level 5 holder is still excluded',
    membersOf(generated, 'ca.all-level-ii-or-iii').includes('frankie.l5@example.org'), false);
}

// ---------------------------------------------------------------------------
section('8. The ladder, and one member standing on it');
{
  const { m } = load();
  const ladder = m.getProfessionalLevelLadder_();
  check('Level 2 is two paths', ladder['2'], ['7', '8']);
  check('Level 3 is one', ladder['3'], ['3']);
  check('non-level paths are not rungs', ladder['12'], undefined);

  const climber = m.summarizeMemberLevelStanding_({ '4': true, '7': true, '8': true, '3': true }, ladder);
  check('highest wins', climber.highest, '3');
  check('completed lists every rung passed', climber.completed.sort(), ['1', '2', '3']);

  const midway = m.summarizeMemberLevelStanding_({ '4': true, '7': true }, ladder);
  check('partial rung reported', midway.partial, ['2']);
  check('highest stays at the last finished rung', midway.highest, '1');
}

// ---------------------------------------------------------------------------
section('9. Value resolution: rungs vs plain paths');
{
  const { m } = load();
  check('a level becomes a rung',
    m.resolveProfessionalLevelSpec_(['Level 2'], 'x').levels, ['2']);
  check('a non-level path stays a path',
    m.resolveProfessionalLevelSpec_(['Squadron Commander Training'], 'x').paths[0].pathIds, ['12']);
  check('and contributes no rung',
    m.resolveProfessionalLevelSpec_(['Squadron Commander Training'], 'x').levels, []);
}

// ---------------------------------------------------------------------------
section('10. Nobody at a rung: no group, and the reason says which case');
{
  const { m, calls } = load();
  const generated = m.getGroupMembers('all-level-iv', 'professionalLevel', 'Level 4', members, squadrons);
  check('no group', Object.keys(generated), []);
  check('reported', calls.warn.some(w => /matched no members/i.test(w.msg)), true);
  check('names the everyone-moved-past case',
    calls.warn.some(w => /gone past it/i.test(JSON.stringify(w))), true);
}

// ---------------------------------------------------------------------------
section('11. A missing PL table is reported, not guessed at');
{
  const { m, calls } = load({ 'PL_Paths.txt': PL_PATHS, 'PL_Lookup.txt': PL_LOOKUP });
  const generated = m.getGroupMembers('all-level-ii', 'professionalLevel', 'Level 2', members, squadrons);

  check('no group created', Object.keys(generated), []);
  check('names the missing file',
    calls.warn.some(w => /PL_MemberPathCredit/.test(JSON.stringify(w))), true);
}

// ---------------------------------------------------------------------------
section('12. An unresolvable value is named in the log');
{
  const { m, calls } = load();
  m.getGroupMembers('all-level-vi', 'professionalLevel', 'Level 6', members, squadrons);
  check('warned', calls.warn.some(w => /could not be resolved to a PL path/i.test(w.msg)), true);
}

done();
