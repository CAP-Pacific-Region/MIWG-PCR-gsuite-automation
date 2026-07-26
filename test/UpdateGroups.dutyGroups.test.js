/**
 * UpdateGroups.gs — which member lands in a duty-position DL.
 *
 * The bug this pins: a wing office DL built with `dutyPositionIdsWingHQ` used to
 * test the MEMBER'S HOME UNIT rather than the org the duty is held at. Nearly all
 * wing staff are members of a squadron and hold their wing duty on top of that, so
 * the DL came back holding whichever holder happened to also be a Wing HQ member —
 * in the live case the assistant, not the primary.
 *
 * Also pinned: duty titles are compared after the CAPR 30-1 renames, so a row
 * keyed on the current title still matches eServices rows carrying the retired one.
 *
 * All names, CAPIDs and addresses here are synthetic.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'UpdateGroups.gs');
const { section, check, done } = makeChecker();

const CONFIG = {
  WING: 'CA',
  EMAIL_DOMAIN: '@example.org',
  CAPWATCH_DATA_FOLDER_ID: 'folder-id'
};

/** Wing HQ, one group echelon org, one squadron under it. */
const squadrons = {
  '188': { orgid: '188', name: 'CALIFORNIA WING HQ', charter: 'PCR-CA-001', unit: '001', nextLevel: '', scope: 'WING', wing: 'CA' },
  '900': { orgid: '900', name: 'GROUP 10', charter: 'PCR-CA-010', unit: '010', nextLevel: '188', scope: 'GROUP', wing: 'CA' },
  '2500': { orgid: '2500', name: 'EXAMPLE SR SQDN 80', charter: 'PCR-CA-080', unit: '080', nextLevel: '900', scope: 'UNIT', wing: 'CA' }
};

/**
 * @param {Object} over - capsn/orgid/group/email overrides
 * @param {Array} duties - [{ id, charter, assistant }]
 */
function member(over, duties) {
  return Object.assign({
    capsn: over.capsn,
    orgid: over.orgid,
    group: over.group,
    email: over.email,
    type: 'SENIOR',
    dutyPositions: duties.map(d => ({
      value: `${d.id} (${d.assistant ? 'A' : 'P'}) (${d.charter})`,
      id: d.id,
      level: d.level || 'WING',
      assistant: !!d.assistant
    })),
    dutyPositionIds: duties.map(d => d.id)
  }, over);
}

const members = {
  // Primary wing Recruiting Officer: a member of a squadron, duty held at Wing HQ.
  // This is the person the old home-unit test dropped.
  '100001': member(
    { capsn: '100001', orgid: '2500', group: '900', email: 'alex.primary@example.org' },
    [{ id: 'Recruiting Officer', charter: 'PCR-CA-001' }]
  ),
  // Assistant: happens to also be a Wing HQ member, so the old code found them.
  '100002': member(
    { capsn: '100002', orgid: '188', group: '', email: 'blair.assistant@example.org' },
    [{ id: 'Recruiting Officer', charter: 'PCR-CA-001', assistant: true }]
  ),
  // Wing HQ member holding the SAME title at a squadron. The old code included
  // them on the strength of their home unit alone.
  '100003': member(
    { capsn: '100003', orgid: '188', group: '', email: 'casey.squadron@example.org' },
    [{ id: 'Recruiting Officer', charter: 'PCR-CA-080', level: 'UNIT' }]
  ),
  // Wing HQ duty still recorded under the pre-ICL title.
  '100004': member(
    { capsn: '100004', orgid: '2500', group: '900', email: 'devon.stale@example.org' },
    [{ id: 'Recruiting & Retention Officer', charter: 'PCR-CA-001' }]
  ),
  // No Workspace account yet.
  '100005': member(
    { capsn: '100005', orgid: '2500', group: '900', email: null },
    [{ id: 'Recruiting Officer', charter: 'PCR-CA-001' }]
  )
};

/** The real normalizer from UpdateMembers.gs, which shares this Apps Script scope. */
function formatDutyTitle_(dutyId) {
  const overrides = {
    'RECRUITING & RETENTION OFFICER': 'Recruiting Officer',
    'RECRUITING AND RETENTION OFFICER': 'Recruiting Officer'
  };
  const title = String(dutyId || '').trim().replace(/\s+/g, ' ');
  return overrides[title.toUpperCase()] || title;
}

function load(globals) {
  return loadModule(MODULE, Object.assign({
    Logger: makeLogger().logger,
    CONFIG: CONFIG
  }, globals || {}), ['getGroupMembers']);
}

const withOverrides = load({ formatDutyTitle_: formatDutyTitle_ });
const withoutOverrides = load({});

/** Sorted member addresses of one generated group id. */
function membersOf(generated, groupId) {
  return Object.keys(generated[groupId] || {}).sort();
}

// ---------------------------------------------------------------------------
section('1. dutyPositionIdsWingHQ matches the org the DUTY is held at');
{
  const generated = withOverrides.getGroupMembers(
    'dty.director-recruiting', 'dutyPositionIdsWingHQ', 'Recruiting Officer', members, squadrons
  );

  check('the primary is in, though they are a squadron member',
    membersOf(generated, 'ca.dty.director-recruiting').includes('alex.primary@example.org'), true);
  check('the assistant is in too',
    membersOf(generated, 'ca.dty.director-recruiting').includes('blair.assistant@example.org'), true);
  check('a Wing HQ member holding the title at a squadron is out',
    membersOf(generated, 'ca.dty.director-recruiting').includes('casey.squadron@example.org'), false);
  check('no per-group child DLs',
    Object.keys(generated), ['ca.dty.director-recruiting']);
}

// ---------------------------------------------------------------------------
section('2. Retired duty titles still match the current one');
{
  const generated = withOverrides.getGroupMembers(
    'dty.director-recruiting', 'dutyPositionIdsWingHQ', 'Recruiting Officer', members, squadrons
  );
  check('holder recorded as "Recruiting & Retention Officer" is included',
    membersOf(generated, 'ca.dty.director-recruiting').includes('devon.stale@example.org'), true);

  check('full membership, deterministic',
    membersOf(generated, 'ca.dty.director-recruiting'),
    ['alex.primary@example.org', 'blair.assistant@example.org', 'devon.stale@example.org']);
}

// ---------------------------------------------------------------------------
section('3. A member with no Workspace account contributes nothing');
{
  const generated = withOverrides.getGroupMembers(
    'dty.director-recruiting', 'dutyPositionIdsWingHQ', 'Recruiting Officer', members, squadrons
  );
  check('accountless holder absent',
    membersOf(generated, 'ca.dty.director-recruiting').some(e => e === null || e === 'null'), false);
}

// ---------------------------------------------------------------------------
section('4. Manual Members: no charter to read, so the duty is held at their own org');
{
  // loadManualMembers() builds dutyPositions as { id, assistant } — no `value`,
  // so there is no charter to parse and the member's own org has to stand in.
  const manual = {
    '200001': {
      capsn: '200001', orgid: '188', group: '', email: 'quinn.manual@example.org', type: 'SENIOR',
      dutyPositions: [{ id: 'Recruiting Officer', assistant: false }],
      dutyPositionIds: ['Recruiting Officer']
    },
    '200002': {
      capsn: '200002', orgid: '2500', group: '900', email: 'rowan.manual@example.org', type: 'SENIOR',
      dutyPositions: [{ id: 'Recruiting Officer', assistant: false }],
      dutyPositionIds: ['Recruiting Officer']
    }
  };

  const generated = withOverrides.getGroupMembers(
    'dty.director-recruiting', 'dutyPositionIdsWingHQ', 'Recruiting Officer', manual, squadrons
  );

  check('manual member attached to Wing HQ is in',
    membersOf(generated, 'ca.dty.director-recruiting'), ['quinn.manual@example.org']);
  check('manual member attached to a squadron is not',
    membersOf(generated, 'ca.dty.director-recruiting').includes('rowan.manual@example.org'), false);
}

// ---------------------------------------------------------------------------
section('5. A row matching nobody creates no group at all');
{
  const generated = withOverrides.getGroupMembers(
    'dty.director-of-recruiting', 'dutyPositionIdsWingHQ', 'Director of Recruiting', members, squadrons
  );
  check('nothing generated', Object.keys(generated), []);
}

// ---------------------------------------------------------------------------
section('6. Title normalization is the only thing that changed for plain dutyPositionIds');
{
  const generated = withoutOverrides.getGroupMembers(
    'dty.recruiting', 'dutyPositionIds', 'Recruiting Officer', members, squadrons
  );

  // Wing-wide list: every holder of the title anywhere in the wing, which is a
  // different question from "who holds it at Wing HQ" and stays that way.
  check('wing-level list holds every titled member with an account',
    membersOf(generated, 'ca.dty.recruiting'),
    ['alex.primary@example.org', 'blair.assistant@example.org', 'casey.squadron@example.org']);

  check('group-echelon child DL built from the members parent org',
    membersOf(generated, 'ca010.dty.recruiting'), ['alex.primary@example.org']);

  check('without the overrides in scope, the retired title does not match',
    membersOf(generated, 'ca.dty.recruiting').includes('devon.stale@example.org'), false);
}

// ---------------------------------------------------------------------------
section('7. An unrecognized Attribute creates nothing');
{
  const generated = withOverrides.getGroupMembers(
    'dty.typo', 'dutyPositionIdsWingHq', 'Recruiting Officer', members, squadrons
  );
  check('no empty group left behind by a typo', Object.keys(generated), []);
}

done();
