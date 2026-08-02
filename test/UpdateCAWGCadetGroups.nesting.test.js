/**
 * UpdateCAWGCadetGroups.gs — WHICH address gets nested, and WHETHER the row it
 * lands in is allowed to hold it. Both were wrong, in ways that cancelled out into
 * silence.
 *
 * WHICH (v1.3.0). buildCAWGCadetSourceGroupEmail_ took `scope` and ignored it, so
 * a wing target asked for ca.cadets@cawgcadets.org. Nothing creates that: `.cadets`
 * groups on the cadet tenant come from updateAllSquadronGroups(), which walks UNIT
 * scope only, and no CAPWATCH org is wing scope. The add 404s and is swallowed as
 * "cannot add external member". Worse, the function's managed-address pattern
 * matches `.all@` while it generated none — so it deleted the hand-added row for
 * ca.all@cawgcadets.org, the address that actually works, on every run.
 *
 * WHETHER (v1.2.0). It set "Add EXT" on the rows it creates but not on the ".all"
 * rows it nests those groups into, which a human writes. A blank column there is
 * not inert: updateEmailGroups() reads it as "this group holds no outside
 * addresses" and both declines the add and writes allowExternalMembers=false onto
 * the group, every run. So the sheet asked for a nesting it simultaneously forbade,
 * and an Admin-console repair reverted by morning.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'UpdateCAWGCadetGroups.gs');
const { section, check, done } = makeChecker();
const { logger } = makeLogger();

/**
 * Squadrons as getSquadrons() returns them: a wing org, a GROUP echelon, and one
 * unit whose nextLevel points at that group.
 */
const SQUADRONS = {
  '100': { scope: 'WING', wing: 'ca', unit: '001', name: 'CALIFORNIA WING', nextLevel: '' },
  '200': { scope: 'GROUP', wing: 'ca', unit: '205', name: 'GROUP 3', nextLevel: '100' },
  '300': { scope: 'UNIT', wing: 'ca', unit: '007', name: 'SQUADRON 7', nextLevel: '200' }
};

/** One ACTIVE CADET in orgid 300 — Member.txt columns 11 / 21 / 24. */
function memberRow(orgid) {
  const row = new Array(25).fill('');
  row[11] = orgid;
  row[21] = 'CADET';
  row[24] = 'ACTIVE';
  return row;
}

const m = loadModule(MODULE, {
  Logger: logger,
  CONFIG: { WING: 'CA', AUTOMATION_SPREADSHEET_ID: 'x', CADETS_TENANT_DOMAIN: 'cawgcadets.org' },
  getSquadrons: () => SQUADRONS,
  parseFile: name => (name === 'Member' ? [memberRow('300')] : [])
}, ['collectCAWGCadetExternalNestTargets_', 'upsertCAWGCadetGroupDefinitionRows_',
    'buildCAWGCadetSourceGroupEmail_', 'isManagedCAWGCadetGroupEmail_',
    'buildCAWGCadetManagedRows_']);

/**
 * A stand-in for the Groups tab that records what was written to it.
 *
 * @param {Array<Array<*>>} rows - starting contents, header first
 * @returns {Object} sheet with a .written property
 */
function fakeSheet(rows) {
  const sheet = {
    written: null,
    getDataRange: () => ({ getValues: () => rows.map(r => r.slice()) }),
    clear: () => {},
    setFrozenRows: () => {},
    autoResizeColumns: () => {},
    getRange: (_r, _c, numRows, numCols) => ({
      setValues: v => { sheet.written = v.map(row => row.slice(0, numCols)).slice(0, numRows); }
    })
  };
  return sheet;
}

const HEADER = ['Category', 'Group Name', 'Attribute', 'Values', 'Description', 'Add EXT'];

/** Finds a written row by its Group Name. */
function rowFor(sheet, name) {
  return (sheet.written || []).find(r => String(r[1]).toLowerCase() === name);
}

// ---------------------------------------------------------------------------
section('1. Which row names a nesting lands in');
{
  const targets = m.collectCAWGCadetExternalNestTargets_([
    { email: 'ca.cadets@cawgcadets.org', groups: 'ca.cadets,ca.all' },
    { email: 'ca007.cadets@cawgcadets.org', groups: 'ca007.cadets,ca007.all,ca006.cadets,ca006.all' },
    { email: 'ca007.parents@cawgcadets.org', groups: 'ca007.parents' }
  ]);

  // The Groups tab names a row once for every echelon that uses it, so the wing
  // list and a unit list are both the row named "all".
  check('wing and unit .all collapse to one name', targets.indexOf('all') > -1, true);
  check('the rows it creates are targets too', targets.indexOf('cadets') > -1, true);
  check('parents included', targets.indexOf('parents') > -1, true);
  check('exactly those three', targets.length, 3);
  check('sorted', targets.join(','), 'all,cadets,parents');
}

// ---------------------------------------------------------------------------
section('2. A blank Add EXT on a nest target is filled in');
{
  const sheet = fakeSheet([
    HEADER,
    ['membership', 'all', 'type', 'CADET,SENIOR', 'All members', ''],
    ['achievements', 'achievements', 'achievements', 'MITCHELL', 'Achievements', '']
  ]);

  const res = m.upsertCAWGCadetGroupDefinitionRows_(sheet, [], ['all']);

  check('the nest target is stamped', rowFor(sheet, 'all')[5], 'Y');
  check('and reported', res.addExtStamped.join(','), 'all');
  check('a row nothing nests into is left alone', rowFor(sheet, 'achievements')[5], '');
}

// ---------------------------------------------------------------------------
section('3. Nothing else on a human-owned row is touched');
{
  const sheet = fakeSheet([
    HEADER,
    ['membership', 'all', 'type', 'CADET,SENIOR', 'All members', '']
  ]);

  m.upsertCAWGCadetGroupDefinitionRows_(sheet, [], ['all']);
  const row = rowFor(sheet, 'all');

  // The whole value of stamping one cell is that the rest of the row stays the
  // author's. Rewriting Attribute or Values here would silently redefine who is
  // on the wing's all-hands list.
  check('Category kept', row[0], 'membership');
  check('Attribute kept', row[2], 'type');
  check('Values kept', row[3], 'CADET,SENIOR');
  check('Description kept', row[4], 'All members');
}

// ---------------------------------------------------------------------------
section('4. An existing permission is not rewritten');
{
  ['Y', 'yes', 'x', 'TRUE'].forEach(existing => {
    const sheet = fakeSheet([HEADER, ['membership', 'all', 'type', '', '', existing]]);
    const res = m.upsertCAWGCadetGroupDefinitionRows_(sheet, [], ['all']);
    check('"' + existing + '" is left as written', rowFor(sheet, 'all')[5], existing);
    check('"' + existing + '" is not reported as stamped', res.addExtStamped.length, 0);
  });

  // "Add Lite" implies external members in UpdateGroups, but it lives in its own
  // column; a row carrying only that still reads as blank here and is stamped.
  // Harmless — the two say the same thing — and it keeps this function reading
  // exactly one column.
  const sheet = fakeSheet([HEADER, ['membership', 'all', 'type', '', '', 'n']]);
  m.upsertCAWGCadetGroupDefinitionRows_(sheet, [], ['all']);
  check('an explicit "n" is overridden, because the nesting requires it', rowFor(sheet, 'all')[5], 'Y');
}

// ---------------------------------------------------------------------------
section('5. Rows this function owns outright still win');
{
  const sheet = fakeSheet([
    HEADER,
    ['membership', 'all', 'type', 'CADET,SENIOR', 'All members', ''],
    ['custom', 'cadets', 'stale', 'nonsense', 'Old', '']
  ]);

  const res = m.upsertCAWGCadetGroupDefinitionRows_(sheet, [
    { category: 'custom', groupName: 'cadets', attribute: 'manualOnly',
      values: 'ca007.cadets', description: 'Cadets', addExt: 'Y' }
  ], ['all', 'cadets']);

  const cadets = rowFor(sheet, 'cadets');
  check('the managed row is replaced, not stamped in place', cadets[2], 'manualOnly');
  check('and carries its own Add EXT', cadets[5], 'Y');
  check('only the human row was stamped', res.addExtStamped.join(','), 'all');
  check('one cadets row, not two', (sheet.written || []).filter(r => r[1] === 'cadets').length, 1);
}

// ---------------------------------------------------------------------------
section('6. No targets means no changes');
{
  const sheet = fakeSheet([HEADER, ['membership', 'all', 'type', '', '', '']]);

  const res = m.upsertCAWGCadetGroupDefinitionRows_(sheet, [], []);
  check('nothing stamped', res.addExtStamped.length, 0);
  check('row untouched', rowFor(sheet, 'all')[5], '');

  const sheet2 = fakeSheet([HEADER, ['membership', 'all', 'type', '', '', '']]);
  const res2 = m.upsertCAWGCadetGroupDefinitionRows_(sheet2, []);
  check('an omitted argument is not an error', res2.addExtStamped.length, 0);
  check('and changes nothing', rowFor(sheet2, 'all')[5], '');
}

// ---------------------------------------------------------------------------
section('7. Which cadet-tenant address a target nests');
{
  // scope was a parameter this function took and ignored, so a wing target asked
  // for ca.cadets@cawgcadets.org — an address nothing creates. `.cadets` groups
  // on that tenant come from updateAllSquadronGroups(), which walks UNIT scope
  // only, and no CAPWATCH org is wing scope. Verified 2026-07-31: that address
  // 404s; ca.all@cawgcadets.org holds 2,646 members and is named "CAWG - Cadets".
  const src = m.buildCAWGCadetSourceGroupEmail_;

  check('a unit still nests its .cadets group',
    src('ca007', 'UNIT', 'cadets', 'cawgcadets.org'), 'ca007.cadets@cawgcadets.org');
  check('the wing nests the all-hands that exists',
    src('ca', 'WING', 'cadets', 'cawgcadets.org'), 'ca.all@cawgcadets.org');
  check('scope is read case-insensitively',
    src('ca', 'wing', 'cadets', 'cawgcadets.org'), 'ca.all@cawgcadets.org');
  check('an unknown scope keeps the unit shape, not the wing one',
    src('ca006', '', 'cadets', 'cawgcadets.org'), 'ca006.cadets@cawgcadets.org');

  // Parents was left alone on purpose: whether a wing-level parents group exists
  // on the cadet tenant is unestablished, and this tenant cannot check.
  check('wing parents is unchanged',
    src('ca', 'WING', 'parents', 'cawgcadets.org'), 'ca.parents@cawgcadets.org');
  check('unit parents is unchanged',
    src('ca007', 'UNIT', 'parents', 'cawgcadets.org'), 'ca007.parents@cawgcadets.org');

  let threw = false;
  try { src('ca', 'WING', 'nonsense', 'cawgcadets.org'); } catch (e) { threw = true; }
  check('an unknown kind still throws', threw, true);
}

// ---------------------------------------------------------------------------
section('8. The wing row is now one this function owns');
{
  // The hand-added ca.all@cawgcadets.org row was being DELETED every run: the
  // managed-address pattern matches `.all@` while nothing generated one, so
  // upsertCAWGCadetGroupRows_ dropped it as an orphan of this function's own
  // making. Generating it is what stops that, so the pattern and the generator
  // have to agree about this address.
  check('the address it now generates is one it claims to manage',
    m.isManagedCAWGCadetGroupEmail_('ca.all@cawgcadets.org', 'cawgcadets.org'), true);
  check('and so is a unit source',
    m.isManagedCAWGCadetGroupEmail_('ca007.cadets@cawgcadets.org', 'cawgcadets.org'), true);
  check('an unrelated cadet-tenant address is not claimed',
    m.isManagedCAWGCadetGroupEmail_('somebody@cawgcadets.org', 'cawgcadets.org'), false);
  check('a wing-tenant address is never claimed',
    m.isManagedCAWGCadetGroupEmail_('ca.all@cawgcap.org', 'cawgcadets.org'), false);
}

// ---------------------------------------------------------------------------
section('9. Cadets and parents reach the wing by different routes');
{
  // Neither ca.cadets@ nor ca.parents@ exists on the cadet tenant — nothing there
  // creates wing-level groups. Cadets have a substitute (that tenant's own
  // all-hands IS its cadets); parents do not, so the units feed the wing list.
  const rows = m.buildCAWGCadetManagedRows_('cawgcadets.org').userAdditionsRows;
  const by = email => rows.find(r => r.email === email);

  check('the wing cadet source is the all-hands that exists',
    !!by('ca.all@cawgcadets.org'), true);
  check('and it feeds both wing lists',
    by('ca.all@cawgcadets.org').groups, 'ca.cadets,ca.all');

  check('no wing cadets source row — that address 404s',
    by('ca.cadets@cawgcadets.org'), undefined);
  check('no wing parents source row either — same reason',
    by('ca.parents@cawgcadets.org'), undefined);

  const unitParents = by('ca007.parents@cawgcadets.org');
  check('a unit feeds its own parents list', unitParents.groups.indexOf('ca007.parents') > -1, true);
  check('and its group echelon', unitParents.groups.indexOf('ca205.parents') > -1, true);
  check('and now the wing list directly', unitParents.groups.indexOf('ca.parents') > -1, true);

  // Cadets deliberately do NOT reach up: the wing row already nests one aggregate
  // covering every cadet, so adding them here would duplicate that population.
  const unitCadets = by('ca007.cadets@cawgcadets.org');
  check('a unit does not feed ca.cadets', unitCadets.groups.indexOf('ca.cadets'), -1);
  check('it feeds its own and its group echelon',
    unitCadets.groups, 'ca007.cadets,ca007.all,ca205.cadets,ca205.all');

  check('exactly three rows: wing cadets, unit cadets, unit parents', rows.length, 3);
}

// ---------------------------------------------------------------------------
section('10. ca.parents stays a managed destination');
{
  // Dropping the wing SOURCE row must not drop the wing DESTINATION group, or the
  // 56 unit groups would be nested into something this function stopped managing.
  const defs = m.buildCAWGCadetManagedRows_('cawgcadets.org').groupDefinitions;
  const parents = defs.find(d => d.groupName === 'parents');
  const cadets = defs.find(d => d.groupName === 'cadets');

  check('ca.parents is still in the parents row values',
    parents.values.split(',').indexOf('ca.parents') > -1, true);
  check('alongside the unit destination',
    parents.values.split(',').indexOf('ca007.parents') > -1, true);
  check('ca.cadets likewise', cadets.values.split(',').indexOf('ca.cadets') > -1, true);
  check('both rows still request external members', parents.addExt + cadets.addExt, 'YY');
}

done();
