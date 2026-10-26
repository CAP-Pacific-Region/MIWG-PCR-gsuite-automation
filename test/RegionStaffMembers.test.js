/**
 * RegionStaffMembers.gs — deriveRegionStaffMembers_(): region-HQ members with a duty
 * in one of this wing's units. Synthetic CAPIDs only.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'RegionStaffMembers.gs');
const { section, check, done } = makeChecker();
const m = loadModule(MODULE, { Logger: makeLogger().logger }, ['deriveRegionStaffMembers_', 'explainRegionStaffExclusions_']);

const squadrons = {
  '100': { charter: 'PCR-CA-070' }, '200': { charter: 'PCR-CA-001' }, '1297': { charter: 'PCR-CA-000' }
};
const opts = { types: ['SENIOR', 'FIFTY YEAR'], excludedOrgIds: ['1297'] };

function mem(capid, orgid, type, status) {
  const r = new Array(25).fill('');
  r[0] = capid; r[2] = 'L' + capid; r[3] = 'F' + capid; r[11] = orgid; r[14] = 'Maj';
  r[21] = type || 'SENIOR'; r[24] = status || 'ACTIVE';
  return r;
}
function duty(capid, id, asst, orgid) {
  const r = new Array(8).fill('');
  r[0] = capid; r[1] = id; r[4] = asst ? '1' : '0'; r[7] = orgid;
  return r;
}
const derive = (mems, duties) => m.deriveRegionStaffMembers_(mems, duties, squadrons, opts);

section('qualification');
check('HQ member with a wing duty qualifies',
  Object.keys(derive([mem('1', '434')], [duty('1', 'IT Officer', false, '100')])), ['1']);
check('HQ member with only an HQ duty does not',
  Object.keys(derive([mem('1', '434')], [duty('1', 'Commander', false, '434')])), []);
check('wing-primary member (not HQ) is not derived',
  Object.keys(derive([mem('1', '100')], [duty('1', 'IT Officer', false, '100')])), []);
check('inactive HQ member does not qualify',
  Object.keys(derive([mem('1', '434', 'SENIOR', 'EXPIRED')], [duty('1', 'X', false, '100')])), []);
check('ineligible member type does not qualify',
  Object.keys(derive([mem('1', '434', 'CADET')], [duty('1', 'X', false, '100')])), []);
check('duty in an excluded holding unit does not count',
  Object.keys(derive([mem('1', '434')], [duty('1', 'X', false, '1297')])), []);

section('placement');
{
  const r = derive([mem('1', '434')], [
    duty('1', 'Deputy', true, '200'), duty('1', 'Director', false, '100'), duty('1', 'Chief', false, '200')]);
  check('placed in the region HQ org, not a duty unit', r['1'].placementOrgid, '434');
  check('all wing duties are kept', r['1'].duties.length, 3);
  const a = derive([mem('2', '434')], [duty('2', 'Deputy', true, '100')]);
  check('assistant-only member is still placed at HQ', a['2'].placementOrgid, '434');
}
section('explainRegionStaffExclusions_');
{
  const mems = [mem('1', '100'), mem('2', '434', 'SENIOR', 'EXPIRED'), mem('3', '434', 'CADET'),
                mem('4', '434'), mem('5', '434')];
  const duties = [duty('4', 'X', false, '434'), duty('5', 'Y', false, '100')];
  const why = m.explainRegionStaffExclusions_(['9', '1', '2', '3', '4', '5'], mems, duties, squadrons, opts);
  check('absent from extract', why['9'], 'not in region extract');
  check('wing-primary', why['1'], 'primary org 100 (not HQ)');
  check('inactive', why['2'], 'status EXPIRED');
  check('ineligible type', why['3'], 'type CADET not eligible');
  check('HQ-only duty', why['4'], 'no duty in a wing unit (1 duties elsewhere)');
  check('would qualify', why['5'], 'qualifies (rule should have derived it)');
}
done();
