/**
 * RegionStaffMembers.gs - provisioning side: mergeDerivedRegionStaff_() and the
 * lifecycle hooks (expirations, fail-safe, Level I). Synthetic CAPIDs only.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'RegionStaffMembers.gs');
const { section, check, done } = makeChecker();

const squadrons = { '100': { charter: 'PCR-CA-070', name: 'Alpha' } };

function memberRow(capid, orgid, type, status) {
  const r = new Array(25).fill('');
  r[0] = capid; r[2] = 'L' + capid; r[3] = 'F' + capid; r[11] = orgid; r[14] = 'Maj';
  r[16] = '12/31/2027'; r[21] = type || 'SENIOR'; r[24] = status || 'ACTIVE';
  return r;
}
function dutyRow(capid, orgid) {
  const r = new Array(8).fill('');
  r[0] = capid; r[1] = 'IT Officer'; r[4] = '0'; r[7] = orgid;
  return r;
}

/** opts: { flag, regionMembers, orgPaths } */
function load(opts) {
  const calls = { contact: [], duty: [] };
  const region = {
    Member: opts.regionMembers || [],
    DutyPosition: opts.regionDuties || [],
    MbrContact: [['1', 'EMAIL', 'PRIMARY', 'a@example.org']],
    Organization: [['434', 'PCR', 'PCR', '1', '', 'Pacific Region']],
    MbrAchievements: [['1', '96', 'ACTIVE'], ['2', '96', 'EXPIRED']],
    SeniorLevel: [['3', 'LV1']]
  };
  const mod = loadModule(MODULE, {
    Logger: makeLogger().logger,
    Utilities: { formatString: (f, a, b, c) => a + '-' + b + '-' + String(c).padStart(3, '0') },
    CONFIG: {
      DERIVE_REGION_STAFF: !!opts.flag,
      REGION_CAPWATCH_DATA_FOLDER_ID: 'folder',
      MEMBER_TYPES: { ACTIVE: ['SENIOR'] },
      EXCLUDED_ORG_IDS: []
    },
    parseCapwatchFromFolder_: (id, name) => region[name] || [],
    parseFile: name => name === 'OrgPaths' ? (opts.orgPaths || [['434', '/PCR-001']]) : [],
    getSquadrons: () => squadrons,
    addContactInfo: (m, rows) => calls.contact.push(Object.keys(m)),
    addDutyPositions: (m, rows) => calls.duty.push(Object.keys(m))
  }, ['mergeDerivedRegionStaff_', 'derivedRegionStaffExpirations_', 'regionExtractReadable_',
      'addRegionLevel1Completions_', 'addDerivedRegionStaffCapids_']);
  mod.calls = calls;
  return mod;
}

const one = { regionMembers: [memberRow('1', '434')], regionDuties: [dutyRow('1', '100')] };

section('flag off');
{
  const m = load(Object.assign({ flag: false }, one));
  const roster = {};
  check('nothing merged', m.mergeDerivedRegionStaff_(roster, squadrons), 0);
  check('roster untouched', Object.keys(roster), []);
  check('no expirations', m.derivedRegionStaffExpirations_(), {});
  check('region extract counts as readable', m.regionExtractReadable_(), true);
  check('Level I set untouched', Array.from(m.addRegionLevel1Completions_(new Set())), []);
}

section('flag on');
{
  const m = load(Object.assign({ flag: true }, one));
  const roster = {};
  check('one member merged', m.mergeDerivedRegionStaff_(roster, squadrons), 1);
  check('placed in the HQ OU', roster['1'].orgPath, '/PCR-001');
  check('home org is HQ', roster['1'].orgid, '434');
  check('marked as derived', roster['1'].derivedFromRegion, true);
  check('contact reader fed only the added member', m.calls.contact, [['1']]);
  check('duty reader fed only the added member', m.calls.duty, [['1']]);
  check('counts as active for the lifecycle', m.derivedRegionStaffExpirations_(), { '1': '12/31/2027' });
  check('readable', m.regionExtractReadable_(), true);
  const lv = Array.from(m.addRegionLevel1Completions_(new Set(['9']))).sort();
  check('region Level I completions added (ACTIVE 96 and LV1 only)', lv, ['1', '3', '9']);
  const allow = m.addDerivedRegionStaffCapids_({ '7': 1 });
  check('allow-list gains derived CAPIDs', Object.keys(allow).sort(), ['1', '7']);
}

section('a wing-roster member is never overwritten');
{
  const m = load(Object.assign({ flag: true }, one));
  const wingCopy = { capsn: '1', orgid: '100' };
  const roster = { '1': wingCopy };
  check('nothing added', m.mergeDerivedRegionStaff_(roster, squadrons), 0);
  check('wing entry kept', roster['1'] === wingCopy, true);
}

section('fail-safe: unreadable region extract');
{
  const m = load({ flag: true, regionMembers: [], regionDuties: [] });
  check('reported unreadable', m.regionExtractReadable_(), false);
  check('merge adds nothing', m.mergeDerivedRegionStaff_({}, squadrons), 0);
  check('no expirations', m.derivedRegionStaffExpirations_(), {});
}
{
  const m = load(Object.assign({ flag: true, orgPaths: [] }, one));
  check('missing HQ OU path is also unreadable (cannot place them)', m.regionExtractReadable_(), false);
}
done();
