/**
 * UpdateMembers.gs — loadLevel1CompletedCapids(), the Level I gate for new senior
 * accounts. A real member's legacy Level I completion (2009, pre-dating the modern
 * achievement-tracking cutover) shows on eServices' profile LEVEL tab but has no
 * corresponding ACTIVE row in MbrAchievements — that legacy completion lives only in
 * SeniorLevel.txt (CAPID, Lvl, Completed), the table eServices actually reads for
 * that tab. Pinned here with synthetic CAPIDs — mirrors the CPP Basic/Advanced
 * dual-table gap this codebase already handles in ManageLicenses.gs.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'UpdateMembers.gs');
const { section, check, done } = makeChecker();

/** A fake parseFile(fileName) returning canned rows for 'MbrAchievements' / 'SeniorLevel'. */
function makeParseFile(tables) {
  return fileName => tables[fileName] || [];
}

function load(tables) {
  return loadModule(MODULE, {
    Logger: makeLogger().logger,
    parseFile: makeParseFile(tables)
  }, ['loadLevel1CompletedCapids']);
}

// ---------------------------------------------------------------------------
section('loadLevel1CompletedCapids — checks BOTH MbrAchievements and SeniorLevel.txt');
{
  check('MbrAchievements AchvID 96 ACTIVE counts',
    Array.from(load({
      MbrAchievements: [['111111', '96', 'ACTIVE']]
    }).loadLevel1CompletedCapids()),
    ['111111']);

  check('MbrAchievements AchvID 96 EXPIRED does NOT count',
    Array.from(load({
      MbrAchievements: [['111111', '96', 'EXPIRED']]
    }).loadLevel1CompletedCapids()),
    []);

  check('a different AchvID does NOT count, even ACTIVE',
    Array.from(load({
      MbrAchievements: [['111111', '55', 'ACTIVE']]
    }).loadLevel1CompletedCapids()),
    []);

  // The actual gap: no MbrAchievements row at all, but SeniorLevel.txt has LV1 — the
  // legacy-completion case that surfaced this.
  check('SeniorLevel.txt LV1 row counts on its own',
    Array.from(load({
      MbrAchievements: [],
      SeniorLevel: [['999001', 'LV1', '06/30/2009', 'converted', '07/08/2009']]
    }).loadLevel1CompletedCapids()),
    ['999001']);

  check('SeniorLevel.txt LV2/LV3/etc. do NOT count as Level I',
    Array.from(load({
      SeniorLevel: [['222222', 'LV2', '04/22/2013']]
    }).loadLevel1CompletedCapids()),
    []);

  check('a CAPID complete in EITHER table is not double-counted (Set, not array)',
    Array.from(load({
      MbrAchievements: [['444444', '96', 'ACTIVE']],
      SeniorLevel: [['444444', 'LV1', '01/01/2010']]
    }).loadLevel1CompletedCapids()),
    ['444444']);

  check('two different CAPIDs, one per table, both come through',
    Array.from(load({
      MbrAchievements: [['555555', '96', 'ACTIVE']],
      SeniorLevel: [['666666', 'LV1', '01/01/2010']]
    }).loadLevel1CompletedCapids()).sort(),
    ['555555', '666666']);

  check('a member with LV2/LV3 but no LV1 row is NOT counted just because higher levels exist',
    Array.from(load({
      SeniorLevel: [
        ['777777', 'LV2', '04/22/2013'],
        ['777777', 'LV3', '08/24/2020']
      ]
    }).loadLevel1CompletedCapids()),
    []);

  check('empty tables yield an empty set',
    Array.from(load({}).loadLevel1CompletedCapids()),
    []);
}

done();
