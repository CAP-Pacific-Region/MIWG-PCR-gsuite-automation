/**
 * UpdateMembers.gs — buildSendAsDisplayName_(), the Gmail/Directory Send-As name
 * builder shared by addOrUpdateUser() and updateAllSendAsNames(). Chaplain-corps
 * duty holders get "Ch" inserted before their grade ("Last, First M Ch Grade"),
 * per CAP naming convention — this was previously missing from both call sites.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'UpdateMembers.gs');
const { section, check, done } = makeChecker();

function load() {
  return loadModule(MODULE, {
    Logger: makeLogger().logger
  }, ['buildSendAsDisplayName_']);
}

function member(overrides) {
  return Object.assign({
    lastName: 'Smith',
    firstName: 'Taylor',
    middleName: '',
    suffix: '',
    rank: 'Capt',
    dutyPositions: []
  }, overrides);
}

// ---------------------------------------------------------------------------
section('buildSendAsDisplayName_ — non-chaplains are unaffected');
{
  const { buildSendAsDisplayName_ } = load();

  check('basic name, no middle name, no duty positions',
    buildSendAsDisplayName_(member({})),
    'Smith, Taylor Capt');

  check('middle name is initialed',
    buildSendAsDisplayName_(member({ middleName: 'Alex' })),
    'Smith, Taylor A Capt');

  check('suffix is appended to the last name',
    buildSendAsDisplayName_(member({ suffix: 'Jr' })),
    'Smith Jr, Taylor Capt');

  check('a non-chaplain duty position does not add "Ch"',
    buildSendAsDisplayName_(member({
      dutyPositions: [{ id: 'Personnel Officer' }]
    })),
    'Smith, Taylor Capt');
}

// ---------------------------------------------------------------------------
section('buildSendAsDisplayName_ — chaplain-corps duty holders get "Ch"');
{
  const { buildSendAsDisplayName_ } = load();

  check('"Chaplain" duty position inserts Ch before grade',
    buildSendAsDisplayName_(member({
      dutyPositions: [{ id: 'Chaplain' }]
    })),
    'Smith, Taylor Ch Capt');

  check('a qualifying duty title like "Deputy Wing Chaplain" still matches',
    buildSendAsDisplayName_(member({
      dutyPositions: [{ id: 'Deputy Wing Chaplain' }]
    })),
    'Smith, Taylor Ch Capt');

  check('matches case-insensitively',
    buildSendAsDisplayName_(member({
      dutyPositions: [{ id: 'squadron chaplain' }]
    })),
    'Smith, Taylor Ch Capt');

  check('Ch is inserted alongside a middle initial and suffix',
    buildSendAsDisplayName_(member({
      middleName: 'Alex',
      suffix: 'Jr',
      dutyPositions: [{ id: 'Chaplain' }]
    })),
    'Smith Jr, Taylor A Ch Capt');

  check('one non-chaplain duty among several is enough to NOT match alone, but any chaplain duty is enough to match',
    buildSendAsDisplayName_(member({
      dutyPositions: [{ id: 'Personnel Officer' }, { id: 'Chaplain' }]
    })),
    'Smith, Taylor Ch Capt');
}

done();
