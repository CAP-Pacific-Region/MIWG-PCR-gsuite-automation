/**
 * signature-webapp/ — the member self-service signature app.
 *
 * WHAT THIS FILE IS FOR
 *
 * The app is a SECOND implementation of a signature src/ already pushes.
 * src/accounts-and-groups/UpdateMembers.gs owns the original; the web app carries
 * a port, because the two live in different Apps Script projects and a project
 * cannot call another's functions without becoming a library with its own
 * deploy-version step.
 *
 * A copy that drifts is worse than no copy at all. A member would approve one
 * signature here and have the next bulk push replace it with a subtly different
 * one — different enough to notice, not different enough to explain.
 *
 * So the first and largest section below runs BOTH generators over the same
 * members and requires byte-identical output. It is the price of the duplication
 * and the reason the duplication is acceptable.
 *
 * The rest pins the decisions the app makes on its own:
 *   - the phone row is the ONLY thing a member may change, and only an explicit
 *     `true` turns it on;
 *   - a cadet never has a phone to publish in the first place;
 *   - a signature is never written to an address the organization does not own.
 *
 * All members here are synthetic.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeDrive, makeChecker, Utilities } = require('./helpers/apps-script');
const { MEMBER_HEADER, COL } = require('./helpers/capwatch-fixtures');

const { section, check, done } = makeChecker();

const SRC_MEMBERS = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'UpdateMembers.gs');
const SRC_UTILS = path.join(__dirname, '..', 'src', 'utils.gs');
const APP = path.join(__dirname, '..', 'signature-webapp');

const WING_FOLDER = 'folder-id';
const REGION_FOLDER = 'region-folder-id';

/**
 * A DriveApp over SEVERAL folders, because the whole point of the region
 * supplement is that the two extracts are different files in different places.
 * An unknown folder id throws, the way Drive does — that is the path a mistyped
 * or unshared folder takes, and both modules must survive it.
 */
const DRIVE_FOLDERS = {};
function makeFolderedDrive(byFolder) {
  return {
    getFolderById: id => {
      const files = byFolder[id];
      if (!files) throw new Error('No item with the given ID could be found: ' + id);
      return {
        getFilesByName: name => {
          const exists = Object.prototype.hasOwnProperty.call(files, name);
          let consumed = false;
          return {
            hasNext: () => exists && !consumed,
            next: () => {
              consumed = true;
              return { getBlob: () => ({ getDataAsString: () => files[name] }) };
            }
          };
        }
      };
    }
  };
}

const CONFIG = {
  WING: 'CA',
  EMAIL_DOMAIN: '@example.org',
  SECONDARY_EMAIL_DOMAIN: '@example.cap.gov',
  CAPWATCH_DATA_FOLDER_ID: WING_FOLDER,
  REGION_CAPWATCH_DATA_FOLDER_ID: REGION_FOLDER
};

const SIG_CONFIG = {
  WING: 'CA',
  EMAIL_DOMAIN: '@example.org',
  SECONDARY_EMAIL_DOMAIN: '@example.cap.gov',
  CAPWATCH_DATA_FOLDER_ID: WING_FOLDER,
  REGION_CAPWATCH_DATA_FOLDER_ID: REGION_FOLDER,
  ALLOWED_GROUP: '',
  SUPPORT_EMAIL: 'it@example.org',
  ORG_LABEL: 'CAWG'
};

// utils.gs declares its own Logger, so that global is NOT injected here.
const utils = loadModule(SRC_UTILS, {
  CONFIG: CONFIG,
  Utilities: Utilities
}, ['toTitleCase', 'sanitizeEmail', 'calculateGroup']);

const src = loadModule(SRC_MEMBERS, {
  CONFIG: CONFIG,
  Logger: makeLogger().logger,
  Utilities: Utilities,
  // Resolves the folder maps at CALL time, so the fixtures below can be declared
  // in reading order rather than hoisted above this load.
  DriveApp: makeFolderedDrive(DRIVE_FOLDERS),
  toTitleCase: utils.toTitleCase,
  sanitizeEmail: utils.sanitizeEmail,
  calculateGroup: utils.calculateGroup
}, ['generateEmailSignature', 'getDutyBlock', 'dutyKey_', 'createMemberObject',
    'addContactInfo', 'addDutyPositions', 'addCadetDutyPositions',
    'addOutOfWingDutyPositions_']);

const tpl = loadModule(path.join(APP, 'SignatureTemplate.gs'), {
  SIG_CONFIG: SIG_CONFIG
}, ['sigGenerateEmailSignature_', 'sigDutyBlock_', 'sigName_', 'sigFormatPhone_',
    'sigDutyKey_', 'sigDutyLevelRank_']);

// ---------------------------------------------------------------------------
// Members, in the shape both generators consume.
// ---------------------------------------------------------------------------

function duty(id, level, orgName, orgScope, assistant) {
  return { id: id, level: level, orgName: orgName, orgScope: orgScope, assistant: !!assistant };
}

const MEMBERS = {
  'graded senior, wing + squadron duty, phone': {
    rank: 'Maj', firstName: 'Rowan', middleName: 'K', lastName: 'Ashford', suffix: '',
    phone: '+15550101234', orgName: 'EXAMPLE SR SQDN 80', type: 'SENIOR',
    dutyPositions: [
      duty('Information Technologies Officer', 'UNIT', 'EXAMPLE SR SQDN 80', 'UNIT'),
      duty('Director of IT', 'WING', 'CALIFORNIA WING HQ', 'WING')
    ]
  },

  'two wing duties must not crowd out a squadron command': {
    rank: 'Lt Col', firstName: 'Wren', middleName: '', lastName: 'Calloway', suffix: '',
    phone: '', orgName: 'EXAMPLE COMP SQ 45', type: 'SENIOR',
    dutyPositions: [
      duty('Web Security Administrator', 'WING', 'CALIFORNIA WING HQ', 'WING'),
      duty('Director of IT', 'WING', 'CALIFORNIA WING HQ', 'WING'),
      duty('Commander', 'UNIT', 'EXAMPLE COMP SQ 45', 'UNIT')
    ]
  },

  'ungraded senior (SM), no duty, no phone': {
    rank: 'SM', firstName: 'Devi', middleName: 'Marlowe', lastName: 'Osei', suffix: '',
    phone: '', orgName: 'EXAMPLE SR SQDN 80', type: 'SENIOR',
    dutyPositions: []
  },

  'cadet with a cadet duty': {
    rank: 'C/SMSgt', firstName: 'Imani', middleName: 'T', lastName: 'Brightwater', suffix: '',
    phone: '', orgName: 'EXAMPLE CDT SQDN 12', type: 'CADET',
    dutyPositions: [duty('Cadet Commander', 'UNIT', 'EXAMPLE CDT SQDN 12', 'UNIT')]
  },

  'assistant duties only — the block is dropped entirely': {
    rank: 'Capt', firstName: 'Soren', middleName: '', lastName: 'Vasquez', suffix: 'Sr',
    phone: '+15550109876', orgName: 'EXAMPLE SR SQDN 80', type: 'SENIOR',
    dutyPositions: [
      duty('Personnel Officer', 'UNIT', 'EXAMPLE SR SQDN 80', 'UNIT', true),
      duty('Safety Officer', 'GROUP', 'CENTRAL CALIF GROUP 6', 'GROUP', true)
    ]
  },

  'a retired duty title still in the feed': {
    rank: '1st Lt', firstName: 'Lior', middleName: '', lastName: 'Nakamura', suffix: '',
    phone: '', orgName: 'EXAMPLE SR SQDN 80', type: 'SENIOR',
    dutyPositions: [duty('Recruiting & Retention Officer', 'UNIT', 'EXAMPLE SR SQDN 80', 'UNIT')]
  },

  'wing HQ org is cut back to the echelon': {
    rank: 'Col', firstName: 'Astrid', middleName: 'J', lastName: "O'Brien", suffix: '',
    phone: '+15550100000', orgName: 'CALIFORNIA WING HQ', type: 'SENIOR',
    dutyPositions: [duty('Commander', 'WING', 'CALIFORNIA WING HQ', 'WING')]
  },

  'duty at an org outside the wing falls back to the home unit': {
    rank: 'Maj', firstName: 'Kai', middleName: '', lastName: 'Delacroix-Stone', suffix: 'Jr',
    phone: '+15550105555', orgName: 'FALLBROOK SENIOR SQD 87', type: 'SENIOR',
    // orgName '' is what a duty held outside this wing looks like: getSquadrons()
    // only maps this wing's orgs, and MemberRecord.gs mirrors that filter.
    dutyPositions: [duty('Director of Operations', 'REGION', '', '')]
  }
};

// ---------------------------------------------------------------------------
section('1. The port renders EXACTLY what the main project renders');
{
  Object.keys(MEMBERS).forEach(label => {
    const member = MEMBERS[label];
    check(label, tpl.sigGenerateEmailSignature_(member), src.generateEmailSignature(member));
  });

  // A guard on the guard: if the fixtures were all trivially equal (say both
  // generators returned ''), everything above would pass vacuously.
  const rendered = tpl.sigGenerateEmailSignature_(MEMBERS['graded senior, wing + squadron duty, phone']);
  check('the compared output is a real signature — the name line is there',
    rendered.indexOf('Maj. Rowan Ashford') !== -1, true);
  check('...with the duty block, highest echelon first',
    rendered.indexOf('California Wing Director of IT<br />Example Senior Squadron 80 Information Technologies Officer') !== -1,
    true);
  check('...and the phone row', rendered.indexOf('555.010.1234') !== -1, true);

  // Vertical rhythm. Both of these shipped wrong and were only visible once a real
  // signature sat in a real mail client, which is a slow way to find out.
  check('nothing precedes the name line — no leading blank line in the mail client',
    /^\s*<!DOCTYPE html>\s*<html>\s*<body>\s*<h1\b/.test(rendered), true);
  check('and the logo sits on the same 5px rhythm as every other row, not 20',
    rendered.indexOf('height:42px; margin: 15px 0 5px 0;') !== -1, true);
  check('...which is the ONLY bottom margin in the block',
    (rendered.match(/margin: 0 0 5px;|margin: 15px 0 5px 0;/g) || []).length,
    (rendered.match(/margin:/g) || []).length);
}

// ---------------------------------------------------------------------------
section('2. Suppressing the phone changes the phone and nothing else');
{
  const member = MEMBERS['graded senior, wing + squadron duty, phone'];
  const withPhone = tpl.sigGenerateEmailSignature_(member);
  const without = tpl.sigGenerateEmailSignature_(Object.assign({}, member, { phone: '' }));

  check('the number is gone', without.indexOf('555.010.1234') === -1, true);
  check('the tel: link is gone', without.indexOf('tel:+1') === -1, true);
  check('the name line survives', without.indexOf('Maj. Rowan Ashford') !== -1, true);
  check('the duty block survives', without.indexOf('Director of IT') !== -1, true);
  // Pull the href out and compare it WHOLE rather than asking whether the host
  // appears somewhere in the document. A substring check would pass just as
  // happily on https://cawg.cap.gov.example.com, and it is a signature: the link
  // is the wing's or it is somebody else's.
  const wingHref = (without.match(/<a\s+href="(https:\/\/[^"]*)"/g) || [])
    .map(tag => tag.replace(/^<a\s+href="/, '').replace(/"$/, ''))
    .filter(href => href !== 'https://www.GoCivilAirPatrol.com');
  check('the wing website line survives', wingHref, ['https://cawg.cap.gov']);

  // Line-level set difference, so "only the phone row differs" is asserted rather
  // than eyeballed. Duplicate lines are matched one-for-one: the three paragraphs
  // share an opening line, and a naive contains-check would hide a lost one.
  function only(a, b) {
    const rest = b.slice();
    return a.filter(line => {
      const at = rest.indexOf(line);
      if (at === -1) return true;
      rest.splice(at, 1);
      return false;
    });
  }
  const nonBlank = lines => lines.filter(l => l.trim() !== '');
  const added = only(withPhone.split('\n'), without.split('\n'));
  const lost = only(without.split('\n'), withPhone.split('\n'));

  check('the phone row is the only thing the number adds', nonBlank(added).length, 5);
  check('...and one of those lines is the number itself',
    nonBlank(added).filter(l => l.indexOf('(M)') !== -1).length, 1);
  check('nothing else in the block is lost with it', nonBlank(lost), []);
}

// ---------------------------------------------------------------------------
section('3. The phone toggle is the ONLY thing the client can decide');
{
  const record = {
    capid: '123456',
    member: MEMBERS['graded senior, wing + squadron duty, phone']
  };

  /** SignatureApi.gs over an injectable apply result. */
  function loadApi(applyResult, resolved) {
    return loadModule(path.join(APP, 'SignatureApi.gs'), {
      SIG_CONFIG: SIG_CONFIG,
      Logger: makeLogger().logger,
      HtmlService: {},
      Session: { getActiveUser: () => ({ getEmail: () => 'rowan.ashford@example.org' }) },
      resolveActor_: () => 'rowan.ashford@example.org',
      mayUseSignatureApp_: () => true,
      requireMember_: () => 'rowan.ashford@example.org',
      sigResolveMemberForActor_: () => resolved || record,
      SIG_MAX_DUTIES: 2,
      sigGenerateEmailSignature_: tpl.sigGenerateEmailSignature_,
      sigName_: tpl.sigName_,
      sigDutyBlock_: tpl.sigDutyBlock_,
      sigDutyKey_: tpl.sigDutyKey_,
      sigDutyLevelRank_: tpl.sigDutyLevelRank_,
      sigFormatPhone_: tpl.sigFormatPhone_,
      sigSendAsSnapshot_: () => ({ orgOwned: ['rowan.ashford@example.org'] }),
      sigApplyToSendAsIdentities_: () => applyResult,
      sigSupportSentence_: () => 'Contact it@example.org.'
    }, ['sigRender_', 'apiApply', 'apiGetState']);
  }

  const api = loadApi({ updated: [], skipped: [], failed: [] });

  check('true includes the phone', api.sigRender_('a@example.org', true).includedPhone, true);
  check('false suppresses it', api.sigRender_('a@example.org', false).includedPhone, false);

  // Anything that is not an explicit true fails toward LESS disclosure. A garbled
  // value must never be the reason a personal number gets published.
  ['true', 1, 'yes', {}, [], null, undefined].forEach(value => {
    check('a non-boolean (' + JSON.stringify(value) + ') suppresses it',
      api.sigRender_('a@example.org', value).includedPhone, false);
  });

  check('and the suppressed render really has no number',
    api.sigRender_('a@example.org', 'true').html.indexOf('555.010.1234'), -1);

  // A member with no number on file cannot be reported as having published one.
  const noPhone = { capid: '123457', member: MEMBERS['ungraded senior (SM), no duty, no phone'] };
  const apiNoPhone = loadApi({ updated: ['x@example.org'], skipped: [], failed: [] }, noPhone);
  check('asking for a phone nobody has reports no phone',
    apiNoPhone.sigRender_('devi.osei@example.org', true).includedPhone, false);

  function refusal(api_) {
    try {
      api_.apiApply(false);
      return '(no error)';
    } catch (err) {
      return err.message;
    }
  }

  // Two ways an apply can change nothing. Neither may come back as a success the
  // page would render as "applied to " with an empty list.
  check('applying with nothing to write to is an error, not a silent no-op',
    refusal(api).indexOf('nothing to update') !== -1, true);
  check('an apply that failed everywhere is an error too',
    refusal(loadApi({
      updated: [], skipped: [], failed: [{ address: 'rowan.ashford@example.org', code: 403 }]
    })).indexOf('could not be written to rowan.ashford@example.org') !== -1, true);
  check('but a partial success is reported, not thrown',
    loadApi({
      updated: ['rowan.ashford@example.org'], skipped: [],
      failed: [{ address: 'rowan.ashford@example.cap.gov', code: 403 }]
    }).apiApply(false).updated, ['rowan.ashford@example.org']);
}

// ---------------------------------------------------------------------------
section('3b. Member-chosen duties: the member picks the contents, never the order');
{
  const MULTI = {
    rank: 'Maj', firstName: 'Rowan', middleName: 'K', lastName: 'Ashford', suffix: '',
    phone: '', orgName: 'EXAMPLE SR SQDN 80', type: 'SENIOR',
    dutyPositions: [
      duty('Commander', 'UNIT', 'EXAMPLE SR SQDN 80', 'UNIT'),
      duty('Director of IT', 'WING', 'CALIFORNIA WING HQ', 'WING'),
      duty('Safety Officer', 'WING', 'CALIFORNIA WING HQ', 'WING', true),
      duty('Personnel Officer', 'UNIT', 'EXAMPLE SR SQDN 80', 'UNIT', true)
    ]
  };
  const WING_PRIMARY = 'Director of IT|WING|CALIFORNIA WING HQ|P';
  const WING_ASSISTANT = 'Safety Officer|WING|CALIFORNIA WING HQ|A';
  const UNIT_PRIMARY = 'Commander|UNIT|EXAMPLE SR SQDN 80|P';
  const UNIT_ASSISTANT = 'Personnel Officer|UNIT|EXAMPLE SR SQDN 80|A';

  const block = keys => tpl.sigDutyBlock_(Object.assign({}, MULTI, { selectedDutyKeys: keys }));

  // The default, untouched: assistants out, one per echelon, highest first.
  check('with no selection, nothing changes',
    tpl.sigDutyBlock_(MULTI),
    'California Wing Director of IT<br />Example Senior Squadron 80 Commander');
  check('and src/ agrees, as always', tpl.sigDutyBlock_(MULTI), src.getDutyBlock(MULTI));

  // An assistant can be chosen — that is the new capability. And it must SAY it is
  // an assistant: CAPWATCH keeps that flag in a separate column from the title, so
  // printing the title bare ("California Wing Supply Officer") claims the billet
  // rather than assisting it. Nobody could see this while assistants were filtered
  // out of every signature.
  check('a chosen assistant duty appears, and says it is one',
    block([WING_ASSISTANT]), 'California Wing Assistant Safety Officer');
  check('...and src/ renders the identical block',
    block([WING_ASSISTANT]),
    src.getDutyBlock(Object.assign({}, MULTI, { selectedDutyKeys: [WING_ASSISTANT] })));

  // ECHELON IS THE OUTER KEY. A wing assistant still precedes a squadron principal:
  // the style guide asks for highest organizational level first, and the assistant
  // rule orders duties WITHIN a level.
  check('a wing assistant still outranks a squadron principal',
    block([UNIT_PRIMARY, WING_ASSISTANT]),
    'California Wing Assistant Safety Officer<br />Example Senior Squadron 80 Commander');
  check('...whichever order they were ticked in',
    block([WING_ASSISTANT, UNIT_PRIMARY]), block([UNIT_PRIMARY, WING_ASSISTANT]));

  // ...and within one echelon, the principal comes first.
  check('at the same level, principal before assistant',
    block([WING_ASSISTANT, WING_PRIMARY]),
    'California Wing Director of IT<br />California Wing Assistant Safety Officer');
  check('...also when ticked the other way round',
    block([WING_PRIMARY, WING_ASSISTANT]), block([WING_ASSISTANT, WING_PRIMARY]));

  // Two duties at one echelon are allowed BY SELECTION, though the default pick
  // spreads across echelons. An explicit choice is already the answer.
  check('a selection overrides the one-per-echelon default',
    block([WING_PRIMARY, WING_ASSISTANT]).indexOf('<br />') !== -1, true);
  check('an empty selection shows no duty block at all', block([]), '');
  check('one duty gives one line', block([UNIT_ASSISTANT]),
    'Example Senior Squadron 80 Assistant Personnel Officer');

  // The cap is a guarantee of the generator, not only of the page.
  check('a third duty is never printed even if it reaches the generator',
    block([WING_PRIMARY, WING_ASSISTANT, UNIT_PRIMARY]).split('<br />').length, 2);

  // The word "Assistant" is the member's assignment, not decoration. Getting this
  // wrong published a claim to a billet somebody else holds.
  check('a principal never gains the word', block([WING_PRIMARY]), 'California Wing Director of IT');
  check('...and src/ agrees on the assistant form',
    block([UNIT_ASSISTANT]),
    src.getDutyBlock(Object.assign({}, MULTI, { selectedDutyKeys: [UNIT_ASSISTANT] })));

  // A CAPWATCH title that already carries the word must not be doubled.
  const already = {
    orgName: 'EXAMPLE SR SQDN 80',
    dutyPositions: [duty('Assistant Finance Officer', 'UNIT', 'EXAMPLE SR SQDN 80', 'UNIT', true)]
  };
  const alreadyKey = 'Assistant Finance Officer|UNIT|EXAMPLE SR SQDN 80|A';
  check('a title that already says Assistant is not said twice',
    tpl.sigDutyBlock_(Object.assign({}, already, { selectedDutyKeys: [alreadyKey] })),
    'Example Senior Squadron 80 Assistant Finance Officer');

  // And the default path — where assistants never appear — is untouched by all of
  // this. The retired-title override still applies underneath the prefix.
  check('the default block still shows no assistant at all',
    tpl.sigDutyBlock_(MULTI).indexOf('Assistant'), -1);
  const retiredAssistant = {
    orgName: 'EXAMPLE SR SQDN 80',
    dutyPositions: [duty('Recruiting & Retention Officer', 'UNIT', 'EXAMPLE SR SQDN 80', 'UNIT', true)]
  };
  check('a renamed title keeps its rename under the prefix',
    tpl.sigDutyBlock_(Object.assign({}, retiredAssistant, {
      selectedDutyKeys: ['Recruiting & Retention Officer|UNIT|EXAMPLE SR SQDN 80|A']
    })),
    'Example Senior Squadron 80 Assistant Recruiting Officer');
}

// ---------------------------------------------------------------------------
section('3c. A duty selection is checked against the member\'s own record');
{
  const member = {
    capsn: '600001', rank: 'Maj', firstName: 'Rowan', lastName: 'Ashford',
    orgName: 'EXAMPLE SR SQDN 80', type: 'SENIOR', phone: '+15550101234',
    dutyPositions: [
      duty('Director of IT', 'WING', 'CALIFORNIA WING HQ', 'WING'),
      duty('Commander', 'UNIT', 'EXAMPLE SR SQDN 80', 'UNIT')
    ]
  };
  const MINE = 'Director of IT|WING|CALIFORNIA WING HQ|P';

  const api = loadModule(path.join(APP, 'SignatureApi.gs'), {
    SIG_CONFIG: SIG_CONFIG,
    SIG_MAX_DUTIES: 2,
    Logger: makeLogger().logger,
    HtmlService: {},
    requireMember_: () => 'rowan.ashford@example.org',
    sigResolveMemberForActor_: () => ({ capid: '600001', member: member }),
    sigGenerateEmailSignature_: tpl.sigGenerateEmailSignature_,
    sigDutyBlock_: tpl.sigDutyBlock_,
    sigDutyKey_: tpl.sigDutyKey_,
    sigDutyLevelRank_: tpl.sigDutyLevelRank_,
    sigName_: tpl.sigName_,
    sigFormatPhone_: tpl.sigFormatPhone_,
    sigSupportSentence_: () => 'Contact it@example.org.'
  }, ['sigRender_', 'sigValidatedDutyKeys_', 'sigDutyOptions_']);

  function refusal(keys) {
    try {
      api.sigRender_('rowan.ashford@example.org', false, keys);
      return '(no error)';
    } catch (err) {
      return err.message;
    }
  }

  check('omitting the selection means "pick for me"',
    api.sigValidatedDutyKeys_(member, undefined), null);
  check('...and null means the same', api.sigValidatedDutyKeys_(member, null), null);
  check('an empty array is a real choice, not a missing one',
    api.sigValidatedDutyKeys_(member, []), []);
  check('a duty the member holds is accepted', api.sigValidatedDutyKeys_(member, [MINE]), [MINE]);
  check('a duplicate is collapsed, not counted twice',
    api.sigValidatedDutyKeys_(member, [MINE, MINE]), [MINE]);

  // A key naming a duty the member does not hold is the one client input that could
  // put WORDS in a signature. It is refused loudly rather than dropped quietly —
  // dropping it would publish something the member never saw.
  check('a duty from somebody else\'s record is refused',
    refusal(['Commander|WING|CALIFORNIA WING HQ|P']).indexOf('not on your CAPWATCH record') !== -1,
    true);
  check('...as is an invented one', refusal(['Wing King|WING|SOMEWHERE|P'])
    .indexOf('not on your CAPWATCH record') !== -1, true);
  check('...and an assistant key for a duty the member holds as principal',
    refusal(['Director of IT|WING|CALIFORNIA WING HQ|A'])
      .indexOf('not on your CAPWATCH record') !== -1, true);
  check('too many is refused rather than silently truncated',
    refusal([MINE, 'Commander|UNIT|EXAMPLE SR SQDN 80|P', MINE + 'x'])
      .indexOf('at most 2 duty assignments') !== -1, true);
  check('a non-array is refused', refusal('Director of IT|WING|CALIFORNIA WING HQ|P')
    .indexOf('could not be read') !== -1, true);

  // The options the page renders.
  const options = api.sigDutyOptions_(member);
  check('every duty is offered', options.map(o => o.line),
    ['California Wing Director of IT', 'Example Senior Squadron 80 Commander']);
  check('...with the default pick marked', options.map(o => o.selected), [true, true]);
  check('...and keys that round-trip', api.sigValidatedDutyKeys_(member, [options[1].key]),
    [options[1].key]);

  // And the render honors it end to end.
  check('rendering with one duty prints one line',
    api.sigRender_('rowan.ashford@example.org', false, [MINE]).html
      .indexOf('Example Senior Squadron 80 Commander'), -1);
  check('...and reports back what it used',
    api.sigRender_('rowan.ashford@example.org', false, [MINE]).dutyKeys, [MINE]);
}

// ---------------------------------------------------------------------------
section('4. The record is built from CAPWATCH the way the main project builds it');

/** CAPWATCH rows, one source of truth, handed to both implementations. */
const ORG_ROWS = [
  // orgid, region, wing, unit, nextLevel, name, _, _, _, scope
  ['2500', 'PCR', 'CA', '080', '900', 'EXAMPLE SR SQDN 80', '', '', '', 'UNIT'],
  ['900', 'PCR', 'CA', '010', '188', 'CENTRAL CALIF GROUP 6', '', '', '', 'GROUP'],
  ['188', 'PCR', 'CA', '001', '', 'CALIFORNIA WING HQ', '', '', '', 'WING'],
  ['2600', 'PCR', 'CA', '012', '900', 'EXAMPLE CDT SQDN 12', '', '', '', 'UNIT'],
  // Another wing's org. getSquadrons() drops it, so a duty held here has no org
  // name and the signature falls back to the member's home unit.
  ['7700', 'PCR', 'HI', '001', '', 'HAWAII WING HQ', '', '', '', 'WING']
];

const CONTACT_ROWS = [
  // capid, type, priority, contact, _, _, doNotContact
  ['600001', 'CELL PHONE', 'SECONDARY', '(555) 010-7777', '', '', 'FALSE'],
  ['600001', 'CELL PHONE', 'PRIMARY', '555-010-1234', '', '', 'FALSE'],
  ['600001', 'HOME PHONE', 'PRIMARY', '555-010-2222', '', '', 'FALSE'],
  ['600001', 'EMAIL', 'PRIMARY', 'rowan@example.com', '', '', 'FALSE'],
  ['600002', 'CELL PHONE', 'PRIMARY', '555-010-3333', '', '', 'TRUE'],
  ['600003', 'CELL PHONE', 'PRIMARY', '555-010-4444', '', '', 'FALSE'],
  ['600004', 'CELL PHONE', 'PRIMARY', '555-010-5555', '', '', 'FALSE'],
  ['600006', 'CELL PHONE', 'PRIMARY', '555-010-6666', '', '', 'FALSE']
];

const DUTY_ROWS = [
  // capid, duty, _, level, assistant, _, _, orgid
  ['600001', 'Director of IT', '', 'WING', '0', '', '', '188'],
  ['600001', 'Information Technologies Officer', '', 'UNIT', '0', '', '', '2500'],
  ['600001', 'Safety Officer', '', 'UNIT', '1', '', '', '2500'],
  ['600004', 'Director of Operations', '', 'WING', '0', '', '', '7700']
];

const CADET_DUTY_ROWS = [
  ['600003', 'Cadet Commander', '', 'UNIT', '0', '', '', '2600']
];

/** Member.txt rows, built off the real header so the column indices are real. */
function memberRow(spec) {
  const cells = new Array(MEMBER_HEADER.length).fill('');
  cells[COL.CAPID] = spec.capid;
  cells[COL.NameLast] = spec.last;
  cells[COL.NameFirst] = spec.first;
  cells[COL.NameMiddle] = spec.middle || '';
  cells[COL.NameSuffix] = spec.suffix || '';
  cells[COL.ORGID] = spec.orgid;
  cells[COL.Wing] = 'CA';
  cells[COL.Rank] = spec.rank;
  cells[COL.Type] = spec.type;
  cells[COL.MbrStatus] = spec.status || 'ACTIVE';
  return cells;
}

const MEMBER_ROWS = [
  memberRow({ capid: '600001', last: 'Ashford', first: 'Rowan', middle: 'K', orgid: '2500', rank: 'Maj', type: 'SENIOR' }),
  memberRow({ capid: '600002', last: 'Calloway', first: 'Wren', orgid: '2500', rank: 'Lt Col', type: 'SENIOR' }),
  memberRow({ capid: '600003', last: 'Brightwater', first: 'Imani', orgid: '2600', rank: 'C/SMSgt', type: 'CADET' }),
  memberRow({ capid: '600004', last: 'Delacroix', first: 'Kai', orgid: '2500', rank: 'Maj', type: 'SENIOR' }),
  memberRow({ capid: '600005', last: 'Pemberton', first: 'Nia', orgid: '2500', rank: 'Capt', type: 'SENIOR', status: 'EXPIRED' }),
  // A cadet-lite grade: below C/SSgt, so CONFIG.CADET_LITE_EXCLUDED_GRADES keeps
  // this member out of getMembers() on the cadets tenant. See section 7.
  memberRow({ capid: '600006', last: 'Ferreira', first: 'Nico', orgid: '2600', rank: 'C/Amn', type: 'CADET' })
];

/** Rows -> a CAPWATCH .txt file, header included (parseFile drops row 0). */
function csv(header, rows) {
  const quote = cells => cells.map(v => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"').join(',');
  return quote(header) + '\n' + rows.map(quote).join('\n') + '\n';
}

/**
 * The REGION-wide extract, in its own folder.
 *
 * CAPWATCH scopes an extract to the echelon downloaded, so these rows — a wing
 * member's region billet, and the region org itself — exist ONLY here. The
 * fixture deliberately also repeats a duty the wing extract already has (600001's
 * wing directorship, orgid 188) and carries a member the wing does not have
 * (700001), because those are the two ways this source could do damage.
 */
const REGION_ORG_ROWS = [
  ['434', 'PCR', '', '001', '', 'PACIFIC REGION CAP', '', '', '', 'REGION'],
  ['1', 'NAT', '', '000', '', 'NATIONAL HEADQUARTERS', '', '', '', 'NAT'],
  ['188', 'PCR', 'CA', '001', '434', 'CALIFORNIA WING HQ', '', '', '', 'WING'],
  ['2500', 'PCR', 'CA', '080', '900', 'EXAMPLE SR SQDN 80', '', '', '', 'UNIT']
];

const REGION_DUTY_ROWS = [
  // The billet the wing pull cannot see: assistant, held at region.
  ['600001', 'Director of Operations', '', 'REGION', '1', '', '', '434'],
  // A national one, principal.
  ['600002', 'Advisor', '', 'NAT', '0', '', '', '1'],
  // Already in the wing extract — must not be added twice.
  ['600001', 'Director of IT', '', 'WING', '0', '', '', '188'],
  // A member of another wing entirely. Must never reach this tenant's roster.
  ['700001', 'Commander', '', 'WING', '0', '', '', '434']
];

const REGION_FILES = {
  'Organization.txt': csv(['ORGID', 'Region', 'Wing', 'Unit', 'NextLevel', 'Name', 'a', 'b', 'c', 'Scope'], REGION_ORG_ROWS),
  'DutyPosition.txt': csv(['CAPID', 'Duty', 'a', 'Lvl', 'Asst', 'b', 'c', 'ORGID'], REGION_DUTY_ROWS),
  'CadetDutyPositions.txt': csv(['CAPID', 'Duty', 'a', 'Lvl', 'Asst', 'b', 'c', 'ORGID'], [])
};

const FILES = {
  'Member.txt': csv(MEMBER_HEADER, MEMBER_ROWS),
  'Organization.txt': csv(['ORGID', 'Region', 'Wing', 'Unit', 'NextLevel', 'Name', 'a', 'b', 'c', 'Scope'], ORG_ROWS),
  'MbrContact.txt': csv(['CAPID', 'Type', 'Priority', 'Contact', 'a', 'b', 'DoNotContact'], CONTACT_ROWS),
  'DutyPosition.txt': csv(['CAPID', 'Duty', 'a', 'Lvl', 'Asst', 'b', 'c', 'ORGID'], DUTY_ROWS),
  'CadetDutyPositions.txt': csv(['CAPID', 'Duty', 'a', 'Lvl', 'Asst', 'b', 'c', 'ORGID'], CADET_DUTY_ROWS)
};

// Both extracts, now that the fixtures exist. The src module loaded above reads
// these through the same closure.
DRIVE_FOLDERS[WING_FOLDER] = FILES;
DRIVE_FOLDERS[REGION_FOLDER] = REGION_FILES;

/** Loads MemberRecord.gs over a fresh copy of the fixture extract. */
function loadRecordModule(cacheStore, sigConfig) {
  const store = cacheStore || {};
  return loadModule(path.join(APP, 'MemberRecord.gs'), {
    SIG_CONFIG: sigConfig || SIG_CONFIG,
    SIG_CAPID_RE: /^\d{5,7}$/,
    SIG_CACHE_TTL_SECONDS: 600,
    Logger: makeLogger().logger,
    Utilities: Utilities,
    DriveApp: makeFolderedDrive(DRIVE_FOLDERS),
    CacheService: {
      getUserCache: () => ({
        get: key => (Object.prototype.hasOwnProperty.call(store, key) ? store[key] : null),
        put: (key, value) => { store[key] = value; }
      })
    },
    AdminDirectory: {
      Users: {
        get: email => {
          const accounts = {
            'rowan.ashford@example.org': { externalIds: [{ type: 'organization', value: '600001' }] },
            'twinned@example.org': {
              externalIds: [
                { type: 'organization', value: '600001' },
                { type: 'organization', value: '600002' }
              ]
            },
            'retired.twin@example.org': {
              externalIds: [
                { type: 'custom', customType: 'duplicate_retired_capid', value: '600002' },
                { type: 'organization', value: '600001' }
              ]
            },
            'nocapid@example.org': { externalIds: [] },
            'expired.member@example.org': { externalIds: [{ type: 'organization', value: '600005' }] }
          };
          if (!accounts[email]) throw new Error('Not found');
          return accounts[email];
        }
      }
    }
  }, ['sigBuildMemberRecord_', 'sigResolveMemberForActor_', 'sigCapidsFromUser_',
      'sigDirectoryPhone_', 'sigOrgMap_']);
}

/** The same member, assembled by the MAIN project's own field derivations. */
function srcRecord(capid) {
  const squadrons = {};
  ORG_ROWS.forEach(row => {
    if (row[2] !== CONFIG.WING) return;   // getSquadrons()'s wing filter
    squadrons[row[0]] = {
      orgid: row[0], name: row[5], charter: row[1] + '-' + row[2] + '-' + row[3],
      unit: row[3], nextLevel: row[4], scope: row[9], wing: row[2], orgPath: ''
    };
  });

  const row = MEMBER_ROWS.filter(r => r[COL.CAPID] === capid)[0];
  const members = {};
  members[capid] = src.createMemberObject(row, squadrons);
  src.addContactInfo(members, CONTACT_ROWS);
  src.addDutyPositions(members, DUTY_ROWS, squadrons);
  src.addCadetDutyPositions(members, CADET_DUTY_ROWS, squadrons);
  // The same order getMembers() calls them in: our own extract first, then the
  // region supplement for what it cannot see.
  src.addOutOfWingDutyPositions_(members, squadrons);
  return members[capid];
}

/** The signature-relevant shape of a member, however it was built. */
function shape(member) {
  return {
    rank: member.rank || '',
    firstName: member.firstName || '',
    middleName: member.middleName || '',
    lastName: member.lastName || '',
    suffix: member.suffix || '',
    orgName: member.orgName || '',
    phone: member.phone || '',
    duties: (member.dutyPositions || []).map(d => [d.id, d.level, !!d.assistant, d.orgName || '', d.orgScope || ''])
  };
}

{
  const app = loadRecordModule();

  ['600001', '600003', '600004'].forEach(capid => {
    const mine = app.sigBuildMemberRecord_(capid);
    check('CAPID ' + capid + ': the same record as src/ builds', shape(mine), shape(srcRecord(capid)));
    check('CAPID ' + capid + ': and therefore the same signature',
      tpl.sigGenerateEmailSignature_(mine), src.generateEmailSignature(srcRecord(capid)));
  });

  check('a PRIMARY cell displaces a SECONDARY one',
    app.sigBuildMemberRecord_('600001').phone, '+15550101234');
  check('a DoNotContact number is never published',
    app.sigBuildMemberRecord_('600002').phone, '');
  check('a cadet has no phone at all, DoNotContact or not',
    app.sigBuildMemberRecord_('600003').phone, '');
  check('an assistant duty is carried but flagged',
    app.sigBuildMemberRecord_('600001').dutyPositions
      .filter(d => d.assistant && !d.outOfWing).length, 1);
  check('a duty at another wing\'s org carries no org name',
    app.sigBuildMemberRecord_('600004').dutyPositions[0].orgName, '');
  check('...so its line names the member\'s own unit',
    tpl.sigDutyBlock_(app.sigBuildMemberRecord_('600004')),
    'Example Senior Squadron 80 Director of Operations');
  check('an org outside this wing is not in the map at all',
    app.sigOrgMap_()['7700'], undefined);
  check('a member who is not ACTIVE has no record', app.sigBuildMemberRecord_('600005'), null);
  check('nor does an unknown CAPID', app.sigBuildMemberRecord_('999999'), null);
}

// ---------------------------------------------------------------------------
section('4b. Duties our own CAPWATCH pull cannot see');
{
  // CAPWATCH scopes an extract to the echelon downloaded, so a wing pull has NO
  // ROW for a member's region or national billet. Reading the region tenant's
  // region-wide extract is the only way it can appear — and that source is also
  // the one that could do the most damage, so both guards are pinned here.
  const app = loadRecordModule();

  const region = app.sigBuildMemberRecord_('600001').dutyPositions
    .filter(d => d.outOfWing);
  check('the region billet appears at all', region.length, 1);
  check('...named by the REGION extract\'s own org list, not ours',
    [region[0].orgName, region[0].orgScope], ['PACIFIC REGION CAP', 'REGION']);
  check('...at its real echelon', region[0].level, 'REGION');
  check('...and still marked assistant', region[0].assistant, true);
  check('...so it renders as the region, not the member\'s squadron',
    tpl.sigDutyBlock_(Object.assign({}, app.sigBuildMemberRecord_('600001'), {
      selectedDutyKeys: ['Director of Operations|REGION|PACIFIC REGION CAP|A']
    })),
    'Pacific Region Assistant Director of Operations');

  // A national assignment, which is the same problem one echelon further up.
  const national = app.sigBuildMemberRecord_('600002').dutyPositions.filter(d => d.outOfWing);
  check('a national billet appears too', national.map(d => [d.level, d.orgName]),
    [['NAT', 'NATIONAL HEADQUARTERS']]);
  check('...and outranks everything, per DUTY_LEVEL_ORDER',
    tpl.sigDutyBlock_(app.sigBuildMemberRecord_('600002')),
    'National Headquarters Advisor');

  // INVARIANT 1: the region extract must never introduce a member. It carries
  // every wing in the region; treating it as a roster would provision Nevada.
  check('a member who is only in the region extract stays unknown',
    app.sigBuildMemberRecord_('700001'), null);

  // The region extract repeats duties our own pull already has. Taking those too
  // would print a member's wing directorship twice.
  const wingDuties = app.sigBuildMemberRecord_('600001').dutyPositions
    .filter(d => d.id === 'Director of IT');
  check('a duty our own extract already has is not added twice', wingDuties.length, 1);
  check('...and the copy kept is ours, with our org name',
    [wingDuties[0].orgName, !!wingDuties[0].outOfWing], ['CALIFORNIA WING HQ', false]);

  // Self-disabling. A tenant that never sets the property, or sets it to a folder
  // it cannot read, must behave exactly as it did before this existed.
  const unset = loadRecordModule({}, Object.assign({}, SIG_CONFIG, {
    REGION_CAPWATCH_DATA_FOLDER_ID: ''
  }));
  check('with no region folder configured, nothing changes',
    unset.sigBuildMemberRecord_('600001').dutyPositions.filter(d => d.outOfWing).length, 0);

  const unreadable = loadRecordModule({}, Object.assign({}, SIG_CONFIG, {
    REGION_CAPWATCH_DATA_FOLDER_ID: 'not-shared-with-us'
  }));
  check('an unreadable folder degrades to no extra duties, not an error',
    unreadable.sigBuildMemberRecord_('600001').dutyPositions.filter(d => d.outOfWing).length, 0);
  check('...and the member still gets their signature',
    unreadable.sigBuildMemberRecord_('600001').dutyPositions.length > 0, true);
}

// ---------------------------------------------------------------------------
section('4c. INVARIANT 2: a region billet never changes group membership');
{
  // dutyPositionIds / dutyPositionIdsAndLevel are what UpdateGroups.gs matches
  // duty-based groups on. A wing member's region billet silently adding them to a
  // wing duty group is not something anyone asked for, so the supplement feeds
  // dutyPositions — which is what signatures read — and nothing else.
  const squadrons = {};
  ORG_ROWS.forEach(row => {
    if (row[2] !== CONFIG.WING) return;
    squadrons[row[0]] = {
      orgid: row[0], name: row[5], charter: row[1] + '-' + row[2] + '-' + row[3],
      unit: row[3], nextLevel: row[4], scope: row[9], wing: row[2], orgPath: ''
    };
  });

  const members = {};
  members['600001'] = src.createMemberObject(
    MEMBER_ROWS.filter(r => r[COL.CAPID] === '600001')[0], squadrons);
  src.addDutyPositions(members, DUTY_ROWS, squadrons);

  const idsBefore = members['600001'].dutyPositionIds.slice();
  const levelsBefore = members['600001'].dutyPositionIdsAndLevel.slice();
  const dutiesBefore = members['600001'].dutyPositions.length;

  src.addOutOfWingDutyPositions_(members, squadrons);

  check('the signature-facing list grows',
    members['600001'].dutyPositions.length > dutiesBefore, true);
  check('dutyPositionIds does NOT', members['600001'].dutyPositionIds, idsBefore);
  check('nor does dutyPositionIdsAndLevel',
    members['600001'].dutyPositionIdsAndLevel, levelsBefore);
  check('and the added duty says where it came from',
    members['600001'].dutyPositions.filter(d => d.outOfWing).map(d => d.id),
    ['Director of Operations']);

  // The out-of-wing member is still absent afterwards.
  check('the region-only member was not created', members['700001'], undefined);
}

// ---------------------------------------------------------------------------
section('5. Identity: the account acted on is the signed-in one, and only it');
{
  const app = loadRecordModule();

  check('the CAPID comes off the caller\'s own account',
    app.sigResolveMemberForActor_('rowan.ashford@example.org').capid, '600001');

  // The retired-twin marker parks a dead CAPID on an account. Resolving it would
  // put someone else's record on this member's signature.
  check('a retired-twin CAPID is ignored',
    app.sigCapidsFromUser_({
      externalIds: [
        { type: 'custom', customType: 'duplicate_retired_capid', value: '600002' },
        { type: 'organization', value: '600001' }
      ]
    }), ['600001']);
  check('...and the account still resolves to its live CAPID',
    app.sigResolveMemberForActor_('retired.twin@example.org').capid, '600001');

  function refusal(email) {
    try {
      app.sigResolveMemberForActor_(email);
      return '(no error)';
    } catch (err) {
      return err.message;
    }
  }

  check('two live CAPIDs is refused rather than guessed',
    refusal('twinned@example.org').indexOf('more than one CAPID') !== -1, true);
  check('an account with no CAPID gets a sentence it can act on',
    refusal('nocapid@example.org').indexOf('not linked to a CAPID') !== -1, true);
  check('an expired member is told why there is nothing to build',
    refusal('expired.member@example.org').indexOf('not an active member') !== -1, true);
  check('and the refusal names who to ask',
    refusal('nocapid@example.org').indexOf('it@example.org') !== -1, true);

  // The cache must be a speed-up, never a channel: a second call for one member
  // returns that member, and never anybody else's record.
  const store = {};
  const cached = loadRecordModule(store);
  const first = cached.sigResolveMemberForActor_('rowan.ashford@example.org');
  const second = cached.sigResolveMemberForActor_('rowan.ashford@example.org');
  check('a cache hit returns the same record', second, first);
  check('and it is keyed by the authenticated address',
    Object.keys(store), ['sigmember:rowan.ashford@example.org']);
}

// ---------------------------------------------------------------------------
section('6. A signature is only ever written to an address the organization owns');
{
  const auth = loadModule(path.join(APP, 'Auth.gs'), {
    SIG_CONFIG: SIG_CONFIG,
    Logger: makeLogger().logger,
    Session: { getActiveUser: () => ({ getEmail: () => 'rowan.ashford@example.org' }) },
    AdminDirectory: { Members: { hasMember: (group, email) => ({ isMember: email === 'allowed@example.org' }) } }
  }, ['isOnATenantDomain_', 'mayUseSignatureApp_', 'resolveActor_', 'requireMember_']);

  check('the primary domain is ours', auth.isOnATenantDomain_('rowan.ashford@example.org'), true);
  check('so is the secondary domain', auth.isOnATenantDomain_('rowan.ashford@example.cap.gov'), true);
  check('a personal address is not', auth.isOnATenantDomain_('rowan@gmail.com'), false);
  check('and neither is a lookalike', auth.isOnATenantDomain_('rowan@example.org.evil.com'), false);
  check('nor a subdomain of ours', auth.isOnATenantDomain_('rowan@mail.example.org'), false);

  check('with no group configured, any member of the tenant may use the app',
    auth.mayUseSignatureApp_('anyone@example.org'), true);
  check('but never someone outside it', auth.mayUseSignatureApp_('outsider@gmail.com'), false);
  check('and never a blank identity', auth.mayUseSignatureApp_(''), false);

  const gated = loadModule(path.join(APP, 'Auth.gs'), {
    SIG_CONFIG: Object.assign({}, SIG_CONFIG, { ALLOWED_GROUP: 'pilot@example.org' }),
    Logger: makeLogger().logger,
    Session: { getActiveUser: () => ({ getEmail: () => 'allowed@example.org' }) },
    AdminDirectory: { Members: { hasMember: (group, email) => ({ isMember: email === 'allowed@example.org' }) } }
  }, ['mayUseSignatureApp_']);
  check('with a group configured, a member of it may', gated.mayUseSignatureApp_('allowed@example.org'), true);
  check('and a non-member may not', gated.mayUseSignatureApp_('other@example.org'), false);

  const broken = loadModule(path.join(APP, 'Auth.gs'), {
    SIG_CONFIG: Object.assign({}, SIG_CONFIG, { ALLOWED_GROUP: 'typo@example.org' }),
    Logger: makeLogger().logger,
    Session: { getActiveUser: () => ({ getEmail: () => 'allowed@example.org' }) },
    AdminDirectory: { Members: { hasMember: () => { throw new Error('Resource Not Found: groupKey'); } } }
  }, ['mayUseSignatureApp_']);
  check('a group check that cannot be completed denies', broken.mayUseSignatureApp_('allowed@example.org'), false);

  // Now the write itself, over a mailbox carrying a personal Send-As identity —
  // the case the guard exists for.
  //
  // The two endpoints GmailSignature.gs is expected to reach, built the way it
  // builds them so the stub matches the real thing rather than a guess at it.
  const MAILBOX = 'rowan.ashford@example.org';
  const TOKEN_URL = 'https://oauth2.googleapis.com/token';
  const SENDAS_URL = 'https://gmail.googleapis.com/gmail/v1/users/' +
    encodeURIComponent(MAILBOX) + '/settings/sendAs';

  const patched = [];
  const gmail = loadModule(path.join(APP, 'GmailSignature.gs'), {
    Logger: makeLogger().logger,
    isOnATenantDomain_: auth.isOnATenantDomain_,
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: key => ({
          SA_IMPERSONATION_EMAIL: 'sa@example.iam.gserviceaccount.com',
          SA_PRIVATE_KEY: '-----BEGIN PRIVATE KEY-----\\nx\\n-----END PRIVATE KEY-----'
        })[key] || null
      })
    },
    Utilities: {
      base64EncodeWebSafe: v => 'b64',
      computeRsaSha256Signature: () => 'sig'
    },
    // Routed on the WHOLE url, not on a host appearing somewhere in it: the stub
    // then also asserts which endpoints the module talks to, and an unexpected one
    // fails loudly instead of being served whatever branch it happened to reach.
    UrlFetchApp: {
      fetch: (url, options) => {
        if (url === TOKEN_URL) {
          return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ access_token: 'tok' }) };
        }
        if (url === SENDAS_URL && options.method === 'get') {
          return {
            getResponseCode: () => 200,
            getContentText: () => JSON.stringify({
              sendAs: [
                { sendAsEmail: 'rowan.ashford@example.org' },
                { sendAsEmail: 'rowan.ashford@example.cap.gov' },
                { sendAsEmail: 'rowan@gmail.com' },
                { sendAsEmail: 'rowan@example.org.evil.com' }
              ]
            })
          };
        }
        if (url.startsWith(SENDAS_URL + '/') && options.method === 'patch') {
          patched.push(decodeURIComponent(url.slice((SENDAS_URL + '/').length)));
          return { getResponseCode: () => 200, getContentText: () => '{}' };
        }
        throw new Error('the module called an endpoint this test did not expect: ' +
          options.method + ' ' + url);
      }
    }
  }, ['sigApplyToSendAsIdentities_', 'sigSendAsSnapshot_']);

  const result = gmail.sigApplyToSendAsIdentities_(MAILBOX, '<html></html>');
  check('only the org-owned identities are written',
    patched.sort(), ['rowan.ashford@example.cap.gov', 'rowan.ashford@example.org']);
  check('the personal ones are reported as skipped, not touched',
    result.skipped.sort(), ['rowan@example.org.evil.com', 'rowan@gmail.com']);
  check('and the member is told what was updated',
    result.updated.sort(), ['rowan.ashford@example.cap.gov', 'rowan.ashford@example.org']);
}

// ---------------------------------------------------------------------------
section('7. On the CADETS tenant');
{
  // Same code, different Script Properties: cadet domain, and NO secondary domain
  // (config-tenants/cadets.json does not set one).
  const CADET_CONFIG = Object.assign({}, SIG_CONFIG, {
    EMAIL_DOMAIN: '@example-cadets.org',
    SECONDARY_EMAIL_DOMAIN: ''
  });

  const auth = loadModule(path.join(APP, 'Auth.gs'), {
    SIG_CONFIG: CADET_CONFIG,
    Logger: makeLogger().logger,
    Session: { getActiveUser: () => ({ getEmail: () => 'imani.brightwater@example-cadets.org' }) },
    AdminDirectory: { Members: { hasMember: () => ({ isMember: true }) } }
  }, ['isOnATenantDomain_', 'mayUseSignatureApp_']);

  check('a cadet-domain address is writable', auth.isOnATenantDomain_('x@example-cadets.org'), true);
  // A cadet may hold a senior-domain Send-As (a cadet who is also staff, or one
  // mid-transition). The CADETS deployment must not write it: that address belongs
  // to the seniors tenant's own deployment, which has its own record for it.
  check('but the senior tenant\'s domain is NOT this deployment\'s to write',
    auth.isOnATenantDomain_('x@example.org'), false);
  check('and with no secondary domain configured nothing else slips through',
    auth.isOnATenantDomain_('x@example.cap.gov'), false);
  check('a cadet may use the app', auth.mayUseSignatureApp_('imani.brightwater@example-cadets.org'), true);

  const app = loadRecordModule();

  // A cadet's number is never published, and this is the case that proves the rule
  // rather than the absence of data: 600003 and 600006 both have a clean, PRIMARY,
  // non-DoNotContact cell row in the fixture extract.
  check('a cadet with a perfectly good cell still gets no phone row',
    app.sigBuildMemberRecord_('600003').phone, '');
  check('...and neither does a cadet-lite one', app.sigBuildMemberRecord_('600006').phone, '');

  // Deliberate divergence, documented in docs/SIGNATURE_WEB_APP.md: getMembers()
  // filters cadet-lite grades out because they get no ACCOUNT, so pushAllSignatures()
  // never covers them. This app is reached BY a signed-in account, so eligibility to
  // be provisioned is not the question — the member is already here, and refusing to
  // format a signature for an account that exists would help nobody.
  const cadetLite = app.sigBuildMemberRecord_('600006');
  check('a cadet-lite member who does have an account still gets a signature',
    cadetLite === null, false);
  check('...rendered with their cadet grade',
    tpl.sigName_(cadetLite), 'Cadet Airman Nico Ferreira');
  check('...and identical to what src/ would render for the same record',
    tpl.sigGenerateEmailSignature_(cadetLite), src.generateEmailSignature(cadetLite));

  // The page: no toggle, and a reason instead.
  const api = loadModule(path.join(APP, 'SignatureApi.gs'), {
    SIG_CONFIG: CADET_CONFIG,
    Logger: makeLogger().logger,
    HtmlService: {},
    requireMember_: () => 'imani.brightwater@example-cadets.org',
    sigResolveMemberForActor_: () => ({ capid: '600003', member: app.sigBuildMemberRecord_('600003') }),
    SIG_MAX_DUTIES: 2,
    sigGenerateEmailSignature_: tpl.sigGenerateEmailSignature_,
    sigName_: tpl.sigName_,
    sigDutyBlock_: tpl.sigDutyBlock_,
    sigDutyKey_: tpl.sigDutyKey_,
    sigDutyLevelRank_: tpl.sigDutyLevelRank_,
    sigFormatPhone_: tpl.sigFormatPhone_,
    sigSendAsSnapshot_: () => ({ orgOwned: ['imani.brightwater@example-cadets.org'] }),
    sigApplyToSendAsIdentities_: () => ({ updated: [], skipped: [], failed: [] }),
    sigSupportSentence_: () => 'Contact it@example.org.'
  }, ['apiGetState', 'sigRender_']);

  const state = api.apiGetState();
  check('the page is told this is a cadet', state.isCadet, true);
  check('...with no phone to offer', state.phone, '');
  check('...so ticking the box anyway changes nothing',
    state.html, api.sigRender_('imani.brightwater@example-cadets.org', true).html);
  check('and the cadet duty is still shown', state.dutyLines, ['Example Cadet Squadron 12 Cadet Commander']);
}

// ---------------------------------------------------------------------------
section('8. An unset Script Property is absent, not blank');
{
  // The Apps Script UI will not store an empty value, so an "optional, leave blank"
  // property is simply never created and getProperty() returns null. Everything
  // below loads the REAL Config.gs over a properties store that behaves that way.
  function loadConfig(properties) {
    return loadModule(path.join(APP, 'Config.gs'), {
      PropertiesService: {
        getScriptProperties: () => ({
          getProperty: key => (Object.prototype.hasOwnProperty.call(properties, key)
            ? properties[key]
            : null)
        })
      }
    }, ['SIG_CONFIG', 'sigMissingConfig_']);
  }

  const SET = {
    TENANT_EMAIL_DOMAIN: '@example.org',
    TENANT_WING: 'CA',
    TENANT_CAPWATCH_DATA_FOLDER_ID: 'folder-id',
    SA_IMPERSONATION_EMAIL: 'sa@example.iam.gserviceaccount.com',
    SA_PRIVATE_KEY: '-----BEGIN PRIVATE KEY-----'
  };

  // The four properties an operator would have been asked to "leave blank", now
  // simply not created. None of them may break the app.
  const lean = loadConfig(SET);
  check('an omitted allowed-group means every member may use it', lean.SIG_CONFIG.ALLOWED_GROUP, '');
  check('an omitted wing abbreviation still derives one', lean.SIG_CONFIG.ORG_LABEL, 'CAWG');
  check('an omitted secondary domain is just absent', lean.SIG_CONFIG.SECONDARY_EMAIL_DOMAIN, '');
  check('an omitted support address is absent, not "null"', lean.SIG_CONFIG.SUPPORT_EMAIL, '');
  check('and with the required ones set, the app is usable', lean.sigMissingConfig_(), []);

  // The reader must not turn a missing key into the STRING "null" — it is
  // concatenated into member-facing sentences and compared against domains.
  check('a missing key never becomes the string "null"',
    [lean.SIG_CONFIG.SECONDARY_EMAIL_DOMAIN, lean.SIG_CONFIG.ALLOWED_GROUP,
     lean.SIG_CONFIG.SUPPORT_EMAIL].join('|'), '||');

  // An explicitly blank value (a property created and then emptied) reads the same
  // as an absent one, so an operator cannot get a third behavior by choosing.
  const blanked = loadConfig(Object.assign({}, SET, {
    SIGNATURE_WEBAPP_ALLOWED_GROUP: '', TENANT_WING_ABBREVIATION: '   '
  }));
  check('blank and absent behave identically', blanked.SIG_CONFIG.ALLOWED_GROUP, '');
  check('...including whitespace-only', blanked.SIG_CONFIG.ORG_LABEL, 'CAWG');

  // A REQUIRED property missing is the case that used to masquerade as an
  // authorization failure: isOnATenantDomain_() would match nothing and every
  // member would be told their own account was the problem.
  const noDomain = loadConfig({ TENANT_WING: 'CA', TENANT_CAPWATCH_DATA_FOLDER_ID: 'f',
    SA_IMPERSONATION_EMAIL: 'x', SA_PRIVATE_KEY: 'y' });
  check('a missing mail domain is reported as missing config',
    noDomain.sigMissingConfig_(), ['TENANT_EMAIL_DOMAIN']);
  check('missing service-account credentials count too',
    loadConfig({ TENANT_EMAIL_DOMAIN: '@example.org', TENANT_WING: 'CA',
      TENANT_CAPWATCH_DATA_FOLDER_ID: 'f' }).sigMissingConfig_(),
    ['SA_IMPERSONATION_EMAIL / SA_PRIVATE_KEY']);
  check('half a credential is not a credential',
    loadConfig(Object.assign({}, SET, { SA_PRIVATE_KEY: null })).sigMissingConfig_(),
    ['SA_IMPERSONATION_EMAIL / SA_PRIVATE_KEY']);
  check('the private key itself is never carried in the config object',
    Object.keys(lean.SIG_CONFIG).filter(k => /KEY|PRIVATE/.test(k)), []);

  // ...and the gate says so, instead of blaming the caller.
  const logged = makeLogger();
  const auth = loadModule(path.join(APP, 'Auth.gs'), {
    SIG_CONFIG: noDomain.SIG_CONFIG,
    sigMissingConfig_: noDomain.sigMissingConfig_,
    sigSupportSentence_: () => 'Contact it@example.org if you think this is wrong.',
    Logger: logged.logger,
    Session: { getActiveUser: () => ({ getEmail: () => 'rowan.ashford@example.org' }) },
    AdminDirectory: { Members: { hasMember: () => ({ isMember: true }) } }
  }, ['requireMember_']);

  let threw = '';
  try {
    auth.requireMember_();
  } catch (err) {
    threw = err.message;
  }
  check('an unconfigured tenant does NOT read as "you are not authorized"',
    threw.indexOf('not authorized') === -1 && threw.indexOf('not signed in') === -1, true);
  check('...it says the page is not set up', threw.indexOf('not set up') !== -1, true);
  check('...and names the property in the log, not to the member',
    logged.calls.error[0].ctx.missing, ['TENANT_EMAIL_DOMAIN']);
  check('...which the member never sees', threw.indexOf('TENANT_'), -1);

  // The two pages that refuse are where a member is most stuck, so both hand them
  // CAP's own generator rather than only telling them what did not work.
  function servedPage(globals) {
    let served = '';
    const api = loadModule(path.join(APP, 'SignatureApi.gs'), Object.assign({
      SIG_CONFIG: SIG_CONFIG,
      SIG_GENERATOR_URL: 'https://cap-brand-tools.netlify.app/signature-generator/index.html',
      Logger: makeLogger().logger,
      HtmlService: {
        createHtmlOutput: html => { served = html; return { setTitle: () => served }; }
      },
      sigSupportSentence_: () => 'Contact it@example.org if you think this is wrong.',
      sigMissingConfig_: () => [],
      resolveActor_: () => '',
      mayUseSignatureApp_: () => false
    }, globals || {}), ['doGet']);
    api.doGet();
    return served;
  }

  const notConfigured = servedPage({ sigMissingConfig_: () => ['TENANT_EMAIL_DOMAIN'] });
  check('the not-set-up page offers the generator',
    notConfigured.indexOf('signature-generator/index.html') !== -1, true);
  check('...and still does not name the property to the member',
    notConfigured.indexOf('TENANT_'), -1);

  const notAuthorized = servedPage({});
  check('the not-signed-in page offers it too',
    notAuthorized.indexOf('signature-generator/index.html') !== -1, true);
  check('...as a link that opens in a new tab',
    notAuthorized.indexOf('target="_blank" rel="noopener"') !== -1, true);
}

done();
