/**
 * RecoveryEmailNotify.gs — end-to-end.
 *
 * Drives the real notifyRecoveryEmailCompliance() against a fake Drive, a Gmail
 * stub that records every send, and a clock a test can move.
 *
 * The question this exists to answer is the inverse of the LSCode module's. That
 * one must be SILENT on its first run; this one must be LOUD — it reports a
 * standing condition, so the first run is meant to surface the whole backlog.
 * What must not happen is the second run doing it again: a monthly job that
 * re-mails the same commander about the same member every month gets filtered to
 * trash, and the three-month window is the only thing preventing that. Moving the
 * clock is the only way to prove the window really opens and closes.
 *
 * Run: npm test
 */
const path = require('path');
const {
  loadModule, makeLogger, makeDrive, makeGmail, makeChecker, makeClock,
  Session, Utilities
} = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'notifications', 'RecoveryEmailNotify.gs');
const { section, check, done } = makeChecker();

const STATE_FILE = 'RecoveryComplianceState.txt';

// Commanders.txt: ORGID=0, CAPID=4, NameLast=8, NameFirst=9, Rank=12
const COMMANDERS = [
  ['100', 'PCR', 'CA', '070', '900001', '', '', '', 'Alpha', 'Ada', '', '', 'Maj'],
  ['200', 'PCR', 'CA', '080', '900002', '', '', '', 'Bravo', 'Ben', '', '', 'Maj']
];
// DutyPosition.txt: CAPID=0, Duty=1, Level=3, Asst=4, ORGID=7
const DUTIES = [
  ['900003', 'Personnel Officer', '', 'UNIT', '0', '', '', '100'],
  ['900004', 'Deputy Commander', '', 'UNIT', '0', '', '', '100']
];
// Raw Member.txt rows, for recipient names: CAPID=0, NameLast=2, NameFirst=3, Rank=14
const MEMBER_ROWS = [
  ['900003', '', 'Papa', 'Pat', '', '', '', '', '', '', '', '100', '', '', 'Capt'],
  ['900004', '', 'Delta', 'Dana', '', '', '', '', '', '', '', '100', '', '', 'Maj']
];

/**
 * @param {Object} spec - { capid, orgid, primary, recovery, manual }
 * @returns {Object} A member as getMembers() returns it
 */
function member(spec) {
  const m = {
    capsn: spec.capid,
    firstName: 'First' + spec.capid,
    lastName: 'Last' + spec.capid,
    rank: 'Capt',
    type: 'SENIOR',
    orgid: spec.orgid || '100',
    orgName: spec.orgid === '200' ? 'Bravo Squadron' : 'Alpha Squadron',
    charter: spec.orgid === '200' ? 'PCR-CA-080' : 'PCR-CA-070',
    primaryEmailValue: spec.primary
  };
  if (!spec.manual) m.recoveryEmail = spec.recovery === undefined ? '' : spec.recovery;
  return m;
}

/** A compliant member: CAP address primary, personal address for recovery. */
function ok(capid, orgid) {
  return member({ capid, orgid, primary: 'm' + capid + '@cawgcap.org', recovery: 'm' + capid + '@gmail.com' });
}
/** A member with no personal address anywhere. */
function noRecovery(capid, orgid) {
  return member({ capid, orgid, primary: 'm' + capid + '@cawgcap.org', recovery: '' });
}

/**
 * Builds a module instance over the given Drive contents.
 *
 * @param {Object} files - filename -> content, mutated in place by the module
 * @param {Object} opts - { members, enabled, failFrom, today, commanders, duties }
 * @returns {Object} { m, sent, log, clock }
 */
function harness(files, opts) {
  const o = opts || {};
  const enabled = o.enabled !== false;
  const { logger, calls } = makeLogger();
  const { gmail, sent } = makeGmail({ failFrom: o.failFrom || [] });
  const clock = makeClock(o.today || '2026-01-01');
  const state = { members: o.members || {} };

  const m = loadModule(MODULE, {
    Logger: logger,
    Session: Session,
    Utilities: Utilities,
    Date: clock.Date,
    DriveApp: makeDrive(files),
    GmailApp: gmail,
    CONFIG: {
      CAPWATCH_DATA_FOLDER_ID: 'fake-folder-id',
      DOMAIN: 'cawgcap.org',
      SECONDARY_EMAIL_DOMAIN: '@cawg.cap.gov',
      EMAIL_DOMAIN: '@cawgcap.org',
      COMMAND_EMAIL_DOMAIN: '@cawgcap.org',
      ORG_LABEL: 'CAWG',
      MEMBER_TYPES: { ACTIVE: ['SENIOR'] }
    },
    PROFILE_: { RUN_RECOVERY_EMAIL_NOTIFICATIONS: enabled },
    TENANT_PROFILE: enabled ? 'seniors' : 'region',
    clearCache: () => {},
    getMembers: () => state.members,
    parseFile: name => {
      if (name === 'Commanders') return o.commanders || COMMANDERS;
      if (name === 'DutyPosition') return o.duties || DUTIES;
      if (name === 'Member') return MEMBER_ROWS;
      return [];
    },
    createEmailMap: () => ({}),
    getActiveUsers: () => [],
    executeWithRetry: fn => fn(),
    AUTOMATION_SENDER_EMAIL: 'automation@cawgcap.org',
    SENDER_NAME: 'CAWG Information Technology',
    ITSUPPORT_EMAIL: 'it.support@cawgcap.org',
    TEST_EMAIL: 'test@cawgcap.org'
  }, [
    'notifyRecoveryEmailCompliance', 'previewRecoveryEmailCompliance',
    'installRecoveryComplianceMonthlyTrigger'
  ]);

  return { m, sent, log: calls, clock, state };
}

// ---------------------------------------------------------------------------
section('1. The first run reports the standing backlog — it is NOT silent');
{
  const files = {};
  const { m, sent } = harness(files, {
    members: { '1': noRecovery('1'), '2': ok('2'), '3': member({ capid: '3', primary: 'p@gmail.com', recovery: 'p@gmail.com' }) }
  });
  const summary = m.notifyRecoveryEmailCompliance();

  check('mails the unit once', sent.length, 1);
  check('flagging the two non-compliant members', summary.flagged, 2);
  check('leaving the compliant one alone', summary.compliant, 1);
  check('names both flagged members',
    /<td>1<\/td>/.test(sent[0].html) && /<td>3<\/td>/.test(sent[0].html), true);
  check('and not the compliant one', /<td>2<\/td>/.test(sent[0].html), false);

  const state = JSON.parse(files[STATE_FILE]);
  check('records only the members it reported',
    Object.keys(state.members).sort(), ['1', '3']);
  check('dated today', state.members['1'].lastNotified, '2026-01-01');
  check('with the reason', [state.members['1'].reasons, state.members['3'].reasons],
    ['RECOVERY', 'PRIMARY']);
}

// ---------------------------------------------------------------------------
section('2. Running again the next month says nothing — the whole point');
{
  const files = {};
  const members = { '1': noRecovery('1') };
  const { m, sent, clock } = harness(files, { members: members });

  m.notifyRecoveryEmailCompliance();
  check('reported once', sent.length, 1);

  clock.set('2026-02-01');
  const feb = m.notifyRecoveryEmailCompliance();
  check('February: still nothing sent', sent.length, 1);
  check('the member is counted as suppressed', feb.suppressed, 1);
  check('and nobody was reported', feb.reported, 0);

  clock.set('2026-03-01');
  m.notifyRecoveryEmailCompliance();
  check('March: still quiet', sent.length, 1);

  // The window runs from when they were told, so it stays 1 January.
  check('the recorded date has not drifted',
    JSON.parse(files[STATE_FILE]).members['1'].lastNotified, '2026-01-01');
}

// ---------------------------------------------------------------------------
section('3. Three months on, an uncorrected member is reported again');
{
  const files = {};
  const { m, sent, clock } = harness(files, { members: { '1': noRecovery('1') } });

  m.notifyRecoveryEmailCompliance();
  clock.set('2026-04-01');
  const apr = m.notifyRecoveryEmailCompliance();

  check('a second digest goes out', sent.length, 2);
  check('reporting the member again', apr.reported, 1);
  check('and the window restarts from today',
    JSON.parse(files[STATE_FILE]).members['1'].lastNotified, '2026-04-01');
}

// ---------------------------------------------------------------------------
section('4. A member who fixes it drops out — and a relapse is reported at once');
{
  const files = {};
  const members = { '1': noRecovery('1') };
  const { m, sent, clock, state } = harness(files, { members: members });

  m.notifyRecoveryEmailCompliance();
  check('reported', sent.length, 1);

  // They add a personal address in February.
  clock.set('2026-02-01');
  state.members = { '1': ok('1') };
  m.notifyRecoveryEmailCompliance();

  check('nothing further is sent', sent.length, 1);
  check('and they are cleared from state',
    Object.keys(JSON.parse(files[STATE_FILE]).members).length, 0);

  // In March it lapses again. Had they been left sitting in state, they would
  // have stayed silently suppressed until April.
  clock.set('2026-03-01');
  state.members = { '1': noRecovery('1') };
  const mar = m.notifyRecoveryEmailCompliance();

  check('the relapse is reported immediately, not in April', sent.length, 2);
  check('counted as reported', mar.reported, 1);
}

// ---------------------------------------------------------------------------
section('5. The commander is the addressee; the staff are copied');
{
  const files = {};
  const { m, sent } = harness(files, { members: { '1': noRecovery('1') } });
  m.notifyRecoveryEmailCompliance();

  check('addressed to the commander', sent[0].to, 'ada.alpha@cawgcap.org');
  check('copying the personnel officer and deputy commander',
    sent[0].cc.split(',').sort(), ['dana.delta@cawgcap.org', 'pat.papa@cawgcap.org']);
  check('subject names the charter', sent[0].subject.indexOf('PCR-CA-070') > -1, true);
  check('sent as the automation identity', sent[0].from, 'automation@cawgcap.org');
  check('replies go to IT', sent[0].replyTo, 'it.support@cawgcap.org');
}

// ---------------------------------------------------------------------------
section('6. Each unit hears only about its own members');
{
  const files = {};
  const { m, sent } = harness(files, {
    members: { '1': noRecovery('1', '100'), '9': noRecovery('9', '200') }
  });
  m.notifyRecoveryEmailCompliance();

  check('two digests', sent.length, 2);
  const alpha = sent.filter(s => s.to === 'ada.alpha@cawgcap.org')[0];
  const bravo = sent.filter(s => s.to === 'ben.bravo@cawgcap.org')[0];

  check('Alpha hears about its own member', /<td>1<\/td>/.test(alpha.html), true);
  check('and not about Bravo\'s', /<td>9<\/td>/.test(alpha.html), false);
  check('Bravo hears about its own member', /<td>9<\/td>/.test(bravo.html), true);
  // Bravo has no personnel officer or deputy in the fixtures.
  check('a unit with no other staff copies nobody', bravo.cc, undefined);
}

// ---------------------------------------------------------------------------
section('7. A unit with no reachable staff stays pending, and IT is told');
{
  const files = {};
  const { m, sent } = harness(files, {
    members: { '9': noRecovery('9', '200') },
    commanders: [],   // no commander on record anywhere
    duties: []
  });
  const summary = m.notifyRecoveryEmailCompliance();

  check('no digest could be sent', summary.sent, 0);
  check('the unit is reported as unreachable', summary.noRecipientOrgs.length, 1);
  check('the member is NOT recorded as notified',
    Object.keys(JSON.parse(files[STATE_FILE]).members).length, 0);

  const alarm = sent.filter(s => s.to === 'it.support@cawgcap.org');
  check('IT is told once', alarm.length, 1);

  // Pending rather than lost: it must fire once a commander exists.
  const { m: m2, sent: sent2 } = harness(files, { members: { '9': noRecovery('9', '200') } });
  m2.notifyRecoveryEmailCompliance();
  check('once a commander is assigned, it goes out',
    sent2.filter(s => s.to === 'ben.bravo@cawgcap.org').length, 1);
}

// ---------------------------------------------------------------------------
section('8. Wrong sender identity: digests fail, the alarm still gets out');
{
  // The live 2026-07-16 LSCode failure, reproduced: the executing account cannot
  // send as automation@, so every digest bounces with "Invalid argument".
  const files = {};
  const { m, sent } = harness(files, {
    members: { '1': noRecovery('1') },
    failFrom: ['automation@cawgcap.org']
  });
  const summary = m.notifyRecoveryEmailCompliance();

  check('the digest failed', summary.failedOrgs.length, 1);
  check('nothing recorded as sent', summary.sent, 0);
  check('the member stays pending for a retry',
    Object.keys(JSON.parse(files[STATE_FILE]).members).length, 0);

  const alarm = sent.filter(s => s.to === 'it.support@cawgcap.org');
  check('the IT alarm was delivered anyway', alarm.length, 1);
  check('because it sets no from override', alarm[0].from, undefined);
  check('no digest slipped through', sent.filter(s => s.from === 'automation@cawgcap.org').length, 0);
}

// ---------------------------------------------------------------------------
section('9. preview sends nothing, writes nothing, consumes nothing');
{
  const files = {};
  const { m, sent } = harness(files, { members: { '1': noRecovery('1') } });

  const preview = m.previewRecoveryEmailCompliance();
  check('no mail', sent.length, 0);
  check('no state file', files[STATE_FILE], undefined);
  check('but it reports what would go', preview.sent, 1);
  check('and how many members', preview.reported, 1);

  const real = m.notifyRecoveryEmailCompliance();
  check('the real run afterwards still delivers it', sent.length, 1);
  check('having consumed nothing', real.reported, 1);
}

// ---------------------------------------------------------------------------
section('10. A disabled profile does nothing at all');
{
  const files = {};
  const { m, sent } = harness(files, { members: { '1': noRecovery('1') }, enabled: false });
  const summary = m.notifyRecoveryEmailCompliance();

  check('skips', summary.skipped, true);
  check('sends nothing', sent.length, 0);
  check('writes no state file', files[STATE_FILE], undefined);
}

// ---------------------------------------------------------------------------
section('11. An empty roster aborts rather than wiping the suppression state');
{
  const files = {};
  const { m, sent, state } = harness(files, { members: { '1': noRecovery('1') } });
  m.notifyRecoveryEmailCompliance();
  const good = files[STATE_FILE];

  state.members = {};
  const summary = m.notifyRecoveryEmailCompliance();

  check('aborts', summary.aborted, true);
  check('sends nothing', sent.length, 1);
  // Wiping here would drop every suppression window, so the next healthy run
  // would re-mail the entire wing.
  check('leaves state intact', files[STATE_FILE], good);
}

// ---------------------------------------------------------------------------
section('12. A corrupt state file refuses to run rather than re-reporting everyone');
{
  const files = {};
  files[STATE_FILE] = '{ this is not json';
  const { m, sent } = harness(files, { members: { '1': noRecovery('1') } });

  let threw = false;
  try { m.notifyRecoveryEmailCompliance(); } catch (e) { threw = true; }
  check('throws', threw, true);
  check('sends nothing', sent.length, 0);
}

// ---------------------------------------------------------------------------
section('13. installRecoveryComplianceMonthlyTrigger dedupes and schedules monthly');
{
  const created = [];
  const deleted = [];
  const triggers = [
    { getHandlerFunction: () => 'notifyRecoveryEmailCompliance' },
    { getHandlerFunction: () => 'someOtherJob' },
    { getHandlerFunction: () => 'notifyRecoveryEmailCompliance' }
  ];
  const ScriptApp = {
    getProjectTriggers: () => triggers,
    deleteTrigger: t => deleted.push(t.getHandlerFunction()),
    newTrigger: fn => {
      const b = {
        timeBased: () => b,
        onMonthDay: d => { b._day = d; return b; },
        atHour: h => { b._hour = h; return b; },
        create: () => created.push({ fn: fn, day: b._day, hour: b._hour })
      };
      return b;
    }
  };
  const mod = loadModule(MODULE, { Logger: makeLogger().logger, ScriptApp: ScriptApp },
    ['installRecoveryComplianceMonthlyTrigger']);

  mod.installRecoveryComplianceMonthlyTrigger();

  check('removes both existing triggers', deleted,
    ['notifyRecoveryEmailCompliance', 'notifyRecoveryEmailCompliance']);
  check('leaves unrelated triggers alone', deleted.indexOf('someOtherJob'), -1);
  check('creates exactly one', created.length, 1);
  check('on the 1st, at 07:00',
    created[0], { fn: 'notifyRecoveryEmailCompliance', day: 1, hour: 7 });
}

done();
