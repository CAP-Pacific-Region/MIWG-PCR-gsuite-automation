/**
 * LSCodeNotify.gs — end-to-end.
 *
 * Drives the real notifyLSCodeChanges() against a fake Drive and a Gmail stub
 * that records every send. The question it exists to answer is the one that
 * decides whether this module is safe to arm: does the FIRST run mail anyone?
 *
 * It must not. On a wing where most senior members already hold an 'A', a first
 * run that treated those as news would mail every squadron commander their entire
 * roster — the module's whole first-seen rule exists to prevent exactly that, and
 * this is what holds it in place.
 *
 * Run: npm test
 */
const path = require('path');
const {
  loadModule, makeLogger, makeDrive, makeGmail, makeChecker, Session, Utilities
} = require('./helpers/apps-script');
const {
  memberFile, emptyMemberFile, SQUADRONS, COMMANDER_ROWS, COMMANDER_EMAILS
} = require('./helpers/capwatch-fixtures');

const MODULE = path.join(__dirname, '..', 'src', 'notifications', 'LSCodeNotify.gs');
const { section, check, done } = makeChecker();

/**
 * Builds a module instance over the given Drive contents.
 *
 * @param {Object} files - filename -> content, mutated in place by the module
 * @param {Object} [opts] - { enabled: boolean, failFrom: string[] }
 * @returns {Object} { m, sent, log }
 */
function harness(files, opts) {
  const enabled = !opts || opts.enabled !== false;
  const { logger, calls } = makeLogger();
  const { gmail, sent } = makeGmail({ failFrom: (opts && opts.failFrom) || [] });

  const m = loadModule(MODULE, {
    Logger: logger,
    Session: Session,
    Utilities: Utilities,
    DriveApp: makeDrive(files),
    GmailApp: gmail,
    CONFIG: { CAPWATCH_DATA_FOLDER_ID: 'fake-folder-id' },
    PROFILE_: { RUN_LSCODE_NOTIFICATIONS: enabled, EXCLUDED_ORG_IDS: [] },
    TENANT_PROFILE: enabled ? 'seniors' : 'cadets',
    clearCache: () => {},
    getSquadrons: () => SQUADRONS,
    parseFile: name => (name === 'Commanders' ? COMMANDER_ROWS : []),
    createEmailMap: () => COMMANDER_EMAILS,
    executeWithRetry: fn => fn(),
    AUTOMATION_SENDER_EMAIL: 'automation@cawgcap.org',
    SENDER_NAME: 'CAP Information Technology',
    ITSUPPORT_EMAIL: 'it.support@cawgcap.org',
    TEST_EMAIL: 'test@cawgcap.org'
  }, ['notifyLSCodeChanges', 'previewLSCodeChanges', 'armLSCodeLiveTest']);

  return { m, sent, log: calls };
}

// A wing that looks like the real one: most seniors already cleared.
const ESTABLISHED_WING = [
  { capid: '1', orgid: '100', lscode: 'A' },
  { capid: '2', orgid: '100', lscode: 'A' },
  { capid: '3', orgid: '100', lscode: '' },
  { capid: '4', orgid: '200', lscode: 'A' },
  { capid: '5', orgid: '200', lscode: '' }
];

// ---------------------------------------------------------------------------
section('1. First run on a wing where members already hold an A');
{
  const files = { 'Member.txt': memberFile(ESTABLISHED_WING) };
  const { m, sent } = harness(files);
  const summary = m.notifyLSCodeChanges();

  check('sends NO email to anyone', sent.length, 0);
  check('counts no grants', summary.granted, 0);
  check('counts no revocations', summary.revoked, 0);
  check('records all five as first-seen', summary.firstSeen, 5);
  check('reports itself as a baseline', summary.baseline, true);
  check('accounts for every member read',
    summary.firstSeen + summary.unchanged + summary.granted + summary.revoked + summary.ignored,
    summary.members);

  const state = JSON.parse(files['LSCodeState.txt']);
  check('writes a state file', typeof files['LSCodeState.txt'], 'string');
  check('at version 2', state.version, 2);
  check('holding all five members', Object.keys(state.members).sort(), ['1', '2', '3', '4', '5']);
  check('with the A-holders recorded as A',
    [state.members['1'].c, state.members['2'].c, state.members['4'].c], ['A', 'A', 'A']);
  check('and the uncleared recorded as blank',
    [state.members['3'].c, state.members['5'].c], ['', '']);
}

// ---------------------------------------------------------------------------
section('2. Second run, nothing changed — the A-holders are left alone');
{
  const files = { 'Member.txt': memberFile(ESTABLISHED_WING) };
  const { m, sent } = harness(files);
  m.notifyLSCodeChanges();
  const afterBaseline = files['LSCodeState.txt'];
  const summary = m.notifyLSCodeChanges();

  check('still sends nothing', sent.length, 0);
  check('counts all five as unchanged', summary.unchanged, 5);
  check('re-counts nobody as first-seen', summary.firstSeen, 0);
  check('leaves state identical', files['LSCodeState.txt'], afterBaseline);
}

// ---------------------------------------------------------------------------
section('3. A member clears their check — one email, to one commander');
{
  const files = { 'Member.txt': memberFile(ESTABLISHED_WING) };
  const { m, sent } = harness(files);
  m.notifyLSCodeChanges();
  check('baseline sent nothing', sent.length, 0);

  // Member 3, in Alpha Squadron, clears.
  files['Member.txt'] = memberFile(
    ESTABLISHED_WING.map(s => (s.capid === '3' ? { ...s, lscode: 'A' } : s))
  );
  const summary = m.notifyLSCodeChanges();

  check('sends exactly one email', sent.length, 1);
  check('to that squadron\'s commander', sent[0].to, 'ada.alpha@cawgcap.org');
  check('counts one grant', summary.granted, 1);
  check('subject names the charter', sent[0].subject.includes('PCR-CA-070'), true);
  check('names the member who changed', /<td>3<\/td>/.test(sent[0].html), true);
  check('does not name the untouched A-holders',
    /<td>1<\/td>|<td>2<\/td>|<td>4<\/td>/.test(sent[0].html), false);
  check('does not mail the other squadron',
    sent.filter(s => s.to === 'ben.bravo@cawgcap.org').length, 0);
  check('dates the change to a window, not a day',
    /\d+ \w+ \d{4}|\d+–\d+ \w+ \d{4}/.test(sent[0].html), true);
}

// ---------------------------------------------------------------------------
section('4. A member loses their check — reported, and marked for verification');
{
  const files = {
    'Member.txt': memberFile([
      { capid: '1', orgid: '100', lscode: 'A' },
      { capid: '2', orgid: '100', lscode: 'A' }
    ])
  };
  const { m, sent } = harness(files);
  m.notifyLSCodeChanges();

  files['Member.txt'] = memberFile([
    { capid: '1', orgid: '100', lscode: '' },
    { capid: '2', orgid: '100', lscode: 'A' }
  ]);
  const summary = m.notifyLSCodeChanges();

  check('counts one revocation', summary.revoked, 1);
  check('sends one email', sent.length, 1);
  check('subject says no longer current', sent[0].subject.includes('no longer current'), true);
  check('body tells the CC to verify in eServices first',
    sent[0].html.includes('verify in eServices'), true);
}

// ---------------------------------------------------------------------------
section('5. A new member who arrives already cleared is not news');
{
  const files = { 'Member.txt': memberFile([{ capid: '1', orgid: '100', lscode: 'A' }]) };
  const { m, sent } = harness(files);
  m.notifyLSCodeChanges();

  // Member 9 transfers in, already holding a clearance.
  files['Member.txt'] = memberFile([
    { capid: '1', orgid: '100', lscode: 'A' },
    { capid: '9', orgid: '100', lscode: 'A' }
  ]);
  const summary = m.notifyLSCodeChanges();

  check('sends nothing', sent.length, 0);
  check('counts them first-seen, not granted', [summary.firstSeen, summary.granted], [1, 0]);
  check('but records them for next time',
    JSON.parse(files['LSCodeState.txt']).members['9'].c, 'A');
}

// ---------------------------------------------------------------------------
section('6. previewLSCodeChanges() sends nothing, writes nothing, consumes nothing');
{
  const files = {
    'Member.txt': memberFile([
      { capid: '1', orgid: '100', lscode: 'A' },
      { capid: '2', orgid: '100', lscode: '' }
    ])
  };
  const { m, sent } = harness(files);
  m.notifyLSCodeChanges();

  files['Member.txt'] = memberFile([
    { capid: '1', orgid: '100', lscode: 'A' },
    { capid: '2', orgid: '100', lscode: 'A' }
  ]);
  const beforePreview = files['LSCodeState.txt'];

  const preview = m.previewLSCodeChanges();
  check('preview sends no mail', sent.length, 0);
  check('preview still reports the change', preview.granted, 1);
  check('preview leaves state untouched', files['LSCodeState.txt'], beforePreview);

  const real = m.notifyLSCodeChanges();
  check('the real run afterwards still delivers it', sent.length, 1);
  check('and still counts it', real.granted, 1);
}

// ---------------------------------------------------------------------------
section('7. A truncated Member.txt aborts rather than wiping state');
{
  const files = {
    'Member.txt': memberFile([
      { capid: '1', orgid: '100', lscode: 'A' },
      { capid: '2', orgid: '100', lscode: 'A' }
    ])
  };
  const { m, sent } = harness(files);
  m.notifyLSCodeChanges();
  const good = files['LSCodeState.txt'];

  files['Member.txt'] = emptyMemberFile();
  const summary = m.notifyLSCodeChanges();

  check('aborts', summary.aborted, true);
  check('sends nothing', sent.length, 0);
  // If it wiped state here, the next healthy run would read the whole wing as
  // first-seen -- silent, but every pending change lost.
  check('does not wipe state', files['LSCodeState.txt'], good);
}

// ---------------------------------------------------------------------------
section('8. Inactive members are ignored');
{
  const files = {
    'Member.txt': memberFile([
      { capid: '1', orgid: '100', lscode: 'A' },
      { capid: '7', orgid: '100', lscode: '', status: 'EXPIRED' }
    ])
  };
  const { m } = harness(files);
  const summary = m.notifyLSCodeChanges();

  check('only the active member is read', summary.members, 1);
  check('the expired one is not in state',
    Object.keys(JSON.parse(files['LSCodeState.txt']).members), ['1']);
}

// ---------------------------------------------------------------------------
section('9. A disabled profile does nothing at all');
{
  const files = { 'Member.txt': memberFile([{ capid: '1', orgid: '100', lscode: 'A' }]) };
  const { m, sent } = harness(files, { enabled: false });
  const summary = m.notifyLSCodeChanges();

  check('skips', summary.skipped, true);
  check('sends nothing', sent.length, 0);
  check('writes no state file', files['LSCodeState.txt'], undefined);
}

// ---------------------------------------------------------------------------
section('10. armLSCodeLiveTest seeds one scoped change, then notify mails once');
{
  const files = { 'Member.txt': memberFile(ESTABLISHED_WING) };
  const { m, sent } = harness(files);
  m.notifyLSCodeChanges();            // real baseline: everyone recorded, silent
  check('baseline is silent', sent.length, 0);

  // Arm a test in Alpha Squadron (org 100). Member 1 holds an 'A', so this
  // records their prior value as blank -> the next run sees a GRANT.
  const armed = m.armLSCodeLiveTest('100');
  check('arming sends nothing on its own', sent.length, 0);
  check('it targets an org-100 member', ['1', '2', '3'].includes(armed.capid), true);
  check('picks a cleared member, so the test is a GRANT', armed.direction, 'GRANTED');
  check('reports the recipient it will mail', armed.recipient, 'ada.alpha@cawgcap.org');

  // Preview must show exactly one org would be mailed — the blast-radius gate.
  const preview = m.previewLSCodeChanges();
  check('preview would mail exactly one org', preview.sent, 1);
  check('preview still sends no mail', sent.length, 0);

  // The real run sends one digest, to that one commander.
  const summary = m.notifyLSCodeChanges();
  check('one email', sent.length, 1);
  check('to the armed org\'s commander', sent[0].to, 'ada.alpha@cawgcap.org');
  check('one grant', summary.granted, 1);

  // Self-heal: the seeded member is back at their true value, so a further run
  // is silent again.
  const after = m.notifyLSCodeChanges();
  check('state self-heals — next run silent', sent.length, 1);
  check('and everyone reads unchanged', after.unchanged, ESTABLISHED_WING.length);
}

// ---------------------------------------------------------------------------
section('11. armLSCodeLiveTest refuses an argument-less call');
{
  const files = { 'Member.txt': memberFile(ESTABLISHED_WING) };
  const { m } = harness(files);
  let threw = false;
  try { m.armLSCodeLiveTest(); } catch (e) { threw = true; }
  check('throws without an ORGID', threw, true);
}

// ---------------------------------------------------------------------------
section('12. A corrupt state file refuses to run rather than re-baselining');
{
  const files = {
    'Member.txt': memberFile([{ capid: '1', orgid: '100', lscode: 'A' }]),
    'LSCodeState.txt': '{ this is not json'
  };
  const { m, sent } = harness(files);

  let threw = false;
  try { m.notifyLSCodeChanges(); } catch (e) { threw = true; }
  // Treating unparseable state as a baseline would look identical to a clean
  // first run while silently discarding every pending change.
  check('throws', threw, true);
  check('sends nothing', sent.length, 0);
}

// ---------------------------------------------------------------------------
section('13. Wrong sender identity: digests fail, but the alarm still gets out');
{
  // Reproduce the live 2026-07-16 failure: the executing account cannot send as
  // automation@cawgcap.org, so every digest bounces with "Invalid argument".
  const files = { 'Member.txt': memberFile([{ capid: '1', orgid: '100', lscode: 'A' }]) };
  const { m, sent } = harness(files, { failFrom: ['automation@cawgcap.org'] });
  m.notifyLSCodeChanges();            // baseline, silent (no from used)

  // Seed a real change, then run: the digest is attempted with from:
  // automation@, which the stub rejects exactly as the live tenant did.
  m.armLSCodeLiveTest('100');
  const summary = m.notifyLSCodeChanges();

  check('the digest failed (wrong identity)', summary.failedOrgs.length, 1);
  check('nothing was recorded as sent', summary.sent, 0);

  // The alarm to IT is the whole point: it must have gone out despite the same
  // misconfiguration, because it does not set `from`.
  const alarm = sent.filter(s => s.to === 'it.support@cawgcap.org');
  check('the IT failure-summary was delivered', alarm.length, 1);
  check('and it carried no from override', alarm[0].from, undefined);
  check('no digest slipped through with the bad from',
    sent.filter(s => s.from === 'automation@cawgcap.org').length, 0);
}

// ---------------------------------------------------------------------------
section('14. installLSCodeWeeklyTrigger dedupes and schedules weekly');
{
  const created = [];
  const deleted = [];
  // Two existing notify triggers (a stale duplicate) plus an unrelated one.
  const triggers = [
    { getHandlerFunction: () => 'notifyLSCodeChanges' },
    { getHandlerFunction: () => 'someOtherJob' },
    { getHandlerFunction: () => 'notifyLSCodeChanges' }
  ];
  const ScriptApp = {
    WeekDay: { MONDAY: 'MONDAY' },
    getProjectTriggers: () => triggers,
    deleteTrigger: t => deleted.push(t.getHandlerFunction()),
    newTrigger: fn => {
      const b = {
        timeBased: () => b,
        onWeekDay: d => { b._day = d; return b; },
        atHour: h => { b._hour = h; return b; },
        create: () => created.push({ fn: fn, day: b._day, hour: b._hour })
      };
      return b;
    }
  };
  const mod = loadModule(MODULE, { Logger: makeLogger().logger, ScriptApp: ScriptApp },
    ['installLSCodeWeeklyTrigger']);

  mod.installLSCodeWeeklyTrigger();

  check('removes both existing notify triggers', deleted, ['notifyLSCodeChanges', 'notifyLSCodeChanges']);
  check('leaves unrelated triggers alone', deleted.indexOf('someOtherJob'), -1);
  check('creates exactly one trigger', created.length, 1);
  check('scheduled weekly, Monday, hour 7, for notifyLSCodeChanges',
    created[0], { fn: 'notifyLSCodeChanges', day: 'MONDAY', hour: 7 });
}

done();
