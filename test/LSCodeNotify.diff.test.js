/**
 * LSCodeNotify.gs — the diff, the state rollback, and window rendering, in
 * isolation. LSCodeNotify.run.test.js covers the same module end-to-end; this
 * one pins the decisions directly, so a failure names the rule that broke rather
 * than just "no email arrived".
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker, Session, Utilities } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'notifications', 'LSCodeNotify.gs');
const { section, check, done } = makeChecker();

const { logger, calls } = makeLogger();
const m = loadModule(MODULE, {
  Logger: logger,
  Session: Session,
  Utilities: Utilities
}, ['diffLSCodes_', 'revertOrgState_', 'formatWindow_', 'isoDate_']);

const TODAY = '2026-07-15';
const LAST_WEEK = '2026-07-08';

/** A recorded state entry: value `c`, last confirmed on `seen`. */
const rec = (c, seen) => ({ c, seen: seen || LAST_WEEK });

/** A member as readMemberLSCodes_ would produce them. */
const mk = (capid, lscode, orgid) => ({
  capid,
  lscode,
  orgid: orgid || '100',
  firstName: 'First' + capid,
  lastName: 'Last' + capid,
  rank: 'Capt',
  type: 'SENIOR',
  orgName: 'Alpha Squadron',
  charter: 'PCR-CA-070'
});

// ---------------------------------------------------------------------------
section('1. No prior state: everyone is first-seen, nobody is notified');
{
  const current = { '1': mk('1', 'A'), '2': mk('2', ''), '3': mk('3', 'A') };
  const d = m.diffLSCodes_(current, {}, TODAY);
  check('nothing to send', Object.keys(d.byOrg).length, 0);
  check('all three first-seen', d.firstSeen, 3);
  check('no grants', d.granted, 0);
  check('no revocations', d.revoked, 0);
}

// ---------------------------------------------------------------------------
section('2. blank -> A is a grant, carrying its detection window');
{
  const d = m.diffLSCodes_({ '1': mk('1', 'A') }, { '1': rec('') }, TODAY);
  const change = d.byOrg['100'][0];
  check('one grant', d.granted, 1);
  check('no revocation', d.revoked, 0);
  check('grouped by org', Object.keys(d.byOrg), ['100']);
  check('direction', change.direction, 'GRANTED');
  check('from/to', [change.from, change.to], ['', 'A']);
  check('window spans last confirmation to today',
    [change.windowFrom, change.windowTo], [LAST_WEEK, TODAY]);
}

// ---------------------------------------------------------------------------
section('3. A -> blank is a revocation');
{
  const d = m.diffLSCodes_({ '1': mk('1', '') }, { '1': rec('A') }, TODAY);
  check('one revocation', d.revoked, 1);
  check('no grant', d.granted, 0);
  check('direction', d.byOrg['100'][0].direction, 'REVOKED');
}

// ---------------------------------------------------------------------------
section('4. Unchanged members are silent');
{
  const d = m.diffLSCodes_(
    { '1': mk('1', 'A'), '2': mk('2', '') },
    { '1': rec('A'), '2': rec('') },
    TODAY
  );
  check('nothing to send', Object.keys(d.byOrg).length, 0);
  check('both unchanged', d.unchanged, 2);
}

// ---------------------------------------------------------------------------
section('5. A new member already holding a clearance is not news');
{
  const d = m.diffLSCodes_({ '1': mk('1', 'A'), '99': mk('99', 'A') }, { '1': rec('A') }, TODAY);
  check('silent', Object.keys(d.byOrg).length, 0);
  check('first-seen', d.firstSeen, 1);
}

// ---------------------------------------------------------------------------
section('6. Changes group per org, so each CC gets one digest');
{
  const d = m.diffLSCodes_(
    { '1': mk('1', 'A', '100'), '2': mk('2', 'A', '100'), '3': mk('3', 'A', '200') },
    { '1': rec(''), '2': rec(''), '3': rec('') },
    TODAY
  );
  check('two orgs', Object.keys(d.byOrg).sort(), ['100', '200']);
  check('two in the first', d.byOrg['100'].length, 2);
  check('one in the second', d.byOrg['200'].length, 1);
  check('three grants', d.granted, 3);
}

// ---------------------------------------------------------------------------
section('7. A departed member produces nothing');
{
  const d = m.diffLSCodes_({ '1': mk('1', 'A') }, { '1': rec('A'), '77': rec('A') }, TODAY);
  check('no phantom change', Object.keys(d.byOrg).length, 0);
  check('only the present member counted', d.unchanged, 1);
}

// ---------------------------------------------------------------------------
section('8. Changes that never cross the cleared boundary warn, not mail');
{
  calls.warn.length = 0;
  const d = m.diffLSCodes_({ '1': mk('1', 'X') }, { '1': rec('Y') }, TODAY);
  check('no mail for X -> Y', Object.keys(d.byOrg).length, 0);
  check('but it warns', calls.warn.length, 1);
}
{
  // Anything that is not 'A' counts as uncleared, so an unknown code -> A is
  // still a real grant rather than something to swallow.
  const d = m.diffLSCodes_({ '1': mk('1', 'A') }, { '1': rec('X') }, TODAY);
  check('unknown -> A is a grant', d.granted, 1);
}

// ---------------------------------------------------------------------------
section('9. revertOrgState_ leaves a failed org pending, window intact');
{
  const prior = { '1': rec(''), '2': rec('A') };
  const next = {
    '1': { c: 'A', seen: TODAY },
    '2': { c: '', seen: TODAY },
    '3': { c: 'A', seen: TODAY }
  };
  m.revertOrgState_(next, [{ capid: '1' }, { capid: '2' }], prior);
  check('values roll back', [next['1'].c, next['2'].c], ['', 'A']);
  check('and so do their seen dates', [next['1'].seen, next['2'].seen], [LAST_WEEK, LAST_WEEK]);
  check('untouched members keep today', next['3'], { c: 'A', seen: TODAY });
}
{
  const next = { '9': { c: 'A', seen: TODAY } };
  m.revertOrgState_(next, [{ capid: '9' }], {});
  check('a first-seen member is removed outright',
    Object.prototype.hasOwnProperty.call(next, '9'), false);
}

// ---------------------------------------------------------------------------
section('10. The retry that the per-member window exists to get right');
{
  // Wk1 (Jul 8): confirmed blank.
  // Wk2 (Jul 15): change detected, digest FAILS to send.
  // Wk3 (Jul 22): retry must still report Jul 8-15, not Jul 15-22.
  const current = { '1': mk('1', 'A') };
  const priorWk2 = { '1': rec('', '2026-07-08') };

  const wk2 = m.diffLSCodes_(current, priorWk2, '2026-07-15');
  check('week 2 detects it', wk2.granted, 1);
  check('week 2 reports the right window',
    [wk2.byOrg['100'][0].windowFrom, wk2.byOrg['100'][0].windowTo], ['2026-07-08', '2026-07-15']);

  const nextWk2 = { '1': { c: 'A', seen: '2026-07-15' } };
  m.revertOrgState_(nextWk2, wk2.byOrg['100'], priorWk2);

  const wk3 = m.diffLSCodes_(current, nextWk2, '2026-07-22');
  check('week 3 re-detects it', wk3.granted, 1);
  check('week 3 still reports the week it was really detected',
    wk3.byOrg['100'][0].windowFrom, '2026-07-08');
}

// ---------------------------------------------------------------------------
section('11. Window rendering');
{
  check('same month collapses', m.formatWindow_('2026-07-08', '2026-07-15'), '8–15 Jul 2026');
  check('month boundary', m.formatWindow_('2026-06-28', '2026-07-05'), '28 Jun – 5 Jul 2026');
  check('year boundary', m.formatWindow_('2025-12-28', '2026-01-04'), '28 Dec 2025 – 4 Jan 2026');
  check('a single day', m.formatWindow_('2026-07-15', '2026-07-15'), '15 Jul 2026');
  check('unparseable renders empty, not "Invalid Date"', m.formatWindow_('', '2026-07-15'), '');
}

done();
