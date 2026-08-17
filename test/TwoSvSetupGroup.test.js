/**
 * TwoSvSetupGroup.gs — who leaves the 2SV holding pen, and when.
 *
 * Membership of this group is a security EXEMPTION, so both directions of being
 * wrong are pinned here:
 *
 *   - leaving someone exempt forever (the bug this module exists to fix)
 *   - ending an exemption early, out from under a member mid-enrollment, on
 *     evidence the module does not actually have
 *
 * All addresses are synthetic.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'TwoSvSetupGroup.gs');
const { section, check, done } = makeChecker();

const m = loadModule(MODULE, {
  Logger: makeLogger().logger,
  CONFIG: { CAPWATCH_DATA_FOLDER_ID: 'folder', API_RETRY_ATTEMPTS: 3 },
  TENANT: { TWO_SV_SETUP_GROUP: 'ca.2sv-setup@example.org' },
  ERROR_CODES: { NOT_FOUND: 404 },
  DriveApp: {},
  MimeType: { PLAIN_TEXT: 'text/plain' },
  AdminDirectory: {},
  executeWithRetry: (fn) => fn(),
  console: console
}, ['evaluateTwoSvSetupGroup_', 'twoSvDaysBetween_', 'twoSvDateStamp_', 'TWO_SV_SETUP_CONFIG']);

const NOW = new Date('2026-08-16T09:00:00Z');
const TODAY = '2026-08-16';

const member = (email, enrolled) => ({ email: email, enrolledIn2Sv: enrolled });
const evaluate = (members, firstSeen, opts) =>
  m.evaluateTwoSvSetupGroup_(members, firstSeen || {}, NOW, opts);

const reasonFor = (result, email) => {
  const hit = result.remove.concat(result.keep).find(e => e.email === email);
  return hit ? hit.reason : 'absent';
};

// ---------------------------------------------------------------------------
section('1. The two ways out of the group');
{
  // Enrolled: the exemption did its job, and the day it was granted is irrelevant.
  const enrolled = evaluate([member('a.member@example.org', true)], { 'a.member@example.org': TODAY });
  check('enrolled member is removed', enrolled.remove.length, 1);
  check('...on the day they were added', enrolled.remove[0].reason, '2sv-enrolled');
  check('...and nobody is kept', enrolled.keep.length, 0);

  // Not enrolled, window elapsed: the exemption expired.
  const expired = evaluate([member('b.member@example.org', false)], { 'b.member@example.org': '2026-08-09' });
  check('seven days without 2SV is removed', expired.remove.length, 1);
  check('...for the other reason', expired.remove[0].reason, 'grace-expired');
  check('...and the report says how long', expired.remove[0].days, 7);

  // Not enrolled, still inside the window: left alone.
  const waiting = evaluate([member('c.member@example.org', false)], { 'c.member@example.org': '2026-08-10' });
  check('six days in, still exempt', waiting.remove.length, 0);
  check('...and counted as kept', waiting.keep.length, 1);
  check('...with the reason', waiting.keep[0].reason, 'within-grace');
}

// ---------------------------------------------------------------------------
section('2. Never end an exemption on evidence we do not have');
{
  // A directory read that failed says nothing about enrollment. Treating that
  // as "not enrolled" is harmless (they stay); treating a failure as enrolled
  // would strip the exemption of someone who may be mid-setup.
  const unknown = evaluate([member('d.member@example.org', null)], { 'd.member@example.org': '2026-08-15' });
  check('unreadable 2SV state keeps the member', unknown.remove.length, 0);
  check('...and says so, rather than claiming they are unenrolled',
    unknown.keep[0].reason, 'within-grace-2sv-unknown');

  // The clock, however, is ours — it needs no directory read, so an account
  // nobody can read about does not stay exempt forever either.
  const stale = evaluate([member('d.member@example.org', null)], { 'd.member@example.org': '2026-08-01' });
  check('but the grace clock still expires it', stale.remove.length, 1);
  check('...on time, not on 2SV', stale.remove[0].reason, 'grace-expired');
}

// ---------------------------------------------------------------------------
section('3. First sight starts the clock — it never ends it');
{
  // Day one of the ledger: everyone currently in the group is new to it. None of
  // them may be removed for age, or the module's first run would empty a group
  // full of people mid-enrollment.
  const firstRun = evaluate([
    member('e.member@example.org', false),
    member('f.member@example.org', false)
  ], {});
  check('nobody is aged out on the first run', firstRun.remove.length, 0);
  check('...both are stamped today', firstRun.firstSeen['e.member@example.org'], TODAY);
  check('...and counted as newly tracked', firstRun.added, 2);

  // An existing stamp is never overwritten — otherwise the window would reset
  // nightly and nobody would ever age out.
  const kept = evaluate([member('e.member@example.org', false)], { 'e.member@example.org': '2026-08-12' });
  check('an existing first-seen date survives the run',
    kept.firstSeen['e.member@example.org'], '2026-08-12');

  // Someone who left the group drops out of the ledger, so a later re-add gets a
  // fresh window instead of inheriting an expired one and being removed at once.
  const departed = evaluate([member('e.member@example.org', false)], {
    'e.member@example.org': '2026-08-12',
    'gone.member@example.org': '2026-07-01'
  });
  check('a member no longer in the group is dropped from the ledger',
    departed.firstSeen['gone.member@example.org'], undefined);
}

// ---------------------------------------------------------------------------
section('4. Boundaries and housekeeping');
{
  check('the window is seven days', m.TWO_SV_SETUP_CONFIG.GRACE_DAYS, 7);

  const sixDays = evaluate([member('g.member@example.org', false)], { 'g.member@example.org': '2026-08-10' });
  check('day six keeps', sixDays.remove.length, 0);
  const sevenDays = evaluate([member('g.member@example.org', false)], { 'g.member@example.org': '2026-08-09' });
  check('day seven removes', sevenDays.remove.length, 1);

  check('the window is overridable for a one-off run',
    evaluate([member('g.member@example.org', false)], { 'g.member@example.org': '2026-08-13' },
      { graceDays: 3 }).remove.length,
    1);

  // Addresses come back from the Directory in whatever case they were typed in;
  // the ledger is keyed on one form so a member is not tracked twice.
  const mixedCase = evaluate([member('H.Member@Example.org', false)], { 'h.member@example.org': '2026-08-01' });
  check('lookup is case-insensitive', mixedCase.remove.length, 1);
  check('...and the ledger key is lowercased',
    Object.keys(mixedCase.firstSeen)[0], 'h.member@example.org');

  check('a blank address is skipped rather than tracked',
    evaluate([member('', false)], {}).keep.length, 0);

  // A future or malformed stamp must not read as an enormous elapsed age.
  check('a future date is zero days, not negative', m.twoSvDaysBetween_('2026-09-01', NOW), 0);
  check('a malformed date is zero days', m.twoSvDaysBetween_('not-a-date', NOW), 0);
  check('date stamps are UTC dates', m.twoSvDateStamp_(new Date('2026-08-16T23:30:00Z')), '2026-08-16');
}

// ---------------------------------------------------------------------------
section('5. The file must load before config.gs has run');
{
  // Apps Script evaluates every file's top level in editor order, and
  // subfoldered files sort ahead of config.gs — so a top-level read of TENANT
  // throws ReferenceError before TENANT exists, and a top-level throw aborts the
  // load of the WHOLE project, taking every unrelated function with it. That is
  // exactly what happened on live seniors, 2026-08-16.
  //
  // Loading the module with NO tenant config injected reproduces that ordering:
  // it must evaluate cleanly, and only blow up if something actually asks for
  // the group.
  let loadedWithoutTenant = true;
  try {
    loadModule(MODULE, {
      Logger: makeLogger().logger,
      CONFIG: {},
      DriveApp: {},
      MimeType: {},
      AdminDirectory: {},
      executeWithRetry: (fn) => fn(),
      console: console
    }, ['evaluateTwoSvSetupGroup_']);
  } catch (e) {
    loadedWithoutTenant = false;
  }
  check('the module evaluates with TENANT not yet defined', loadedWithoutTenant, true);
}

done();
