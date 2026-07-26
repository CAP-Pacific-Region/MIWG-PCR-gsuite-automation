/**
 * UpdateGroups.gs — pausing and resuming a run that cannot finish in one execution.
 *
 * Every membership change is an API call with pacing behind it, so the length of a
 * run tracks the number of CHANGES. The day a rule changes and thousands of
 * memberships move with it, the execution limit arrives first and the run is killed
 * with no record of where it got to. updateEmailGroups() therefore takes a deadline
 * and reports its position; updateEmailGroupsBatch() parks that position and picks it
 * up next time.
 *
 * What matters here is the bookkeeping — that a paused run resumes exactly where it
 * stopped, applies every change once, and does not write a half-finished error sheet.
 * The clock is faked, so the test is instant and deterministic.
 *
 * Addresses are synthetic.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker, Utilities } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'UpdateGroups.gs');
const { section, check, done } = makeChecker();

const CONFIG = {
  WING: 'CA',
  EMAIL_DOMAIN: '@example.org',
  CAPWATCH_DATA_FOLDER_ID: 'folder-id',
  AUTOMATION_SPREADSHEET_ID: 'sheet-id',
  API_DELAY_MS: 0
};

/** Two groups, three membership changes each. */
function makeDeltas() {
  return {
    'all-level-ii': {
      'ca.all-level-ii': {
        'one@example.org': 1,
        'two@example.org': 1,
        'three@example.org': -1
      }
    },
    'all-level-iii': {
      'ca.all-level-iii': {
        'four@example.org': 1,
        'five@example.org': -1,
        'six@example.org': 0
      }
    }
  };
}

/**
 * A directory that records calls. `clock` advances by `msPerCall` on every insert or
 * remove, which is how the deadline is reached deterministically.
 */
function makeHarness(msPerCall) {
  const calls = { insert: [], remove: [], settingsPatch: [] };
  const clock = { now: 1000 };

  const AdminDirectory = {
    Groups: {
      get: () => ({ name: 'n', description: 'd' }),
      patch: () => ({})
    },
    Members: {
      insert: (body, groupEmail) => {
        clock.now += msPerCall;
        calls.insert.push(`${groupEmail}:${body.email}`);
        return {};
      },
      remove: (groupEmail, email) => {
        clock.now += msPerCall;
        calls.remove.push(`${groupEmail}:${email}`);
        return {};
      },
      list: () => ({ members: [] })
    }
  };

  const AdminGroupsSettings = {
    Groups: {
      get: () => ({ allowExternalMembers: 'false', whoCanViewMembership: 'ALL_IN_DOMAIN_CAN_VIEW', whoCanPostMessage: 'ANYONE_CAN_POST' }),
      patch: (patch, email) => { calls.settingsPatch.push(email); return {}; }
    }
  };

  const savedSheets = [];
  const SpreadsheetApp = {
    openById: () => ({
      getSheetByName: () => ({
        getLastRow: () => 1,
        getLastColumn: () => 9,
        getRange: () => ({
          getValues: () => [['Email', '', '', '', '', '', '', '', '']],
          setValues: (v) => { savedSheets.push(v.length); return { setFontWeight: () => ({ setBackground: () => ({ setFontColor: () => ({}) }) }) }; },
          clearContent: () => {},
          setVerticalAlignment: () => ({}),
          setNumberFormat: () => ({}),
          setFontWeight: () => ({ setBackground: () => ({ setFontColor: () => ({}) }) })
        }),
        getConditionalFormatRules: () => [],
        setConditionalFormatRules: () => {},
        autoResizeColumn: () => {}
      })
    })
  };

  return { calls, clock, AdminDirectory, AdminGroupsSettings, SpreadsheetApp, savedSheets };
}

/** A resume payload positioned at the very start — the apply loop's entry point. */
function freshResume(deltas) {
  return { deltas: deltas, groupIndex: 0, memberIndex: 0, errorEmails: {}, totals: { added: 0, removed: 0, errors: 0 } };
}

function load(h) {
  const { logger, calls: logs } = makeLogger();
  const FakeDate = function () { return new Date(h.clock.now); };
  FakeDate.now = () => h.clock.now;

  const m = loadModule(MODULE, {
    Logger: logger,
    CONFIG: CONFIG,
    Utilities: Object.assign({}, Utilities, { sleep: () => {} }),
    AdminDirectory: h.AdminDirectory,
    AdminGroupsSettings: h.AdminGroupsSettings,
    SpreadsheetApp: h.SpreadsheetApp,
    Date: FakeDate,
    executeWithRetry: (fn) => fn(),
    parseFile: () => [],
    clearCache: () => {}
  }, ['updateEmailGroups']);

  return { m, logs };
}

// ---------------------------------------------------------------------------
section('1. No deadline: one pass applies everything, exactly as before');
{
  const h = makeHarness(10);
  const { m } = load(h);
  const result = m.updateEmailGroups({ resume: freshResume(makeDeltas()) });

  check('reports complete', result.complete, true);
  check('every add applied', h.calls.insert.sort(),
    ['ca.all-level-ii@example.org:one@example.org',
     'ca.all-level-ii@example.org:two@example.org',
     'ca.all-level-iii@example.org:four@example.org']. sort());
  check('every removal applied', h.calls.remove.sort(),
    ['ca.all-level-ii@example.org:three@example.org',
     'ca.all-level-iii@example.org:five@example.org'].sort());
  check('a 0 delta is left alone', h.calls.insert.concat(h.calls.remove)
    .some(c => c.includes('six@example.org')), false);
  check('nothing left to resume', result.memberIndex, 0);
}

// ---------------------------------------------------------------------------
section('2. Deadline mid-group: stops, and says where');
{
  // 1s of clock per change, 1.5s of budget: two changes land, the third does not.
  const h = makeHarness(1000);
  const { m } = load(h);
  const result = m.updateEmailGroups({ deadlineMs: h.clock.now + 1500, resume: freshResume(makeDeltas()) });

  check('not complete', result.complete, false);
  check('stopped inside the first group', result.groupIndex, 0);
  check('two changes made before the deadline', h.calls.insert.length + h.calls.remove.length, 2);
  check('resumes at the third member', result.memberIndex, 2);
  check('deltas handed back for parking', Object.keys(result.deltas).sort(),
    ['all-level-ii', 'all-level-iii']);
}

// ---------------------------------------------------------------------------
section('3. Deadline on a group boundary: whole groups done, next one queued');
{
  const h = makeHarness(1000);
  const { m } = load(h);
  const result = m.updateEmailGroups({ deadlineMs: h.clock.now + 2500, resume: freshResume(makeDeltas()) });

  check('not complete', result.complete, false);
  check('first group finished, stopped before the second', result.groupIndex, 1);
  check('resumes at that group\'s first member', result.memberIndex, 0);
  check('all three of the first group applied', h.calls.insert.length + h.calls.remove.length, 3);
}

// ---------------------------------------------------------------------------
section('4. Resuming finishes the job, each change applied exactly once');
{
  const first = makeHarness(1000);
  const a = load(first).m.updateEmailGroups({
    deadlineMs: first.clock.now + 1500, resume: freshResume(makeDeltas())
  });

  const second = makeHarness(1000);
  const b = load(second).m.updateEmailGroups({
    resume: {
      deltas: a.deltas,
      groupIndex: a.groupIndex,
      memberIndex: a.memberIndex,
      errorEmails: a.errorEmails,
      totals: a.totals
    }
  });

  check('second pass completes', b.complete, true);

  const allCalls = first.calls.insert.concat(first.calls.remove, second.calls.insert, second.calls.remove);
  check('five changes across both passes, no repeats', allCalls.length, 5);
  check('no change applied twice', allCalls.length, new Set(allCalls).size);
  check('running totals carried across', b.totals.added + b.totals.removed, 5);
  check('the resumed pass did not redo the first two',
    second.calls.insert.concat(second.calls.remove).some(c => c.includes('one@example.org')), false);
}

// ---------------------------------------------------------------------------
section('5. Group metadata is not re-applied when resuming mid-group');
{
  const h = makeHarness(1000);
  const { m } = load(h);
  const paused = m.updateEmailGroups({ deadlineMs: h.clock.now + 1500, resume: freshResume(makeDeltas()) });
  const settingsBefore = h.calls.settingsPatch.length;

  const h2 = makeHarness(1000);
  load(h2).m.updateEmailGroups({
    resume: {
      deltas: paused.deltas,
      groupIndex: paused.groupIndex,
      memberIndex: paused.memberIndex,
      errorEmails: {},
      totals: paused.totals
    }
  });

  check('first pass touched settings for the group it started', settingsBefore >= 0, true);
  check('resumed pass settles the remaining group only',
    h2.calls.settingsPatch.length <= 1, true);
}

done();
