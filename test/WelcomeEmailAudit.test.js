/**
 * WelcomeEmailAudit.gs — the verdicts.
 *
 * The audit's whole value is that MISSED means something. Two ways it could
 * become worthless, both pinned here:
 *
 *   - accusing members it cannot actually judge (an unseeded or lost ledger must
 *     produce UNKNOWN, never a wing-wide list of MISSED)
 *   - burying real findings (the seed vouches only for accounts with login
 *     history; never-signed-in accounts stay UNKNOWN rather than being marked
 *     welcomed on no evidence)
 *
 * All names and addresses are synthetic.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'WelcomeEmailAudit.gs');
const { section, check, done } = makeChecker();

const m = loadModule(MODULE, {
  Logger: makeLogger().logger,
  CONFIG: {},
  DriveApp: {},
  MailApp: {},
  ScriptApp: {},
  Utilities: {},
  Session: {},
  ITSUPPORT_EMAIL: 'it@example.org',
  getMembers: () => ({}),
  getActiveUsers: () => [],
  console: console
}, ['classifyWelcomeAudit_', 'welcomeAuditAccountMap_', 'welcomeAuditTimestamp_']);

const NOW = new Date('2026-07-26T12:00:00Z');
const BASELINE = '2026-06-01T00:00:00.000Z';
const NEVER = '1970-01-01T00:00:00.000Z';   // how Google reports "never signed in"

const members = {
  '100001': { rank: 'Capt', firstName: 'Ada', lastName: 'Nwosu', type: 'SENIOR', charter: 'PCR-XX-001', orgName: 'Unit One' },
  '100002': { rank: 'Maj', firstName: 'Bo', lastName: 'Kimura', type: 'SENIOR', charter: 'PCR-XX-002', orgName: 'Unit Two' }
};

const acct = (over) => Object.assign({
  email: 'a.member@example.org',
  creationTime: '2026-07-01T00:00:00.000Z',
  lastLoginTime: NEVER
}, over || {});

const classify = (accountByCapid, ledger, opts) =>
  m.classifyWelcomeAudit_(members, accountByCapid, ledger, NOW, opts);

const seeded = (sent) => ({ seededAt: BASELINE, sent: sent || {} });

// ---------------------------------------------------------------------------
section('1. The finding this exists to produce');
{
  // Created a month after the baseline, never signed into, nothing recorded.
  const r = classify({ '100001': acct() }, seeded());
  check('one MISSED', r.missed.length, 1);
  check('names the account', r.missed[0].email, 'a.member@example.org');
  check('reason is the evidence, not a guess', r.missed[0].reason,
    'created-after-baseline-no-send-recorded');
  check('carries the unit for the IT report', r.missed[0].charter, 'PCR-XX-001');
  check('carries a display name', r.missed[0].name, 'Capt Ada Nwosu');
  check('nothing lands in unknown', r.unknown.length, 0);

  check('a recorded send clears it',
    classify({ '100001': acct() }, seeded({ '100001': { on: '2026-07-01', by: 'send' } })).missed.length,
    0);
  check('...and counts as welcomed',
    classify({ '100001': acct() }, seeded({ '100001': { on: '2026-07-01', by: 'send' } })).welcomed,
    1);
  check('a SEEDED entry clears it just as well',
    classify({ '100001': acct() }, seeded({ '100001': { on: '2026-06-01', by: 'seed' } })).missed.length,
    0);
}

// ---------------------------------------------------------------------------
section('2. Never accuse what cannot be judged');
{
  // The failure that would discredit the whole report: no baseline yet, so every
  // member has no ledger entry. None of them may be called MISSED.
  const unseeded = classify({ '100001': acct(), '100002': acct() }, { seededAt: '', sent: {} });
  check('unseeded ledger produces no accusations', unseeded.missed.length, 0);
  check('...everyone is UNKNOWN instead', unseeded.unknown.length, 2);
  check('...and says why', unseeded.unknown[0].reason, 'no-baseline');

  // Same protection when the ledger file is missing entirely (load returns this shape).
  check('a lost ledger file behaves identically',
    classify({ '100001': acct() }, { seededAt: '', sent: {} }).missed.length,
    0);

  // Predates the baseline: the seed did not vouch for it, but that only means it
  // had no login history back then — not that nobody ever welcomed them.
  const old = classify({ '100001': acct({ creationTime: '2025-03-04T00:00:00.000Z' }) }, seeded());
  check('an account older than the baseline is UNKNOWN, not MISSED', old.missed.length, 0);
  check('...with its own reason', old.unknown[0].reason, 'predates-baseline');

  // An account created moments ago is mid-provisioning, or has a resend in flight.
  const fresh = classify({ '100001': acct({ creationTime: '2026-07-25T18:00:00.000Z' }) }, seeded());
  check('an account younger than the grace period is PENDING', fresh.pending.length, 1);
  check('...not MISSED', fresh.missed.length, 0);
  check('...and says why', fresh.pending[0].reason, 'too-new-to-judge');

  check('the grace period is configurable',
    classify({ '100001': acct({ creationTime: '2026-07-20T00:00:00.000Z' }) },
      seeded(), { graceDays: 30 }).pending.length,
    1);
}

// ---------------------------------------------------------------------------
section('3. Members with no account are not this module\'s business');
{
  // Reporting the Level I-gated and the not-yet-provisioned here would bury the
  // real finding under people who never had an account to be welcomed to.
  const r = classify({}, seeded());
  check('counted, not reported', r.noAccount, 2);
  check('no MISSED', r.missed.length, 0);
  check('no UNKNOWN', r.unknown.length, 0);

  check('an account entry with no email is no account',
    classify({ '100001': acct({ email: '' }) }, seeded()).noAccount,
    2);
}

// ---------------------------------------------------------------------------
section('4. Sign-in state is reported, because it changes what IT should do');
{
  // A MISSED account that HAS been signed into got credentials some other way.
  // Still a ledger gap worth reporting, but the resend will (correctly) refuse it.
  const signedIn = classify({ '100001': acct({ lastLoginTime: '2026-07-20T00:00:00.000Z' }) }, seeded());
  check('still reported', signedIn.missed.length, 1);
  check('flagged as already signed in', signedIn.missed[0].neverSignedIn, false);

  check('the epoch is "never", not a real login',
    classify({ '100001': acct({ lastLoginTime: NEVER }) }, seeded()).missed[0].neverSignedIn,
    true);
  check('an absent lastLoginTime is also "never"',
    classify({ '100001': acct({ lastLoginTime: '' }) }, seeded()).missed[0].neverSignedIn,
    true);
}

// ---------------------------------------------------------------------------
section('5. Ordering — oldest account first, so the longest-suffering is at the top');
{
  const r = classify({
    '100001': acct({ creationTime: '2026-07-10T00:00:00.000Z', email: 'newer@example.org' }),
    '100002': acct({ creationTime: '2026-06-15T00:00:00.000Z', email: 'older@example.org' })
  }, seeded());
  check('oldest first', r.missed.map(x => x.email), ['older@example.org', 'newer@example.org']);
}

// ---------------------------------------------------------------------------
section('6. welcomeAuditAccountMap_ — judge the account in use, not a dead twin');
{
  // A member with duplicate accounts must be judged on the one they actually
  // use; scoring the abandoned twin would report someone who is perfectly fine.
  const map = m.welcomeAuditAccountMap_([
    { capid: '100001', email: 'dead.twin@example.org', lastLoginTime: NEVER },
    { capid: '100001', email: 'in.use@example.org', lastLoginTime: '2026-07-20T00:00:00.000Z' }
  ]);
  check('the signed-into account wins', map['100001'].email, 'in.use@example.org');

  check('order does not matter',
    m.welcomeAuditAccountMap_([
      { capid: '100001', email: 'in.use@example.org', lastLoginTime: '2026-07-20T00:00:00.000Z' },
      { capid: '100001', email: 'dead.twin@example.org', lastLoginTime: NEVER }
    ])['100001'].email,
    'in.use@example.org');

  check('an account with no CAPID is ignored',
    Object.keys(m.welcomeAuditAccountMap_([{ email: 'no.capid@example.org' }])).length,
    0);

  check('numeric and string CAPIDs key the same entry',
    Object.keys(m.welcomeAuditAccountMap_([
      { capid: 100001, email: 'a@example.org', lastLoginTime: NEVER }
    ])),
    ['100001']);
}

// ---------------------------------------------------------------------------
section('7. welcomeAuditTimestamp_');
{
  check('the epoch is 0', m.welcomeAuditTimestamp_(NEVER), 0);
  check('absent is 0', m.welcomeAuditTimestamp_(''), 0);
  check('unparseable is 0, not NaN', m.welcomeAuditTimestamp_('not a date'), 0);
  check('a real timestamp parses',
    m.welcomeAuditTimestamp_('2026-07-20T00:00:00.000Z'),
    Date.parse('2026-07-20T00:00:00.000Z'));
}

done();
