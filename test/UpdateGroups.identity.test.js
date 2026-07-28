/**
 * UpdateGroups.gs — reconciling a group against what it already holds.
 *
 * WHAT WENT WRONG
 *
 * Google ignores dots and +tags in a gmail.com local part. cadet@gmail.com,
 * ca.det@gmail.com and cadet+cap@gmail.com are ONE mailbox wearing three
 * spellings, and CAPWATCH supplies whichever the member happened to type.
 *
 * The delta compared raw strings. So a member whose stored spelling differed
 * from their CAPWATCH spelling looked missing and unwanted at the same time:
 * the desired spelling was inserted — 409 Member already exists — and the
 * stored spelling was removed. Every night. In every group they belong to.
 * One CAWG cadet run: 39 errors, 33 removals, all of them this.
 *
 * SquadronGroups.gs fixed it in diffGroupMembership_() and the sheet path did
 * not, so the two writers disagreed about who was already a member.
 *
 * THE LESSON THESE ASSERTIONS ENCODE
 * Identity is what an address REACHES, not how it is spelled. Any membership
 * decision made by comparing address strings is wrong the moment two spellings
 * of one account meet — and they meet constantly, because people type their own
 * addresses.
 *
 * Delta values are the contract with the apply loop: 1 add, -1 remove, 0 leave.
 *
 * All addresses here are synthetic.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'UpdateGroups.gs');
const UTILS = path.join(__dirname, '..', 'src', 'utils.gs');
const { section, check, done } = makeChecker();

// The REAL googleAccountKey, not a fake. What is under test is the agreement
// between these two files, so stubbing the key would test nothing: UpdateGroups.gs
// reaches for it as a global exactly the way Apps Script's shared namespace
// supplies it.
// utils.gs declares its own Logger, so that global is NOT injected here.
const utils = loadModule(UTILS, {
  CONFIG: {},
  Utilities: { sleep: () => {} }
}, ['googleAccountKey']);

const m = loadModule(MODULE, {
  CONFIG: { EMAIL_DOMAIN: '@example.org', WING: 'CA' },
  Logger: { info: () => {}, warn: () => {}, error: () => {} },
  Utilities: { sleep: () => {} },
  googleAccountKey: utils.googleAccountKey
}, ['reconcileGroupDelta_']);

const reconcile = m.reconcileGroupDelta_;

/** Desired set as getGroupMembers leaves it: every address marked for adding. */
function desired(...emails) {
  const d = {};
  emails.forEach(e => { d[e] = 1; });
  return d;
}

/** What the group holds now, as getCurrentGroup returns it. */
function holds(...entries) {
  return entries.map(e => (typeof e === 'string' ? { email: e, role: 'MEMBER' } : e));
}

// ---------------------------------------------------------------------------
section('1. The churn itself: one account, two spellings');
{
  // The exact shape observed on ca.all night after night.
  const d = desired('caleb.l.anderson@gmail.com');
  const t = reconcile(d, holds('caleblanderson@gmail.com'));

  check('the desired spelling is NOT inserted', d['caleb.l.anderson@gmail.com'], 0);
  check('the stored spelling is NOT removed', d['caleblanderson@gmail.com'], undefined);
  check('and the run says it happened', t.sameAccountDifferentSpelling, 1);
}

// ---------------------------------------------------------------------------
section('2. Before the fix this was an add AND a remove');
{
  // Documents the old behaviour so a regression is unmistakable: string equality
  // would leave the desired entry at 1 (insert -> 409) and mark the held address
  // -1 (remove). Both assertions below fail the moment identity matching is lost.
  const d = desired('j.doe@gmail.com');
  reconcile(d, holds('jdoe@gmail.com'));

  const wouldInsert = Object.keys(d).filter(k => d[k] === 1);
  const wouldRemove = Object.keys(d).filter(k => d[k] === -1);
  check('nothing to insert', wouldInsert, []);
  check('nothing to remove', wouldRemove, []);
}

// ---------------------------------------------------------------------------
section('3. +tags fold too, and dots inside a tag are part of the tag');
{
  const d = desired('cadet+cap@gmail.com');
  reconcile(d, holds('cadet@gmail.com'));
  check('tagged and untagged are one account', d['cadet+cap@gmail.com'], 0);

  const d2 = desired('cadet+c.a.p@gmail.com');
  reconcile(d2, holds('cadet@gmail.com'));
  check('dots in the tag do not change the mailbox', d2['cadet+c.a.p@gmail.com'], 0);
}

// ---------------------------------------------------------------------------
section('4. Only gmail folds — everywhere else dots are significant');
{
  const d = desired('j.doe@example.com');
  reconcile(d, holds('jdoe@example.com'));
  check('the wanted address is still added', d['j.doe@example.com'], 1);
  check('the unwanted one is still removed', d['jdoe@example.com'], -1);
}

// ---------------------------------------------------------------------------
section('5. Real adds and real removals still happen');
{
  const d = desired('stays@example.org', 'joins@example.org');
  const t = reconcile(d, holds('stays@example.org', 'leaves@example.org'));

  check('an existing wanted member is left alone', d['stays@example.org'], 0);
  check('a new member is added', d['joins@example.org'], 1);
  check('an unwanted member is removed', d['leaves@example.org'], -1);
  check('none of that is spelling churn', t.sameAccountDifferentSpelling, 0);
}

// ---------------------------------------------------------------------------
section('6. Managers and owners are never auto-removed');
{
  const d = desired('member@example.org');
  reconcile(d, [
    { email: 'member@example.org', role: 'MEMBER' },
    { email: 'chair@example.org', role: 'OWNER' },
    { email: 'staff@example.org', role: 'MANAGER' }
  ]);

  check('an owner nobody asked for stays', d['chair@example.org'], undefined);
  check('a manager nobody asked for stays', d['staff@example.org'], undefined);
}

// ---------------------------------------------------------------------------
section('7. Cleaning up the damage the old comparison left behind');
{
  // Both spellings ended up in the group: one was inserted before the removal of
  // the other failed, or a second writer added the other. Keep one, drop the
  // extra — otherwise the duplicate sits there forever receiving two copies.
  const d = desired('a.b@gmail.com');
  const t = reconcile(d, holds('ab@gmail.com', 'a.b@gmail.com'));

  check('the account is not re-inserted', d['a.b@gmail.com'], 0);
  check('the second spelling is removed', d['ab@gmail.com'], -1);
  check('and counted as redundant', t.redundantCurrentAddresses, 1);
  check('not as ordinary churn', t.sameAccountDifferentSpelling, 0);

  // Google's listing order must not decide which address survives. It is the
  // spelling the sheet asked for, either way round — and getting this wrong
  // meant the removal of the twin overwrote the keep, deleting the member.
  const d2 = desired('a.b@gmail.com');
  reconcile(d2, holds('a.b@gmail.com', 'ab@gmail.com'));
  check('same outcome when the group lists them the other way',
    [d2['a.b@gmail.com'], d2['ab@gmail.com']], [0, -1]);
}

// ---------------------------------------------------------------------------
section('8. Two desired addresses for one account');
{
  // Two members, or one member listed twice, resolving to the same mailbox.
  // Inserting the second can only ever return 409, so stand it down quietly.
  const d = desired('x.y@gmail.com', 'xy@gmail.com');
  const t = reconcile(d, holds());

  const inserts = Object.keys(d).filter(k => d[k] === 1);
  check('exactly one insert is attempted', inserts.length, 1);
  check('the first spelling wins', inserts[0], 'x.y@gmail.com');
  check('the duplicate is stood down, not removed', d['xy@gmail.com'], 0);
  check('and reported', t.duplicateDesiredAddresses, 1);
}

// ---------------------------------------------------------------------------
section('9. Reading the current membership defensively');
{
  const d = desired('a@example.org');
  check('no current members at all', reconcile(d, []).sameAccountDifferentSpelling, 0);
  check('the wanted member is still added', d['a@example.org'], 1);

  const d2 = desired('a@example.org');
  reconcile(d2, null);
  check('a null member list does not throw', d2['a@example.org'], 1);

  const d3 = desired('a@example.org');
  reconcile(d3, [{ email: '', role: 'MEMBER' }, { email: null }, {}]);
  check('blank entries are ignored', Object.keys(d3).length, 1);

  const d4 = desired('a@example.org');
  reconcile(d4, holds('  A@Example.org  '));
  check('case and padding do not defeat the match', d4['a@example.org'], 0);

  check('an empty desired set does not throw',
    reconcile({}, holds('someone@example.org')).sameAccountDifferentSpelling, 0);
  check('a null desired set does not throw',
    reconcile(null, holds('someone@example.org')).duplicateDesiredAddresses, 0);
}

// ---------------------------------------------------------------------------
section('10. An all-dots local part is never folded');
{
  // Stripping the dots leaves nothing. googleAccountKey() refuses to fold these
  // precisely so unrelated members cannot collapse onto one key.
  const d = desired('...@gmail.com');
  reconcile(d, holds('..@gmail.com'));
  check('two nonsense addresses stay distinct', d['...@gmail.com'], 1);
  check('and the held one is still removed', d['..@gmail.com'], -1);
}

done();
