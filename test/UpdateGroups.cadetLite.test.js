/**
 * UpdateGroups.gs — giving accountless members an address, and only where asked.
 *
 * WHAT WENT WRONG
 * Two writers disagreed about what a unit's `.all` group contains. SquadronGroups
 * added cadet-lite members by their personal CAPWATCH address; the sheet-driven
 * path could not see them at all, because getMembers() filtered them out before
 * the desired set was built — so every one of those addresses looked like a
 * stranger and was marked for removal. On the CAWG cadet tenant that was 1,643
 * removals a night at 05:24, undone at 06:01. Members were off their unit's
 * all-hands for roughly half an hour, daily, and the wing-wide ca.all never had
 * them at all.
 *
 * THE RISK IN FIXING IT
 * Handing every accountless member an address makes them eligible for every
 * group whose criteria they match. That is 1,600-odd external addresses arriving
 * in groups nobody opted into — a change that looks small in a diff and is
 * discovered in a mailbox. So the addressed copy must be exactly that: a copy,
 * handed to one opted-in Groups row and thrown away.
 *
 * All CAPIDs and addresses here are synthetic.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'UpdateGroups.gs');
const { section, check, done } = makeChecker();

const m = loadModule(MODULE, {
  CONFIG: { EMAIL_DOMAIN: '@example.org', WING: 'CA' },
  Logger: { info: () => {}, warn: () => {}, error: () => {} },
  Utilities: { sleep: () => {} }
}, ['withCadetLiteAddresses_']);

const addr = m.withCadetLiteAddresses_;

/** An account holder: the directory already gave them an address. */
function withAccount(capid, email) {
  return { capsn: capid, orgid: '900', type: 'CADET', email: email };
}

/** Accountless: createMemberObject leaves email null and nothing filled it. */
function accountless(capid) {
  return { capsn: capid, orgid: '900', type: 'CADET', email: null };
}

// ---------------------------------------------------------------------------
section('1. The accountless get their CAPWATCH address');
{
  const members = { '100001': accountless('100001') };
  const out = addr(members, { '100001': 'lite.one@example.org' });

  check('addressed', out['100001'].email, 'lite.one@example.org');
  check('everything else about them is intact', out['100001'].orgid, '900');
  check('and their type still drives group matching', out['100001'].type, 'CADET');
}

// ---------------------------------------------------------------------------
section('2. Account holders are never re-addressed');
{
  const members = { '100002': withAccount('100002', 'real@example.org') };
  // A CAPWATCH personal address exists for them too — it must lose.
  const out = addr(members, { '100002': 'personal@example.org' });

  check('the Workspace address wins', out['100002'].email, 'real@example.org');
  check('and the object is passed through untouched',
    out['100002'] === members['100002'], true);
}

// ---------------------------------------------------------------------------
section('3. The original map is never mutated');
{
  // This is the assertion that matters most. A mutation here leaks these
  // addresses into every Groups row processed afterwards.
  const members = {
    '100003': accountless('100003'),
    '100004': withAccount('100004', 'real@example.org')
  };
  const before = members['100003'].email;
  const out = addr(members, { '100003': 'lite@example.org' });

  check('the source member is still addressless', members['100003'].email, before);
  check('and is still null specifically', members['100003'].email, null);
  check('while the copy is addressed', out['100003'].email, 'lite@example.org');
  check('the copy is a different object', out['100003'] === members['100003'], false);
}

// ---------------------------------------------------------------------------
section('4. No address means no group, which is the honest outcome');
{
  const members = { '100005': accountless('100005') };
  const out = addr(members, {});          // nothing on file for them

  check('still addressless', out['100005'].email, null);
  check('so every group\'s `isMatch && .email` test skips them',
    !!out['100005'].email, false);
}

// ---------------------------------------------------------------------------
section('5. Everyone survives the copy');
{
  const members = {
    '100006': accountless('100006'),
    '100007': withAccount('100007', 'a@example.org'),
    '100008': accountless('100008')
  };
  const out = addr(members, { '100006': 'six@example.org' });

  check('same population', Object.keys(out).sort(), ['100006', '100007', '100008']);
  check('one newly addressed', out['100006'].email, 'six@example.org');
  check('one already addressed', out['100007'].email, 'a@example.org');
  check('one still unreachable', out['100008'].email, null);
}

// ---------------------------------------------------------------------------
section('6. Malformed input does not become a wrong address');
{
  check('an empty member map yields an empty map',
    Object.keys(addr({}, { '1': 'x@example.org' })), []);

  const nullMember = { '100009': null };
  check('a null member is passed through, not crashed on',
    addr(nullMember, {})['100009'], null);

  const blankFallback = { '100010': accountless('100010') };
  check('a blank CAPWATCH address is not an address',
    addr(blankFallback, { '100010': '' })['100010'].email, null);
}

done();
