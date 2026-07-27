/**
 * Membership diffing on Google ACCOUNT identity rather than string equality.
 *
 * WHAT WENT WRONG
 * A group held one spelling of a Gmail address while CAPWATCH supplied another —
 * dots carry no meaning on gmail.com, so both are one mailbox. The string diff saw
 * a member to ADD and a stranger to REMOVE. The add came back 409 "Member already
 * exists" and was swallowed; the remove succeeded. The member was dropped from
 * their unit list and only restored on the following run: a day off the list, once
 * per address change, with nothing in the log naming who it happened to.
 *
 * The second shape of the same fault: when BOTH spellings are in the desired set,
 * one is added and the other 409s on every run forever — permanent noise, and an
 * ERROR line for something the caller treats as success.
 *
 * These assertions are about not merging people who merely look alike, as much as
 * about matching the ones who are the same. Folding dots on a domain where dots
 * are significant would silently delete real members, which is worse than the bug.
 *
 * All addresses here are synthetic.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const { section, check, done } = makeChecker();
const { logger } = makeLogger();

// utils.gs declares its own Logger, so that global is NOT injected here.
const utils = loadModule(path.join(__dirname, '..', 'src', 'utils.gs'), {
  CONFIG: {},
  Utilities: { sleep: () => {} }
}, ['googleAccountKey']);

// SquadronGroups.gs reaches for googleAccountKey as a global, the way Apps Script
// resolves across files.
const groups = loadModule(path.join(__dirname, '..', 'src', 'squadron-groups', 'SquadronGroups.gs'), {
  Logger: logger,
  CONFIG: { WING: 'CA', EMAIL_DOMAIN: '@example.org' },
  googleAccountKey: utils.googleAccountKey,
  parseFile: () => []
}, ['diffGroupMembership_']);

const key = utils.googleAccountKey;
const diff = groups.diffGroupMembership_;

/** desiredMembers shape the real callers build. */
function desired(...emails) {
  const out = {};
  emails.forEach(e => { out[e] = { email: e, role: 'MEMBER' }; });
  return out;
}

const added = plan => plan.toAdd.map(a => a.email);

// ---------------------------------------------------------------------------
section('1. One Gmail account, many spellings');
{
  const canonical = key('firstlast@gmail.com');
  check('dots are not significant', key('first.last@gmail.com'), canonical);
  check('several dots', key('f.i.r.s.t.last@gmail.com'), canonical);
  check('a +tag is not part of the mailbox', key('firstlast+cap@gmail.com'), canonical);
  check('tag and dots together', key('first.last+cawg.2026@gmail.com'), canonical);
  check('googlemail is gmail', key('first.last@googlemail.com'), canonical);
  check('case', key('First.Last@Gmail.COM'), canonical);
  check('surrounding whitespace', key('  first.last@gmail.com  '), canonical);
}

// ---------------------------------------------------------------------------
section('2. Everywhere else, a dot is a real character');
{
  check('two people at one domain stay two people',
    key('a.member@example.org') === key('amember@example.org'), false);
  check('a Workspace tenant address is untouched',
    key('First.Last@cawgcap.org'), 'first.last@cawgcap.org');
  check('a +tag off gmail is left alone (it may be the real mailbox)',
    key('sales+web@example.org'), 'sales+web@example.org');
  check('a lookalike domain is not gmail',
    key('a.b@notgmail.com'), 'a.b@notgmail.com');
  check('a subdomain of gmail is not gmail',
    key('a.b@mail.gmail.com'), 'a.b@mail.gmail.com');
}

// ---------------------------------------------------------------------------
section('3. Input that is not an address is returned unchanged');
{
  check('empty', key(''), '');
  check('null', key(null), '');
  check('undefined', key(undefined), '');
  check('no @', key('not-an-address'), 'not-an-address');
  check('nothing before the @', key('@gmail.com'), '@gmail.com');
  check('nothing after the @', key('someone@'), 'someone@');
  check('an all-dots local part is not folded away',
    key('...@gmail.com'), '...@gmail.com');
  check('a bare +tag local part is not folded away',
    key('+tag@gmail.com'), '+tag@gmail.com');
}

// ---------------------------------------------------------------------------
section('4. The reported bug: group holds one spelling, CAPWATCH sends another');
{
  const plan = diff(['dotless.name@gmail.com'], desired('dotlessname@gmail.com'));
  check('nothing to add', added(plan), []);
  check('AND NOTHING TO REMOVE — the member stays', plan.toRemove, []);
}

// ---------------------------------------------------------------------------
section('5. The other shape: both spellings desired');
{
  const plan = diff(['dotlessname@gmail.com'],
    desired('dotlessname@gmail.com', 'dotless.name@gmail.com'));
  check('no second insert to be refused', added(plan), []);
  check('nothing removed', plan.toRemove, []);
  check('the duplicate is reported, not hidden', plan.duplicates.length, 1);
  check('and names both spellings',
    [plan.duplicates[0].kept, plan.duplicates[0].skipped].sort(),
    ['dotless.name@gmail.com', 'dotlessname@gmail.com']);
}

// ---------------------------------------------------------------------------
section('6. Ordinary adds and removes still happen');
{
  const plan = diff(
    ['stays@example.org', 'leaves@example.org'],
    desired('stays@example.org', 'joins@example.org')
  );
  check('the newcomer is added', added(plan), ['joins@example.org']);
  check('the departed is removed', plan.toRemove, ['leaves@example.org']);
  check('role carries through', plan.toAdd[0].role, 'MEMBER');

  const owner = diff([], { 'boss@example.org': { role: 'OWNER' } });
  check('a non-default role is preserved', owner.toAdd[0].role, 'OWNER');
  check('a member with no role object defaults to MEMBER',
    diff([], { 'x@example.org': {} }).toAdd[0].role, 'MEMBER');
}

// ---------------------------------------------------------------------------
section('7. Two different people are never merged');
{
  const plan = diff(
    ['first.person@example.org'],
    desired('first.person@example.org', 'firstperson@example.org')
  );
  check('both are kept apart on a non-gmail domain',
    added(plan), ['firstperson@example.org']);
  check('neither is removed', plan.toRemove, []);
  check('and neither is called a duplicate', plan.duplicates, []);
}

// ---------------------------------------------------------------------------
section('8. Messy input does not become a wrong decision');
{
  check('case differences alone are not a change',
    diff(['Person@Example.org'], desired('person@example.org')).toRemove, []);
  check('whitespace in the desired key is trimmed',
    added(diff(['person@example.org'], desired('  person@example.org  '))), []);
  check('an empty current list adds everyone',
    added(diff([], desired('a@example.org', 'b@example.org'))).length, 2);
  check('an empty desired set removes everyone',
    diff(['a@example.org'], {}).toRemove, ['a@example.org']);
  check('null current list is survivable', diff(null, desired('a@example.org')).toRemove, []);
  check('null desired set is survivable', diff(['a@example.org'], null).toRemove, ['a@example.org']);
  check('blank entries in the current list are ignored',
    diff(['', null, 'a@example.org'], desired('a@example.org')).toRemove, []);
}

done();
