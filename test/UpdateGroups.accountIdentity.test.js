/**
 * UpdateGroups.gs — membership compared on Google ACCOUNT identity, not on the
 * string CAPWATCH happens to hold.
 *
 * Two spellings can be one mailbox: on gmail.com dots carry no meaning and
 * everything from a '+' onward is a tag. When the group holds one spelling and
 * CAPWATCH supplies another, a string diff reads "one member to add, one stranger
 * to remove" — the add returns 409 and is swallowed, the remove succeeds, and the
 * member is off the list until the next run. Silently, every time their address
 * changes shape.
 *
 * SquadronGroups has compared this way since it started nesting external groups.
 * This pass did not, which was inert while these lists held only @<tenant>
 * addresses — every one of those is its own key. It stopped being inert when Add
 * EXT on the wing "all" row admitted CAPWATCH personal addresses, most of them
 * gmail.
 *
 * googleAccountKey() is loaded from utils.gs rather than reimplemented here, so
 * these assertions break if the real folding rules ever change.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeChecker } = require('./helpers/apps-script');

const SRC = path.join(__dirname, '..', 'src');
const { section, check, done } = makeChecker();

// utils.gs declares its own Logger, so nothing may be injected under that name.
const { googleAccountKey } = loadModule(path.join(SRC, 'utils.gs'), {}, ['googleAccountKey']);

const m = loadModule(
  path.join(SRC, 'accounts-and-groups', 'UpdateGroups.gs'),
  {
    googleAccountKey: googleAccountKey,
    CONFIG: { EMAIL_DOMAIN: '@example.org', WING: 'CA' },
    Logger: { info: () => {}, warn: () => {}, error: () => {} },
    Utilities: { sleep: () => {} }
  },
  ['reconcileCurrentAgainstDesired_']
);

/** Members as getCurrentGroup returns them. */
const held = (...emails) => emails.map(e => ({ email: e, role: 'MEMBER' }));

// ---------------------------------------------------------------------------
section('1. The ordinary cases still decide the same way');
{
  const desired = { 'a@example.org': 1, 'b@example.org': 1 };
  const out = m.reconcileCurrentAgainstDesired_(desired, held('a@example.org', 'z@example.org'));

  check('an address held and wanted is keep', out['a@example.org'], 0);
  check('an address wanted and not held is add', out['b@example.org'], 1);
  check('an address held and not wanted is remove', out['z@example.org'], -1);
}

// ---------------------------------------------------------------------------
section('2. One mailbox, two spellings');
{
  // The exact shape of the bug: the group holds the dotted spelling, CAPWATCH
  // now supplies the undotted one.
  const desired = { 'firstlast@gmail.com': 1 };
  const out = m.reconcileCurrentAgainstDesired_(desired, held('first.last@gmail.com'));

  check('the desired spelling is not added again', out['firstlast@gmail.com'], 0);
  check('the held spelling is not removed', out['first.last@gmail.com'], undefined);
  check('nothing else appeared', Object.keys(out).length, 1);
}
{
  // And the other direction, plus a tag.
  const desired = { 'first.last@gmail.com': 1 };
  const out = m.reconcileCurrentAgainstDesired_(desired, held('firstlast+cap@gmail.com'));
  check('a +tag is the same account', out['first.last@gmail.com'], 0);
  check('the tagged spelling is left alone', out['firstlast+cap@gmail.com'], undefined);
}
{
  const desired = { 'firstlast@gmail.com': 1 };
  const out = m.reconcileCurrentAgainstDesired_(desired, held('first.last@googlemail.com'));
  check('googlemail is the same service', out['firstlast@gmail.com'], 0);
}

// ---------------------------------------------------------------------------
section('3. Folding stops at gmail');
{
  // Dots ARE significant elsewhere, so folding them would merge two real people.
  const desired = { 'a.b@example.org': 1 };
  const out = m.reconcileCurrentAgainstDesired_(desired, held('ab@example.org'));

  check('a non-gmail dotted address is still wanted', out['a.b@example.org'], 1);
  check('and the undotted one is a stranger', out['ab@example.org'], -1);
}

// ---------------------------------------------------------------------------
section('4. The wing-account swap still fires');
{
  // Why this matters here: a Level I senior with no account is carried by their
  // CAPWATCH personal address. The day they are provisioned, the desired address
  // becomes their wing one — a different domain, so a different key, so the
  // personal address must still fall out.
  const desired = { 'jane.doe@example.org': 1 };
  const out = m.reconcileCurrentAgainstDesired_(desired, held('jdoe1987@gmail.com'));

  check('the new wing address is added', out['jane.doe@example.org'], 1);
  check('the personal address is removed', out['jdoe1987@gmail.com'], -1);
}

// ---------------------------------------------------------------------------
section('5. Roles other than MEMBER are still left alone');
{
  const desired = {};
  const out = m.reconcileCurrentAgainstDesired_(desired, [
    { email: 'owner@example.org', role: 'OWNER' },
    { email: 'manager@example.org', role: 'MANAGER' },
    { email: 'member@example.org', role: 'MEMBER' }
  ]);

  check('an OWNER is not removed', out['owner@example.org'], undefined);
  check('a MANAGER is not removed', out['manager@example.org'], undefined);
  check('a plain MEMBER is', out['member@example.org'], -1);
}
{
  // A manager who IS wanted still resolves to keep, by account identity.
  const desired = { 'firstlast@gmail.com': 1 };
  const out = m.reconcileCurrentAgainstDesired_(desired, [
    { email: 'first.last@gmail.com', role: 'MANAGER' }
  ]);
  check('a wanted MANAGER is marked keep, not re-added', out['firstlast@gmail.com'], 0);
}

// ---------------------------------------------------------------------------
section('6. Input the live pass actually hands it');
{
  // getCurrentGroup lowercases, but a role can be missing and a group can be empty.
  const out = m.reconcileCurrentAgainstDesired_({ 'a@example.org': 1 }, [{ email: 'a@example.org' }]);
  check('a missing role defaults to MEMBER and matches', out['a@example.org'], 0);

  check('no current members leaves desired untouched',
    m.reconcileCurrentAgainstDesired_({ 'a@example.org': 1 }, [])['a@example.org'], 1);
  check('an undefined member list is not an error',
    m.reconcileCurrentAgainstDesired_({ 'a@example.org': 1 })['a@example.org'], 1);
  check('an undefined desired map is not an error',
    Object.keys(m.reconcileCurrentAgainstDesired_(undefined, held('a@example.org'))).length, 1);

  const blanks = m.reconcileCurrentAgainstDesired_({ 'a@example.org': 1 },
    [{ email: '', role: 'MEMBER' }, { email: '   ', role: 'MEMBER' }]);
  check('blank addresses are ignored, not removed', Object.keys(blanks).length, 1);

  // Google returns what it stores; CAPWATCH is sanitized to lowercase. Guard the
  // seam anyway — a case-only difference must not read as add-plus-remove.
  const cased = m.reconcileCurrentAgainstDesired_({ 'a@example.org': 1 }, held('A@Example.org'));
  check('case alone is not a difference', cased['a@example.org'], 0);
  check('and does not add a second entry', Object.keys(cased).length, 1);
}

done();
