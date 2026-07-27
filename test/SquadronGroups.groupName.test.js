/**
 * SquadronGroups.gs — group names that Google will actually accept.
 *
 * WHAT WENT WRONG
 * groups.insert refuses a name over 73 characters outright — 400 "Invalid Input:
 * groupName" — and creates nothing. On the CAWG senior tenant two lists for one
 * squadron had never been created for exactly this reason, and the failure was
 * invisible because it only fires on CREATE: an over-long name on a group that
 * already exists is never re-sent, so 66 other units looked fine.
 *
 * The irony is that the offending name came from the ABBREVIATION path. A unit
 * called "... Composite Squadron" matches none of the shortening patterns, so it
 * falls through to "Sqdn <n> <full name>" — which says squadron twice and is
 * longer than what it replaced.
 *
 * WHAT THESE ASSERTIONS PROTECT
 * The label must survive intact. It is what tells one of a unit's lists from
 * another, and two groups whose names differ only past the cut are the same name
 * as far as anyone reading the console is concerned. So the unit gives way.
 *
 * All unit names here are synthetic.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'squadron-groups', 'SquadronGroups.gs');
const { section, check, done } = makeChecker();
const { logger } = makeLogger();

const m = loadModule(MODULE, {
  Logger: logger,
  CONFIG: { WING: 'CA', EMAIL_DOMAIN: '@example.org', WING_ABBREVIATION: 'CAWG' },
  parseFile: () => []
}, ['fitGroupName_', 'GROUP_NAME_MAX_LEN_']);

const fit = m.fitGroupName_;
const MAX = m.GROUP_NAME_MAX_LEN_;

const LONG_LABEL = 'Deputy Commander for Cadets';
// 47 characters, the shape that actually failed: an abbreviation that grew.
const LONG_UNIT = 'Sqdn 404 Placeholder Aerospace Composite Squadron';

// ---------------------------------------------------------------------------
section('1. The limit is the limit');
{
  check('73, the Directory API maximum', MAX, 73);
  check('the name that failed in production now fits',
    fit(LONG_UNIT, LONG_LABEL).length <= MAX, true);
  check('so does the longest sibling label',
    fit(LONG_UNIT, 'Deputy Commander for Seniors').length <= MAX, true);
}

// ---------------------------------------------------------------------------
section('2. Names that already fit are untouched');
{
  check('a short pair is joined unchanged',
    fit('Sqdn 41 Example', 'All'), 'Sqdn 41 Example - All');
  check('a name landing exactly on the limit is not trimmed', (() => {
    const label = 'All';
    const unit = 'U'.repeat(MAX - label.length - 3);
    const out = fit(unit, label);
    return out.length === MAX && out === `${unit} - ${label}`;
  })(), true);
  check('one character over IS trimmed', (() => {
    const label = 'All';
    const unit = 'U'.repeat(MAX - label.length - 2);
    return fit(unit, label).length < `${unit} - ${label}`.length;
  })(), true);
}

// ---------------------------------------------------------------------------
section('3. The label always survives whole');
{
  const out = fit(LONG_UNIT, LONG_LABEL);
  check('the label is present in full', out.indexOf(LONG_LABEL) > -1, true);
  check('and it is at the end', out.slice(-LONG_LABEL.length), LONG_LABEL);

  // Two lists of one unit must not collapse into the same name.
  check('sibling lists stay distinguishable',
    fit(LONG_UNIT, 'Deputy Commander for Cadets') !== fit(LONG_UNIT, 'Deputy Commander for Seniors'),
    true);
}

// ---------------------------------------------------------------------------
section('4. The unit is cut at a word boundary, not mid-word');
{
  const out = fit(LONG_UNIT, LONG_LABEL);
  const unitPart = out.slice(0, out.length - LONG_LABEL.length - 3);
  check('no dangling separator', /[\s,\-]$/.test(unitPart), false);
  check('every surviving word is a whole word from the original',
    unitPart.split(' ').every(w => LONG_UNIT.split(' ').indexOf(w) > -1), true);
  check('what survives is the START of the unit name',
    LONG_UNIT.indexOf(unitPart), 0);
}

// ---------------------------------------------------------------------------
section('5. A label with no room left still yields a usable name');
{
  const hugeLabel = 'L'.repeat(MAX + 10);
  check('an over-long label alone is cut to the limit',
    fit('Sqdn 1 Example', hugeLabel).length, MAX);

  const exactLabel = 'L'.repeat(MAX);
  check('a label filling the limit leaves no unit',
    fit('Sqdn 1 Example', exactLabel), exactLabel);

  // A unit whose first word alone cannot fit must not produce " - Label".
  check('no orphan separator when nothing of the unit fits',
    fit('Wwwwwwwwwwwwwwwwwwwwwwwwwwwwww', 'L'.repeat(MAX - 4)).indexOf(' - '), -1);
}

// ---------------------------------------------------------------------------
section('6. Missing pieces');
{
  check('no label', fit('Sqdn 1 Example', ''), 'Sqdn 1 Example');
  check('no unit', fit('', 'All'), 'All');
  check('neither', fit('', ''), '');
  check('null unit', fit(null, 'All'), 'All');
  check('null label', fit('Sqdn 1 Example', null), 'Sqdn 1 Example');
  check('both null', fit(null, null), '');
  check('whitespace is trimmed before joining',
    fit('  Sqdn 1 Example  ', '  All  '), 'Sqdn 1 Example - All');
}

// ---------------------------------------------------------------------------
section('7. Nothing this function returns can be refused for length');
{
  const units = ['', 'S', 'Sqdn 1 Example', LONG_UNIT, 'X'.repeat(300)];
  const labels = ['', 'All', LONG_LABEL, 'Parents & Guardians', 'Y'.repeat(200)];
  const all = [];
  units.forEach(u => labels.forEach(l => all.push(fit(u, l))));
  check('every combination is within the limit',
    all.every(n => n.length <= MAX), true);
  check('and none is undefined or null',
    all.every(n => typeof n === 'string'), true);
}

done();
