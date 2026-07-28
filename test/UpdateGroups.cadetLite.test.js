/**
 * UpdateGroups.gs — which members a Groups-sheet row is allowed to see.
 *
 * WHAT WENT WRONG, TWICE
 *
 * First: the sheet path could not see cadet-lite members at all, because
 * getMembers() filtered them out before the desired set was built. SquadronGroups
 * added them to every unit `.all` by their CAPWATCH address; this path saw
 * strangers and removed them. 1,643 removals a night on the CAWG cadet tenant,
 * undone an hour later.
 *
 * Then, fixing it: the opt-in gate was written as "a cadet-lite member has no
 * address until we supply one", so rows that had not opted in would skip them
 * naturally. That belief was wrong — addContactInfo() fills .email from the
 * CAPWATCH PRIMARY contact for every member, cadet-lite included. The gate never
 * engaged. Cadet-lite members were eligible for EVERY group whose criteria they
 * matched, not the two rows that asked.
 *
 * THE LESSON THESE ASSERTIONS ENCODE
 * Eligibility is decided by WHO A MEMBER IS, never by whether some field happens
 * to be populated. A membership rule resting on a field that another module fills
 * for its own reasons is a rule that can be switched off by a change nowhere near
 * it.
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
}, ['getGroupMembers', 'externalMemberAllowedForGroup_',
    'groupAttributeByName', 'groupAllowExternalByName']);

/**
 * A member as the pipeline actually presents one: addContactInfo has already
 * supplied an address, whether or not they hold a Workspace account.
 */
function member(capid, orgid, opts) {
  const o = opts || {};
  return {
    capsn: capid,
    orgid: orgid,
    group: o.group || '',
    type: o.type || 'CADET',
    // `||` would override an explicit null, which is the one case section 6 is
    // about. Honour the key when it is present, defaulted or not.
    email: Object.prototype.hasOwnProperty.call(o, 'email') ? o.email : (capid + '@example.org')
  };
}

const squadrons = {
  '900': { wing: 'CA', unit: '101', scope: 'UNIT', charter: 'PCR-CA-101' }
};

/** Runs one Groups-sheet row against a member set, as getEmailGroupDeltas does. */
function generate(members) {
  return m.getGroupMembers('all', 'type', 'CADET', members, squadrons);
}

// ---------------------------------------------------------------------------
section('1. A cadet-lite member is addressable, which is why identity must gate');
{
  // The exact condition the broken gate tested for. If this is ever false again,
  // the "no address means no group" reasoning is back, and it is still wrong.
  const lite = member('100001', '900');
  check('they carry a CAPWATCH address like everyone else', !!lite.email, true);
  check('so an address-based gate would let them through everywhere',
    Object.keys(generate({ '100001': lite })['ca.all']).length, 1);
}

// ---------------------------------------------------------------------------
section('2. A row that did not opt in never sees them');
{
  // membersCore: the caller removed cadet-lite by CAPID before calling.
  const core = { '100002': member('100002', '900') };
  const out = generate(core);
  check('only the core member is present',
    Object.keys(out['ca.all']), ['100002@example.org']);
}

// ---------------------------------------------------------------------------
section('3. A row that opted in sees both');
{
  const all = {
    '100003': member('100003', '900'),
    '100004': member('100004', '900')     // cadet-lite, kept in the map
  };
  const out = generate(all);
  check('both reach the wing group',
    Object.keys(out['ca.all']).sort(),
    ['100003@example.org', '100004@example.org']);
  check('and the unit group', Object.keys(out['ca101.all']).sort(),
    ['100003@example.org', '100004@example.org']);
}

// ---------------------------------------------------------------------------
section('4. Membership follows the member set, not the address');
{
  // Same person, same address, present in one call and absent from the other.
  // The ONLY difference is whether the caller included them — which is the
  // property the fix depends on.
  const lite = member('100005', '900');
  const withThem = generate({ '100005': lite });
  const withoutThem = generate({});

  check('included → in the group', Object.keys(withThem['ca.all']).length, 1);
  check('excluded → not in the group', Object.keys(withoutThem['ca.all']).length, 0);
  check('their address was identical in both cases', !!lite.email, true);
}

// ---------------------------------------------------------------------------
section('5. Criteria still apply to whoever is in the set');
{
  const mixed = {
    '100006': member('100006', '900', { type: 'CADET' }),
    '100007': member('100007', '900', { type: 'SENIOR' })
  };
  const out = generate(mixed);   // row asks for CADET only
  check('a senior does not join a cadet row',
    Object.keys(out['ca.all']), ['100006@example.org']);
}

// ---------------------------------------------------------------------------
section('6. A member with no address still reaches nothing');
{
  const out = generate({ '100008': member('100008', '900', { email: null }) });
  check('nowhere to send, so no group', Object.keys(out['ca.all']).length, 0);
}

// ---------------------------------------------------------------------------
section('7. Whether a group may take an off-domain address');
{
  // The apply loop allowed external members for exactly one Attribute value and
  // skipped every other case with a bare `continue` — no log, no counter. So a
  // group could want 1,644 external members, the delta could say so, and the run
  // could report success having added none. That is what happened to ca.all.
  const attr = m.groupAttributeByName;
  const ext = m.groupAllowExternalByName;
  const allowed = m.externalMemberAllowedForGroup_;

  Object.keys(attr).forEach(k => delete attr[k]);
  Object.keys(ext).forEach(k => delete ext[k]);

  attr['contact-row'] = 'contact';
  attr['all'] = 'type';
  ext['all'] = true;              // set by Add EXT or implied by Add Lite
  attr['achievements'] = 'achievements';
  ext['achievements'] = false;

  check('the legacy contact row still works', allowed('contact-row'), true);
  check('a row the sheet marks external is allowed', allowed('all'), true);
  check('a row that asked for neither is not', allowed('achievements'), false);
  check('an unknown row is not', allowed('no-such-row'), false);
  check('whitespace in the name does not defeat the lookup', allowed('  all  '), true);
  check('a blank name is not allowed', allowed(''), false);
  check('null is not allowed', allowed(null), false);
}

done();
