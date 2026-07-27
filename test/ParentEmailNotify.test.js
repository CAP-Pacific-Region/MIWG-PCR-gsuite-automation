/**
 * ParentEmailNotify.gs — who gets told about an unusable parent address, and
 * how often.
 *
 * The two decisions this module makes that can hurt someone:
 *
 *   1. SUPPRESSION. Too loose and a unit gets the same list monthly until they
 *      stop reading it. Too tight and a genuinely broken address goes unreported
 *      while a family silently receives nothing. The window is per member AND
 *      per address — a cadet with two bad addresses who has one corrected must
 *      still be told about the other, which keying on the member alone would
 *      hide for three months.
 *
 *   2. RESOLUTION. The ledger records an ADDRESS; the person is rebuilt from the
 *      current extract. A cadet who left, or whose record was corrected, must
 *      resolve to nothing rather than to a stale row — the alternative is
 *      mailing a commander about a cadet who is not theirs.
 *
 * All names, CAPIDs and addresses here are synthetic.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker, Utilities, Session } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'notifications', 'ParentEmailNotify.gs');
const { section, check, done } = makeChecker();
const { logger } = makeLogger();

/** Member.txt row: [0] CAPID, [2] last, [3] first, [11] ORGID, [13] unit, [21] type, [24] status */
function memberRow(capid, orgid, first, last, opts) {
  const o = opts || {};
  const row = [];
  row[0] = capid;
  row[2] = last;
  row[3] = first;
  row[11] = orgid;
  row[13] = o.unit || '101';
  row[21] = o.type || 'CADET';
  row[24] = o.status || 'ACTIVE';
  return row;
}

/** MbrContact rows: [capid, type, priority, value, , , doNotContact] */
function contact(capid, value, opts) {
  const o = opts || {};
  const row = [];
  row[0] = capid;
  row[1] = o.type || 'CADET PARENT EMAIL';
  row[2] = '1';
  row[3] = value;
  row[6] = o.doNotContact ? 'True' : 'False';
  return row;
}

/**
 * Loads the module with a controllable CAPWATCH extract.
 *
 * @param {Object} fixture - { contacts, members, squadrons }
 */
function load(fixture) {
  const f = fixture || {};
  return loadModule(MODULE, {
    Logger: logger,
    Utilities: Utilities,
    Session: Session,
    CONFIG: { CAPWATCH_DATA_FOLDER_ID: 'folder' },
    PROFILE_: { RUN_PARENT_EMAIL_NOTIFICATIONS: true },
    parseFile: name => {
      if (name === 'MbrContact') return f.contacts || [];
      if (name === 'Member') return f.memberRows || [];
      return [];
    },
    // getMembers is deliberately NOT provided. On the cadets tenant it hides
    // every cadet-lite member, and this module must never reach for it.
    getSquadrons: () => f.squadrons || {}
  }, ['peIsSuppressed_', 'peResolveToUnits_', 'peParseIsoDate_', 'PARENT_EMAIL_NOTIFY_CONFIG']);
}

const m = load({});
const suppressed = m.peIsSuppressed_;

// ---------------------------------------------------------------------------
section('1. The cooldown is three calendar months, not ninety days');
{
  check('configured window', m.PARENT_EMAIL_NOTIFY_CONFIG.SUPPRESSION_MONTHS, 3);
  check('same day, suppressed', suppressed('2026-07-01', '2026-07-01'), true);
  check('one day short of three months, still suppressed',
    suppressed('2026-07-03', '2026-10-02'), true);
  check('exactly three months later, reportable again',
    suppressed('2026-07-03', '2026-10-03'), false);
  check('long past, reportable', suppressed('2026-01-01', '2026-07-01'), false);
  check('across a year boundary', suppressed('2026-11-15', '2027-01-15'), true);
  check('and clear of it', suppressed('2026-11-15', '2027-02-15'), false);
}

// ---------------------------------------------------------------------------
section('2. A member never reported is always reportable');
{
  check('no record', suppressed(undefined, '2026-07-01'), false);
  check('empty string', suppressed('', '2026-07-01'), false);
  check('null', suppressed(null, '2026-07-01'), false);
  check('unparseable date is treated as never reported, not as suppressed',
    suppressed('not-a-date', '2026-07-01'), false);
  check('a partial date is not silently accepted',
    suppressed('2026-07', '2026-07-01'), false);
}

// ---------------------------------------------------------------------------
section('3. Ledger addresses resolve to the cadet who owns the record');
{
  const fixture = {
    contacts: [
      contact('100001', 'parent.one@example.org'),
      contact('100002', 'parent.two@example.org')
    ],
    memberRows: [
      memberRow('100001', '900', 'Alex', 'Rivera'),
      memberRow('100002', '901', 'Sam', 'Okafor')
    ],
    squadrons: { '900': { charter: 'PCR-XX-001' }, '901': { charter: 'PCR-XX-002' } }
  };
  const mod = load(fixture);
  const summary = { resolved: 0, unresolved: 0 };
  const byOrg = mod.peResolveToUnits_([
    { member: 'parent.one@example.org', reason: 'no such account', firstSeen: '2026-07-01' },
    { member: 'parent.two@example.org', reason: 'malformed', firstSeen: '2026-07-01' }
  ], summary);

  check('grouped by unit', Object.keys(byOrg).sort(), ['900', '901']);
  check('resolved count', summary.resolved, 2);
  check('none left unresolved', summary.unresolved, 0);
  check('the cadet is named', byOrg['900'][0].cadetName, 'Alex Rivera');
  check('the charter comes from the squadron', byOrg['900'][0].charter, 'PCR-XX-001');
  check('the reason is carried through', byOrg['901'][0].reason, 'malformed');
  check('the suppression key is member AND address',
    byOrg['900'][0].key, '100001|parent.one@example.org');
}

// ---------------------------------------------------------------------------
section('4. Cadet-lite members resolve — the ones this digest exists for');
{
  // A cadet below the account-holding grade is filtered out of getMembers() on
  // the cadets tenant. Resolving against that set dropped ten of thirteen
  // addresses on the first live preview. These cadets hold no account, so their
  // parent's address is the ONLY way a unit list reaches the family — they are
  // the population the digest is for, not an edge case.
  const mod = load({
    contacts: [contact('100010', 'lite.parent@example.org')],
    memberRows: [memberRow('100010', '900', 'Robin', 'Nakamura', { rank: 'C/Amn' })],
    squadrons: { '900': { charter: 'PCR-XX-001' } }
  });
  const summary = { resolved: 0, unresolved: 0 };
  const byOrg = mod.peResolveToUnits_(
    [{ member: 'lite.parent@example.org', reason: 'no such account' }], summary);

  check('the cadet is found in the raw extract', summary.resolved, 1);
  check('and reaches their unit', Object.keys(byOrg), ['900']);
  check('named from Member.txt', byOrg['900'][0].cadetName, 'Robin Nakamura');
}

// ---------------------------------------------------------------------------
section('5. Only ACTIVE cadets, and only real units');
{
  const base = {
    contacts: [contact('100011', 'p@example.org')],
    squadrons: { '900': { charter: 'PCR-XX-001' } }
  };
  const resolveWith = rows => {
    const s = { resolved: 0, unresolved: 0 };
    const mod = load(Object.assign({}, base, { memberRows: rows }));
    return Object.keys(mod.peResolveToUnits_([{ member: 'p@example.org' }], s));
  };

  check('an expired member is not reported',
    resolveWith([memberRow('100011', '900', 'A', 'B', { status: 'EXPIRED' })]), []);
  check('a senior is not a cadet',
    resolveWith([memberRow('100011', '900', 'A', 'B', { type: 'SENIOR' })]), []);
  check('the 000 holding unit is skipped',
    resolveWith([memberRow('100011', '900', 'A', 'B', { unit: '000' })]), []);
  check('an ordinary active cadet is reported',
    resolveWith([memberRow('100011', '900', 'A', 'B')]), ['900']);
}

// ---------------------------------------------------------------------------
section('6. Nobody is mailed about a cadet who is not there');
{
  const mod = load({
    contacts: [contact('100003', 'gone@example.org')],
    memberRows: [],                                 // cadet no longer active here
    squadrons: {}
  });
  const summary = { resolved: 0, unresolved: 0 };
  const byOrg = mod.peResolveToUnits_(
    [{ member: 'gone@example.org', reason: 'no such account' }], summary);

  check('no unit receives a row', Object.keys(byOrg), []);
  check('and it is counted as unresolved, not silently dropped', summary.unresolved, 1);
}

// ---------------------------------------------------------------------------
section('7. Only parent EMAIL contacts, and only contactable ones');
{
  const memberRows = [memberRow('100004', '900', 'Jo', 'Kim')];
  const squadrons = { '900': { charter: 'PCR-XX-001' } };

  const wrongType = load({
    contacts: [contact('100004', 'x@example.org', { type: 'CADET PARENT PHONE' })],
    memberRows: memberRows, squadrons: squadrons
  });
  let s = { resolved: 0, unresolved: 0 };
  check('a phone row is not an email row',
    Object.keys(wrongType.peResolveToUnits_([{ member: 'x@example.org' }], s)), []);

  const dnc = load({
    contacts: [contact('100004', 'x@example.org', { doNotContact: true })],
    memberRows: memberRows, squadrons: squadrons
  });
  s = { resolved: 0, unresolved: 0 };
  check('do-not-contact is honoured',
    Object.keys(dnc.peResolveToUnits_([{ member: 'x@example.org' }], s)), []);
}

// ---------------------------------------------------------------------------
section('8. One address, several cadets — siblings');
{
  const mod = load({
    contacts: [
      contact('100005', 'shared@example.org'),
      contact('100006', 'shared@example.org')
    ],
    memberRows: [
      memberRow('100005', '900', 'Ada', 'Lin'),
      memberRow('100006', '900', 'Bo', 'Lin')
    ],
    squadrons: { '900': { charter: 'PCR-XX-001' } }
  });
  const summary = { resolved: 0, unresolved: 0 };
  const byOrg = mod.peResolveToUnits_([{ member: 'shared@example.org' }], summary);

  check('both cadets are listed', byOrg['900'].length, 2);
  check('with distinct suppression keys',
    byOrg['900'][0].key !== byOrg['900'][1].key, true);
  check('one ledger address, two rows', summary.resolved, 2);
}

// ---------------------------------------------------------------------------
section('9. Address matching is case- and whitespace-insensitive');
{
  const mod = load({
    contacts: [contact('100007', '  Parent.Mixed@Example.org  ')],
    memberRows: [memberRow('100007', '900', 'Cy', 'Vale')],
    squadrons: { '900': { charter: 'PCR-XX-001' } }
  });
  const summary = { resolved: 0, unresolved: 0 };
  const byOrg = mod.peResolveToUnits_(
    [{ member: 'PARENT.MIXED@EXAMPLE.ORG' }], summary);

  check('matched despite case and padding', summary.resolved, 1);
  check('and stored lowercased', byOrg['900'][0].address, 'parent.mixed@example.org');
}

done();
