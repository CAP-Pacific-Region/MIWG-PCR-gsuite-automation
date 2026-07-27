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
    parseFile: name => (name === 'MbrContact' ? (f.contacts || []) : []),
    getMembers: () => f.members || {},
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
    members: {
      '100001': { capsn: '100001', orgid: '900', firstName: 'Alex', lastName: 'Rivera' },
      '100002': { capsn: '100002', orgid: '901', firstName: 'Sam', lastName: 'Okafor' }
    },
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
section('4. Nobody is mailed about a cadet who is not there');
{
  const mod = load({
    contacts: [contact('100003', 'gone@example.org')],
    members: {},                                    // cadet no longer active here
    squadrons: {}
  });
  const summary = { resolved: 0, unresolved: 0 };
  const byOrg = mod.peResolveToUnits_(
    [{ member: 'gone@example.org', reason: 'no such account' }], summary);

  check('no unit receives a row', Object.keys(byOrg), []);
  check('and it is counted as unresolved, not silently dropped', summary.unresolved, 1);
}

// ---------------------------------------------------------------------------
section('5. Only parent EMAIL contacts, and only contactable ones');
{
  const members = { '100004': { capsn: '100004', orgid: '900', firstName: 'Jo', lastName: 'Kim' } };
  const squadrons = { '900': { charter: 'PCR-XX-001' } };

  const wrongType = load({
    contacts: [contact('100004', 'x@example.org', { type: 'CADET PARENT PHONE' })],
    members: members, squadrons: squadrons
  });
  let s = { resolved: 0, unresolved: 0 };
  check('a phone row is not an email row',
    Object.keys(wrongType.peResolveToUnits_([{ member: 'x@example.org' }], s)), []);

  const dnc = load({
    contacts: [contact('100004', 'x@example.org', { doNotContact: true })],
    members: members, squadrons: squadrons
  });
  s = { resolved: 0, unresolved: 0 };
  check('do-not-contact is honoured',
    Object.keys(dnc.peResolveToUnits_([{ member: 'x@example.org' }], s)), []);
}

// ---------------------------------------------------------------------------
section('6. One address, several cadets — siblings');
{
  const mod = load({
    contacts: [
      contact('100005', 'shared@example.org'),
      contact('100006', 'shared@example.org')
    ],
    members: {
      '100005': { capsn: '100005', orgid: '900', firstName: 'Ada', lastName: 'Lin' },
      '100006': { capsn: '100006', orgid: '900', firstName: 'Bo', lastName: 'Lin' }
    },
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
section('7. Address matching is case- and whitespace-insensitive');
{
  const mod = load({
    contacts: [contact('100007', '  Parent.Mixed@Example.org  ')],
    members: { '100007': { capsn: '100007', orgid: '900', firstName: 'Cy', lastName: 'Vale' } },
    squadrons: { '900': { charter: 'PCR-XX-001' } }
  });
  const summary = { resolved: 0, unresolved: 0 };
  const byOrg = mod.peResolveToUnits_(
    [{ member: 'PARENT.MIXED@EXAMPLE.ORG' }], summary);

  check('matched despite case and padding', summary.resolved, 1);
  check('and stored lowercased', byOrg['900'][0].address, 'parent.mixed@example.org');
}

done();
