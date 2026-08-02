/**
 * SendRetentionEmail.gs — holding units are not units.
 *
 * CA-000 (1297) and CA-999 (368) are administrative parking orgs. Nobody
 * commands them, so retention mail about a member held there has no command
 * channel: the 2026-08-01 run produced a digest for ORGID 1297 addressed to
 * nobody. Excluding at selection keeps those members out of the mail AND the
 * digest in one place.
 *
 * The exclusion is by ORGID, not unit number, and trims — the CAPWATCH feed
 * ships padded values, which is what silently defeated an earlier exact-match
 * query during this work.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker, Session, Utilities } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'recruiting-and-retention', 'SendRetentionEmail.gs');
const { section, check, done } = makeChecker();

// Member.txt columns this module reads:
// CAPID=0, NameLast=2, NameFirst=3, DOB=7, ORGID=11, Rank=14, Expiration=16,
// Type=21, MbrStatus=24.
function memberRow(spec) {
  const row = new Array(25).fill('');
  row[0] = spec.capid;
  row[2] = 'Last' + spec.capid;
  row[3] = 'First';
  row[7] = spec.dob || '';
  row[11] = spec.orgid;
  row[14] = 'Capt';
  row[16] = spec.expiration || '';
  row[21] = spec.type || 'SENIOR';
  row[24] = spec.status || 'ACTIVE';
  return row;
}

const contactRow = (capid, email) => [capid, 'EMAIL', 'PRIMARY', email];

const NOW = new Date('2026-08-15T12:00:00');
const EIGHTEEN_THIS_MONTH = '8/1/2008';   // turns 18 in Aug 2026
const EXPIRES_THIS_MONTH = '08/31/2026';

/**
 * @param {Object} opts - { members, contacts, excluded }
 * @returns {Object} Exported retrieval functions plus recorded logs
 */
function load(opts) {
  const o = opts || {};
  const { logger, calls } = makeLogger();

  const files = { Member: o.members || [], MbrContact: o.contacts || [] };

  const mod = loadModule(MODULE, {
    Logger: logger,
    Session: Session,
    Utilities: Utilities,
    CONFIG: { EXCLUDED_ORG_IDS: o.excluded === undefined ? ['1297', '368'] : o.excluded },
    RETENTION_CONFIG: { AGE_THRESHOLDS: { TRANSITION_TO_SENIOR: 18, CADET_AGE_OUT: 21 } },
    parseFile: name => files[name] || [],
    sanitizeEmail: e => {
      const v = String(e || '').trim().toLowerCase();
      return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v) ? v : null;
    },
    Date: class extends Date {
      constructor(...args) { super(...(args.length ? args : [NOW])); }
    }
  }, ['getMembersTurning18', 'getMembersTurning21', 'getExpiringMembers', 'retentionIsExcludedOrg_']);

  return Object.assign({}, mod, { logCalls: calls });
}

// ---------------------------------------------------------------------------
section('1. The predicate matches by ORGID, and tolerates the padded feed');
{
  const m = load();
  check('CA-000 holding unit', m.retentionIsExcludedOrg_('1297'), true);
  check('CA-999 holding unit', m.retentionIsExcludedOrg_('368'), true);
  check('a real squadron', m.retentionIsExcludedOrg_('070'), false);
  check('padded value still matches', m.retentionIsExcludedOrg_('1297  '), true);
  check('blank is not excluded', m.retentionIsExcludedOrg_(''), false);
  check('undefined is not excluded', m.retentionIsExcludedOrg_(undefined), false);
}

// ---------------------------------------------------------------------------
section('2. Expiring members in a holding unit are not selected');
{
  const m = load({
    members: [
      memberRow({ capid: '600001', orgid: '070', expiration: EXPIRES_THIS_MONTH }),
      memberRow({ capid: '600002', orgid: '1297', expiration: EXPIRES_THIS_MONTH }),
      memberRow({ capid: '600003', orgid: '368', expiration: EXPIRES_THIS_MONTH })
    ],
    contacts: [
      contactRow('600001', 'a@example.com'),
      contactRow('600002', 'b@example.com'),
      contactRow('600003', 'c@example.com')
    ]
  });

  const got = m.getExpiringMembers().map(x => x.capid);
  check('only the real squadron member', got, ['600001']);
  check('both holding units counted as skipped',
    m.logCalls.info.filter(i => i.msg === 'Expiring members retrieved')[0].ctx.holdingUnitSkipped, 2);
}

// ---------------------------------------------------------------------------
section('3. Turning 18 and 21 skip holding units too');
{
  const m18 = load({
    members: [
      memberRow({ capid: '600001', orgid: '070', type: 'CADET', dob: EIGHTEEN_THIS_MONTH }),
      memberRow({ capid: '600002', orgid: '1297', type: 'CADET', dob: EIGHTEEN_THIS_MONTH })
    ],
    contacts: [contactRow('600001', 'a@example.com'), contactRow('600002', 'b@example.com')]
  });
  check('turning 18: real unit only', m18.getMembersTurning18().map(x => x.capid), ['600001']);

  const m21 = load({
    members: [
      memberRow({ capid: '600003', orgid: '070', type: 'CADET', dob: '8/1/2005' }),
      memberRow({ capid: '600004', orgid: '368', type: 'CADET', dob: '8/1/2005' })
    ],
    contacts: [contactRow('600003', 'c@example.com'), contactRow('600004', 'd@example.com')]
  });
  check('turning 21: real unit only', m21.getMembersTurning21().map(x => x.capid), ['600003']);
}

// ---------------------------------------------------------------------------
section('4. Excluding at selection also removes the unit from the digest');
{
  // sendRenewalDigests_ groups whatever getExpiringMembers returns, so a holding
  // unit that yields no members yields no digest — which is what stopped ORGID
  // 1297 producing one addressed to nobody.
  const m = load({
    members: [memberRow({ capid: '600002', orgid: '1297', expiration: EXPIRES_THIS_MONTH })],
    contacts: [contactRow('600002', 'b@example.com')]
  });

  check('nothing selected, so nothing to group', m.getExpiringMembers().length, 0);
}

// ---------------------------------------------------------------------------
section('5. A tenant with no exclusions configured selects everyone');
{
  // The region profile lists a different holding unit; a blank list must not
  // throw or silently drop anybody.
  const m = load({
    excluded: [],
    members: [memberRow({ capid: '600002', orgid: '1297', expiration: EXPIRES_THIS_MONTH })],
    contacts: [contactRow('600002', 'b@example.com')]
  });

  check('no exclusions means no filtering', m.getExpiringMembers().map(x => x.capid), ['600002']);
}

done();
