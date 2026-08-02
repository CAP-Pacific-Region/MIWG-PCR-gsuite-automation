/**
 * SendRetentionEmail.gs — the per-unit renewal digest.
 *
 * A senior's renewal notice carries no unit CC: it is between them and the wing.
 * The unit still needs to know who is lapsing, so it arrives as a digest
 * addressed TO the command channel rather than as a blind copy of somebody
 * else's mail. Cadets appear in the digest too and separately keep the CC on
 * their own notice, which is a cadet protection matter.
 *
 * What is worth pinning: who it reaches when a unit is missing one of the two
 * roles, that it lists every expiring member rather than only the ones this run
 * happened to mail, and that it is deduplicated per unit so a re-run does not
 * mail a unit twice.
 *
 * The staff index is driven through the REAL resolution path rather than faked,
 * so the digest and the CC cannot end up addressing units differently.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker, Session, Utilities } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'recruiting-and-retention', 'SendRetentionEmail.gs');
const NOTIFY = path.join(__dirname, '..', 'src', 'notifications', 'RecoveryEmailNotify.gs');

const { section, check, done } = makeChecker();

const CC_DUTY_TITLES = { AGE_MILESTONE: ['Deputy Commander for Cadets'], RENEWAL: ['Recruiting Officer'] };

const CONFIG = {
  DOMAIN: 'cawgcap.org',
  EMAIL_DOMAIN: '@cawgcap.org',
  COMMAND_EMAIL_DOMAIN: '@cawgcap.org',
  ORG_LABEL: 'CAWG',
  WING_NAME: 'California Wing'
};

const SQUADRONS = {
  '070': { orgid: '070', name: 'Alpha Composite Squadron', charter: 'PCR-CA-070' },
  '071': { orgid: '071', name: 'Bravo Senior Squadron', charter: 'PCR-CA-071' }
};

const member = (capid, orgid, type, last) => ({
  capid: capid, orgid: orgid, type: type || 'SENIOR',
  rank: 'Capt', firstName: 'Test', lastName: last || ('Member' + capid),
  email: 'm' + capid + '@example.com', expiration: '7/31/2026'
});

// Commanders.txt: ORGID=0, CAPID=4, NameLast=8, NameFirst=9, Rank=12.
function commanderRow(orgid, capid) {
  const row = new Array(13).fill('');
  row[0] = orgid; row[4] = capid; row[8] = 'Cmdr'; row[9] = 'Unit' + orgid; row[12] = 'Maj';
  return row;
}

// DutyPosition.txt: CAPID=0, Duty=1, Asst=4, ORGID=7.
function dutyRow(capid, duty, orgid) {
  const row = new Array(8).fill('');
  row[0] = capid; row[1] = duty; row[4] = '0'; row[7] = orgid;
  return row;
}

// Member.txt: CAPID=0, NameLast=2, NameFirst=3, Rank=14.
function memberRow(capid, last, first) {
  const row = new Array(25).fill('');
  row[0] = capid; row[2] = last; row[3] = first; row[14] = 'Capt';
  return row;
}

/**
 * Loads the module with the REAL staff index in play. `accounts` pins each
 * person's address via the directory, the top of the resolution chain.
 *
 * @param {Object} [opts] - { commanders, duties, members, accounts, sendThrows }
 * @returns {Object} Exported internals plus recorded sends
 */
function load(opts) {
  const o = opts || {};
  const sent = [];
  const logged = [];
  const { logger, calls } = makeLogger();

  const notify = loadModule(NOTIFY, {
    Logger: makeLogger().logger, Session: Session, Utilities: Utilities,
    CONFIG: CONFIG, parseFile: () => [], createEmailMap: () => ({}), getActiveUsers: () => []
  }, ['rcBuildCommandDirectoryMap_', 'rcResolveRecipientEmail_', 'rcDeriveCommandEmail_', 'rcEscapeHtml_']);

  const files = {
    Commanders: o.commanders || [],
    DutyPosition: o.duties || [],
    Member: o.members || [],
    MbrContact: []
  };

  const mod = loadModule(MODULE, {
    Logger: logger,
    Session: Session,
    Utilities: Object.assign({}, Utilities, { sleep: () => {} }),
    CONFIG: CONFIG,
    RETENTION_CONFIG: {
      CC_DUTY_TITLES: CC_DUTY_TITLES,
      EMAIL_DELAY_MS: 0,
      SUBJECTS: { RENEWAL_DIGEST: 'Memberships expiring this month in your unit' }
    },
    DIRECTOR_RECRUITING_EMAIL: 'ca.dty.director-recruiting@cawgcap.org',
    AUTOMATION_SENDER_EMAIL: 'automation@cawgcap.org',
    SENDER_NAME: 'CAWG Information Technology',
    parseFile: name => files[name] || [],
    sanitizeEmail: e => {
      const v = String(e || '').trim().toLowerCase();
      return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v) ? v : null;
    },
    formatDutyTitle_: d => String(d || '').trim().replace(/\s+/g, ' '),
    getActiveUsers: () => o.accounts || [],
    getSquadrons: () => SQUADRONS,
    rcBuildCommandDirectoryMap_: notify.rcBuildCommandDirectoryMap_,
    rcResolveRecipientEmail_: notify.rcResolveRecipientEmail_,
    rcDeriveCommandEmail_: notify.rcDeriveCommandEmail_,
    rcEscapeHtml_: notify.rcEscapeHtml_,
    // The module declares retentionLogAppend_ itself, so it cannot be injected —
    // a declaration shadows the parameter. Fake the sheet underneath it instead,
    // which also exercises the real append path.
    RETENTION_LOG_SPREADSHEET_ID: 'log-sheet',
    SpreadsheetApp: {
      openById: () => ({
        getSheetByName: () => ({ appendRow: row => { logged.push(row); } })
      })
    },
    executeWithRetry: fn => fn(),
    GmailApp: {
      sendEmail: (to, subject, body, options) => {
        if (o.sendThrows) throw new Error('Invalid argument');
        sent.push({
          to: to, subject: subject,
          cc: (options || {}).cc || '',
          html: (options || {}).htmlBody || ''
        });
      }
    }
  }, ['sendRenewalDigests_', 'buildRenewalDigestHtml_', 'logEmailSent', 'retentionLogAppend_']);

  return Object.assign({}, mod, { sent: sent, logged: logged, logCalls: calls });
}

const CC070 = 'cc070@cawgcap.org';
const RO070 = 'ro070@cawgcap.org';
const CC071 = 'cc071@cawgcap.org';

const ONLY_CC = { commanders: [commanderRow('070', '1')], accounts: [{ capid: '1', email: CC070 }] };

// ---------------------------------------------------------------------------
section('1. Addressed to the commander, recruiting officer copied');
{
  const m = load({
    commanders: [commanderRow('070', '1')],
    duties: [dutyRow('2', 'Recruiting Officer', '070')],
    members: [memberRow('2', 'Ro', 'Unit')],
    accounts: [{ capid: '1', email: CC070 }, { capid: '2', email: RO070 }]
  });
  const r = m.sendRenewalDigests_([member('600001', '070', 'SENIOR')], { usable: true, keys: {} });

  check('one digest', m.sent.length, 1);
  check('to the commander', m.sent[0].to, CC070);
  check('recruiting officer copied', m.sent[0].cc, RO070);
  check('subject carries the charter', m.sent[0].subject,
    'Memberships expiring this month in your unit — PCR-CA-070');
  check('counted', r.sent, 1);
}

// ---------------------------------------------------------------------------
section('2. No commander: the recruiting officer becomes the addressee');
{
  const m = load({
    commanders: [],
    duties: [dutyRow('2', 'Recruiting Officer', '070')],
    members: [memberRow('2', 'Ro', 'Unit')],
    accounts: [{ capid: '2', email: RO070 }]
  });
  m.sendRenewalDigests_([member('600001', '070')], { usable: true, keys: {} });

  check('addressed to the recruiting officer', m.sent[0].to, RO070);
  check('nobody copied', m.sent[0].cc, '');
}

// ---------------------------------------------------------------------------
section('3. No recruiting officer: commander alone still gets it');
{
  const m = load(ONLY_CC);
  m.sendRenewalDigests_([member('600001', '070')], { usable: true, keys: {} });

  check('addressed to the commander', m.sent[0].to, CC070);
  check('no cc', m.sent[0].cc, '');
}

// ---------------------------------------------------------------------------
section('4. Neither role reachable: reported, not silently dropped');
{
  const m = load({ commanders: [], accounts: [] });
  const r = m.sendRenewalDigests_(
    [member('600001', '070'), member('600002', '070')], { usable: true, keys: {} });

  check('nothing sent', m.sent.length, 0);
  check('reported with a member count', r.noRecipients,
    [{ orgid: '070', orgName: 'Alpha Composite Squadron', members: 2 }]);
  check('warned', m.logCalls.warn.some(w => /No renewal-digest recipient/.test(w.msg)), true);
}

// ---------------------------------------------------------------------------
section('5. One person holding both roles is addressed once, not copied to self');
{
  const m = load({
    commanders: [commanderRow('070', '1')],
    duties: [dutyRow('1', 'Recruiting Officer', '070')],
    members: [memberRow('1', 'Cmdr', 'Unit070')],
    accounts: [{ capid: '1', email: CC070 }]
  });
  m.sendRenewalDigests_([member('600001', '070')], { usable: true, keys: {} });

  check('to', m.sent[0].to, CC070);
  check('no self-cc', m.sent[0].cc, '');
}

// ---------------------------------------------------------------------------
section('6. One digest per unit, listing that unit only');
{
  const m = load({
    commanders: [commanderRow('070', '1'), commanderRow('071', '3')],
    accounts: [{ capid: '1', email: CC070 }, { capid: '3', email: CC071 }]
  });

  m.sendRenewalDigests_([
    member('600001', '070', 'SENIOR', 'Alpha'),
    member('600002', '070', 'CADET', 'Bravo'),
    member('600003', '071', 'SENIOR', 'Charlie')
  ], { usable: true, keys: {} });

  check('two digests', m.sent.length, 2);
  const to070 = m.sent.filter(s => s.to === CC070)[0];
  check('070 lists its own two', /Alpha/.test(to070.html) && /Bravo/.test(to070.html), true);
  check('070 does not list the other unit member', /Charlie/.test(to070.html), false);
}

// ---------------------------------------------------------------------------
section('7. Cadets and seniors both appear — it is the whole unit worklist');
{
  const m = load(ONLY_CC);
  m.sendRenewalDigests_([
    member('600001', '070', 'SENIOR', 'Seniorone'),
    member('600002', '070', 'CADET', 'Cadetone')
  ], { usable: true, keys: {} });

  const html = m.sent[0].html;
  check('senior listed', /Seniorone/.test(html), true);
  check('cadet listed', /Cadetone/.test(html), true);
  check('both counted in the lead line', /2 members/.test(html), true);
}

// ---------------------------------------------------------------------------
section('8. Deduplicated per unit: a re-run does not re-mail');
{
  const m = load(ONLY_CC);
  const r = m.sendRenewalDigests_([member('600001', '070')],
    { usable: true, keys: { 'RENEWAL_DIGEST|070': true } });

  check('nothing sent', m.sent.length, 0);
  check('counted as skipped', r.skipped, 1);
}

// ---------------------------------------------------------------------------
section('9. The dedupe row keys on the ORGID, so the guard can find it');
{
  const m = load(ONLY_CC);
  m.sendRenewalDigests_([member('600001', '070')], { usable: true, keys: {} });

  check('one row logged', m.logged.length, 1);
  check('type in column B', m.logged[0][1], 'RENEWAL_DIGEST');
  check('ORGID in column C', m.logged[0][2], '070');
  check('timestamp in column A is a Date', m.logged[0][0] instanceof Date, true);
  check('eight columns, matching the sheet', m.logged[0].length, 8);
}

// ---------------------------------------------------------------------------
section('10. An unusable log means send, not skip');
{
  const m = load(ONLY_CC);
  // Same key present, but the guard could not read the log this run.
  m.sendRenewalDigests_([member('600001', '070')],
    { usable: false, keys: { 'RENEWAL_DIGEST|070': true } });

  check('fails open, like the per-member guard', m.sent.length, 1);
}

// ---------------------------------------------------------------------------
section('11. A send failure is recorded rather than aborting the run');
{
  const m = load({
    commanders: [commanderRow('070', '1'), commanderRow('071', '3')],
    accounts: [{ capid: '1', email: CC070 }, { capid: '3', email: CC071 }],
    sendThrows: true
  });

  const r = m.sendRenewalDigests_(
    [member('600001', '070'), member('600002', '071')], { usable: true, keys: {} });

  check('both attempted, both recorded', r.failed.length, 2);
  check('none counted sent', r.sent, 0);
  check('nothing logged as sent', m.logged.length, 0);
}

// ---------------------------------------------------------------------------
section('12. Members with no ORGID cannot be grouped, and are skipped');
{
  const m = load(ONLY_CC);
  const r = m.sendRenewalDigests_([
    member('600001', '070'),
    Object.assign(member('600002', '070'), { orgid: '' })
  ], { usable: true, keys: {} });

  check('one digest', r.sent, 1);
  check('lists only the groupable member', /1 member /.test(m.sent[0].html), true);
}

// ---------------------------------------------------------------------------
section('13. Nothing expiring means no digests at all');
{
  const m = load(ONLY_CC);
  const r = m.sendRenewalDigests_([], { usable: true, keys: {} });

  check('no sends', m.sent.length, 0);
  check('no counts', [r.sent, r.skipped, r.failed.length], [0, 0, 0]);
}

// ---------------------------------------------------------------------------
section('14. Member data is escaped into the table');
{
  // Mixed case on purpose. rcEscapeHtml_ replaces angle brackets character by
  // character, so casing is irrelevant to it — but only a mixed-case payload
  // actually demonstrates that, and a lower-case-only assertion would pass just
  // as happily against an escaper that special-cased tag names.
  const m = load(ONLY_CC);
  m.sendRenewalDigests_([member('600001', '070', 'SENIOR', '<SCRIPT>x</ScRiPt>')],
    { usable: true, keys: {} });

  const html = m.sent[0].html;

  check('escaped form present, casing preserved',
    html.indexOf('&lt;SCRIPT&gt;x&lt;/ScRiPt&gt;') !== -1, true);

  // The real invariant is stronger than "no <script>": escaping is at the
  // character level, so NO raw angle bracket from member data survives. Matched
  // case-insensitively and allowing whitespace, so an upper-case or spaced tag
  // cannot slip past the assertion itself.
  check('no raw script tag survives, in any case or spacing',
    /<\s*\/?\s*script/i.test(html), false);
  check('no raw angle bracket from the payload at all',
    html.indexOf('<SCRIPT') === -1 && html.indexOf('</ScRiPt') === -1, true);
}

// ---------------------------------------------------------------------------
section('15. The per-member log actually writes a row');
{
  // REGRESSION. Extracting retentionLogAppend_() out of logEmailSent() left the
  // old `sheet.appendRow(...)` call behind, referencing a variable that no longer
  // existed in that scope. Every per-member write threw ReferenceError into the
  // catch block, which logged "Failed to log email to spreadsheet" and moved on —
  // so the run looked healthy while the Log, and with it the duplicate-send
  // guard, silently recorded nothing. It reached production on 2026-08-01 and
  // cost that month's 251 rows.
  //
  // logEmailSent() had no test at all, which is why. It has one now.
  const m = load(ONLY_CC);

  m.logEmailSent('EXPIRING',
    { capid: '600001', rank: 'Capt', firstName: 'Test', lastName: 'Member', email: 'm@example.com' },
    { capid: '1', rank: 'Maj', firstName: 'Unit', lastName: 'Cmdr', email: CC070 });

  check('a row was written', m.logged.length, 1);
  check('type in column B', m.logged[0][1], 'EXPIRING');
  check('CAPID in column C — what the dedupe guard reads', m.logged[0][2], '600001');
  check('member address in column E', m.logged[0][4], 'm@example.com');
  check('commander recorded', [m.logged[0][5], m.logged[0][7]], ['1', CC070]);
  check('eight columns', m.logged[0].length, 8);
  check('no failure logged', m.logCalls.error.length, 0);
}

// ---------------------------------------------------------------------------
section('16. A member send with no commander still logs');
{
  const m = load(ONLY_CC);
  m.logEmailSent('TURNING_18',
    { capid: '600002', rank: 'C/Amn', firstName: 'A', lastName: 'B', email: 'c@example.com' },
    null);

  check('row written', m.logged.length, 1);
  check('commander columns blank', [m.logged[0][5], m.logged[0][6], m.logged[0][7]], ['', '', '']);
  check('no failure logged', m.logCalls.error.length, 0);
}

done();
