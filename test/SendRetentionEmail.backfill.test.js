/**
 * SendRetentionEmail.gs — reconstructing missing Log rows from sent mail.
 *
 * This writes 'already sent' rows into the audit sheet, and every row it writes
 * suppresses a future email. A wrong row is therefore a silently missed member,
 * which is worse than the duplicate it was meant to prevent. So the cases that
 * matter here are the ones where it must REFUSE to write:
 *
 *   - a member CAPWATCH now selects who was never actually mailed
 *   - a shared family address with more candidates than messages sent
 *
 * and the case where it must write the real send time rather than "now".
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker, Session, Utilities } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'recruiting-and-retention', 'SendRetentionEmail.gs');
const { section, check, done } = makeChecker();

const SUBJECTS = {
  TURNING_18: 'Important Membership Update - Turning 18',
  TURNING_21: 'Important Membership Update - Turning 21',
  EXPIRING: 'Your CAP Membership Expires Soon',
  RENEWAL_DIGEST: 'Memberships expiring this month in your unit'
};

const AUG1 = new Date('2026-08-01T10:43:00');
const SEP2 = new Date('2026-09-02T10:00:00');

const member = (capid, email, type, orgid) => ({
  capid: capid, email: email, type: type || 'SENIOR', orgid: orgid || '070',
  rank: 'Capt', firstName: 'Test', lastName: 'M' + capid, expiration: '08/31/2026'
});

/** A sent message as GmailApp would present it. */
const msg = (subject, to, when) => ({
  getSubject: () => subject,
  getTo: () => to,
  getDate: () => when || AUG1
});

/**
 * @param {Object} [opts] - { messages, t18, t21, expiring, logged }
 * @returns {Object} Exported internals plus recorded appends
 */
function load(opts) {
  const o = opts || {};
  const appended = [];
  const { logger, calls } = makeLogger();

  // Pre-existing Log rows: [Date, type, capid].
  const logRows = o.logged || [];

  // One thread per message keeps the fake honest about the threads->messages walk.
  const threads = (o.messages || []).map(m => ({ getMessages: () => [m] }));

  const mod = loadModule(MODULE, {
    Logger: logger,
    Session: Session,
    Utilities: Utilities,
    RETENTION_CONFIG: {
      SUBJECTS: SUBJECTS,
      // Needed because the real getCommanderInfo runs for cadet rows.
      CC_DUTY_TITLES: { AGE_MILESTONE: ['Deputy Commander for Cadets'], RENEWAL: ['Recruiting Officer'] }
    },
    RETENTION_LOG_SPREADSHEET_ID: 'log',
    GmailApp: {
      search: (query, start) => {
        // Serve only the slice for the subject in this query, once.
        if (start > 0) return [];
        const wanted = (query.match(/subject:"([^"]+)"/) || [])[1];
        return threads.filter(t => t.getMessages()[0].getSubject() === wanted);
      }
    },
    // getCommanderInfo is declared in the module, so it cannot be injected.
    // Starve it instead: an empty extract makes it warn and return null, which
    // is all these cases need.
    parseFile: () => [],
    sanitizeEmail: e => e,
    getActiveUsers: () => [],
    rcBuildCommandDirectoryMap_: () => ({}),
    rcResolveRecipientEmail_: () => null,
    rcDeriveCommandEmail_: () => null,
    formatDutyTitle_: d => String(d || '').trim(),
    // retentionLogAppend_ and retentionAlreadySentThisPeriod_ are declared in
    // the module too, so fake the sheet underneath them rather than replacing
    // them — that way the real read/write path is what these cases exercise.
    SpreadsheetApp: {
      openById: () => ({
        getSheetByName: () => ({
          getLastRow: () => logRows.length + 1,
          getRange: (sr, sc, nr, nc) => ({
            getValues: () => logRows.slice(sr - 2, sr - 2 + nr).map(r => r.slice(sc - 1, sc - 1 + nc))
          }),
          appendRow: row => { appended.push(row); }
        })
      })
    }
  }, ['backfillRetentionLogFromSentMail', 'retentionParseAddresses_', 'retentionGmailWindow_']);

  const candidates = []
    .concat((o.t18 || []).map(m => ({ type: 'TURNING_18', m: m })))
    .concat((o.t21 || []).map(m => ({ type: 'TURNING_21', m: m })))
    .concat((o.expiring || []).map(m => ({ type: 'EXPIRING', m: m })));

  const run = args => mod.backfillRetentionLogFromSentMail(
    Object.assign({ candidates: candidates }, args));

  return Object.assign({}, mod, { run: run, appended: appended, logCalls: calls });
}

// ---------------------------------------------------------------------------
section('1. A mailed member is confirmed and gets the REAL send time');
{
  const when = new Date('2026-08-01T10:43:07');
  const m = load({
    messages: [msg(SUBJECTS.EXPIRING, 'a@example.com', when)],
    expiring: [member('600001', 'a@example.com')]
  });

  const r = m.run({ write: true, period: '2026-08' });

  check('confirmed', r.confirmed, 1);
  check('written', r.written, 1);
  check('timestamp is the send time, not now', m.appended[0][0], when);
  check('type', m.appended[0][1], 'EXPIRING');
  check('capid in the dedupe column', m.appended[0][2], '600001');
  check('address', m.appended[0][4], 'a@example.com');
}

// ---------------------------------------------------------------------------
section('2. Dry run is the default and writes nothing');
{
  const m = load({
    messages: [msg(SUBJECTS.EXPIRING, 'a@example.com')],
    expiring: [member('600001', 'a@example.com')]
  });

  const r = m.run({ period: '2026-08' });

  check('confirmed anyway, so the preview is useful', r.confirmed, 1);
  check('but nothing written', r.written, 0);
  check('and nothing appended', m.appended.length, 0);
}

// ---------------------------------------------------------------------------
section('3. A member CAPWATCH selects now but who was never mailed is NOT written');
{
  // The whole reason for reading sent mail: this member's expiration came into
  // range after the run. Writing a row would suppress a mail they never got.
  const m = load({
    messages: [msg(SUBJECTS.EXPIRING, 'mailed@example.com')],
    expiring: [member('600001', 'mailed@example.com'), member('600002', 'never@example.com')]
  });

  const r = m.run({ write: true, period: '2026-08' });

  check('only the mailed one confirmed', r.confirmed, 1);
  check('only one row', m.appended.length, 1);
  check('and it is the mailed member', m.appended[0][2], '600001');
}

// ---------------------------------------------------------------------------
section('4. Shared family address: two messages, two candidates, both confirmed');
{
  const m = load({
    messages: [
      msg(SUBJECTS.EXPIRING, 'family@example.com'),
      msg(SUBJECTS.EXPIRING, 'family@example.com')
    ],
    expiring: [member('600001', 'family@example.com'), member('600002', 'family@example.com')]
  });

  const r = m.run({ write: true, period: '2026-08' });

  check('both confirmed', r.confirmed, 2);
  check('none ambiguous', r.ambiguous.length, 0);
  check('two rows', m.appended.length, 2);
}

// ---------------------------------------------------------------------------
section('5. Shared address with FEWER messages than candidates refuses to guess');
{
  const m = load({
    messages: [msg(SUBJECTS.EXPIRING, 'family@example.com')],
    expiring: [member('600001', 'family@example.com'), member('600002', 'family@example.com')]
  });

  const r = m.run({ write: true, period: '2026-08' });

  check('only as many confirmed as were sent', r.confirmed, 1);
  check('the excess is reported, not written', r.ambiguous.length, 1);
  check('one row only', m.appended.length, 1);
}

// ---------------------------------------------------------------------------
section('6. Rows already logged are skipped, not duplicated');
{
  const m = load({
    messages: [msg(SUBJECTS.EXPIRING, 'a@example.com')],
    expiring: [member('600001', 'a@example.com')],
    logged: [[AUG1, 'EXPIRING', '600001']]
  });

  const r = m.run({ write: true, period: '2026-08' });

  check('counted as already logged', r.alreadyLogged, 1);
  check('nothing written', m.appended.length, 0);
}

// ---------------------------------------------------------------------------
section('7. Types are kept apart — a birthday mail is not a renewal');
{
  const m = load({
    messages: [
      msg(SUBJECTS.TURNING_18, 'a@example.com'),
      msg(SUBJECTS.EXPIRING, 'b@example.com')
    ],
    t18: [member('600001', 'a@example.com', 'CADET')],
    expiring: [member('600002', 'b@example.com')]
  });

  const r = m.run({ write: true, period: '2026-08' });

  check('both confirmed', r.confirmed, 2);
  const types = m.appended.map(row => row[1]).sort();
  check('one of each type', types, ['EXPIRING', 'TURNING_18']);
}

// ---------------------------------------------------------------------------
section('8. A cross-type address collision does not confirm the wrong mail');
{
  // Only a TURNING_18 went to this address; an EXPIRING candidate shares it.
  const m = load({
    messages: [msg(SUBJECTS.TURNING_18, 'shared@example.com')],
    t18: [member('600001', 'shared@example.com', 'CADET')],
    expiring: [member('600002', 'shared@example.com')]
  });

  const r = m.run({ write: true, period: '2026-08' });

  check('only the birthday mail confirmed', r.confirmed, 1);
  check('and it is the right member', m.appended[0][2], '600001');
  check('the renewal candidate is untouched', m.appended.length, 1);
}

// ---------------------------------------------------------------------------
section('9. Messages outside the period are ignored');
{
  const m = load({
    messages: [msg(SUBJECTS.EXPIRING, 'a@example.com', SEP2)],
    expiring: [member('600001', 'a@example.com')]
  });

  const r = m.run({ write: true, period: '2026-08' });

  check('no sent mail counted for August', r.sentFound, 0);
  check('nothing confirmed', r.confirmed, 0);
  check('nothing written', m.appended.length, 0);
}

// ---------------------------------------------------------------------------
section('10. The unit digest is not mistaken for a member send');
{
  // Digest subjects start with the member-facing text plus a charter suffix.
  const m = load({
    messages: [msg(SUBJECTS.RENEWAL_DIGEST + ' — PCR-CA-070', 'cc@cawgcap.org')],
    expiring: [member('600001', 'cc@cawgcap.org')]
  });

  const r = m.run({ write: true, period: '2026-08' });

  check('digest not counted', r.sentFound, 0);
  check('nothing written', m.appended.length, 0);
}

// ---------------------------------------------------------------------------
section('11. An empty mailbox writes nothing and says why');
{
  const m = load({ expiring: [member('600001', 'a@example.com')] });
  const r = m.run({ write: true, period: '2026-08' });

  check('nothing found', r.sentFound, 0);
  check('nothing written', m.appended.length, 0);
  check('no rows invented from CAPWATCH alone', r.confirmed, 0);
}

// ---------------------------------------------------------------------------
section('12. To-header parsing handles display names and multiple recipients');
{
  const m = load();
  check('bare address', m.retentionParseAddresses_('a@b.org'), ['a@b.org']);
  check('display name form', m.retentionParseAddresses_('Test M <A@B.ORG>'), ['a@b.org']);
  check('multiple', m.retentionParseAddresses_('a@b.org, Name <c@d.org>'), ['a@b.org', 'c@d.org']);
  check('junk dropped', m.retentionParseAddresses_('not-an-address'), []);
  check('empty', m.retentionParseAddresses_(''), []);
}

// ---------------------------------------------------------------------------
section('13. The Gmail window brackets the period');
{
  const m = load();
  check('August 2026', m.retentionGmailWindow_('2026-08'), 'after:2026/7/31 before:2026/9/2');
  check('January rolls back a year', m.retentionGmailWindow_('2026-01'), 'after:2025/12/31 before:2026/2/2');
}

done();
