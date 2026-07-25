/**
 * SendRetentionEmail.gs — the already-sent guard, and the tenant guard.
 *
 * These two decide whether a member gets a duplicate copy of a birthday or
 * expiration notice, so they are worth pinning precisely.
 *
 * The already-sent guard reads the Log sheet the module has always written and
 * never read. The interesting cases are not the happy path but the edges: a
 * missing sheet is a legitimate empty history, an unreadable one is not, and the
 * two must behave differently — one proceeds normally, the other proceeds
 * loudly. It fails OPEN by design: refusing to send because a spreadsheet read
 * failed would turn a logging problem into a silent outage of the whole feature.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker, Session, Utilities } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'recruiting-and-retention', 'SendRetentionEmail.gs');
const { section, check, done } = makeChecker();

const SHEET_ID = 'sheet-abc123';

/**
 * A SpreadsheetApp whose Log sheet holds the given rows.
 *
 * @param {Array<Array>} rows - [Date|any, emailType, capid] triples
 * @param {Object} [opts] - { missing: no Log sheet, throws: open() raises }
 * @returns {Object} SpreadsheetApp stand-in
 */
function makeSpreadsheetApp(rows, opts) {
  const o = opts || {};
  return {
    openById: function (id) {
      if (o.throws) throw new Error('You do not have permission to access the requested document.');
      if (id !== SHEET_ID) throw new Error('Unexpected spreadsheet id: ' + id);
      return {
        getSheetByName: function (name) {
          if (o.missing || name !== 'Log') return null;
          return {
            getLastRow: () => rows.length + 1, // +1 for the header
            getRange: (startRow, startCol, numRows, numCols) => ({
              getValues: () => rows
                .slice(startRow - 2, startRow - 2 + numRows)
                .map(r => r.slice(startCol - 1, startCol - 1 + numCols))
            })
          };
        }
      };
    }
  };
}

/**
 * @param {Object} [opts] - { rows, sheetId, missing, throws, runRetention, profile }
 * @returns {Object} Exported internals plus recorded log calls
 */
function load(opts) {
  const o = opts || {};
  const { logger, calls } = makeLogger();
  const mod = loadModule(MODULE, {
    Logger: logger,
    Session: Session,
    Utilities: Utilities,
    SpreadsheetApp: makeSpreadsheetApp(o.rows || [], o),
    RETENTION_LOG_SPREADSHEET_ID: o.sheetId === undefined ? SHEET_ID : o.sheetId,
    TENANT_PROFILE: o.profile || 'seniors',
    PROFILE_: { RUN_RETENTION_EMAILS: o.runRetention !== false },
    CONFIG: { WING_NAME: 'California Wing', ORG_LABEL: 'CAWG' },
    clearCache: () => { calls.clearCacheCalled = true; },
    parseFile: () => [],
    sanitizeEmail: e => e,
    getActiveUsers: () => []
  }, [
    'retentionPeriodKey_', 'retentionAlreadySentThisPeriod_', 'retentionFilterUnsent_',
    'sendRetentionEmails'
  ]);
  return Object.assign({}, mod, { logCalls: calls });
}

const JULY = new Date('2026-07-25T10:00:00');
const JUNE = new Date('2026-06-25T10:00:00');
const member = capid => ({ capid: capid, rank: 'C/Amn', lastName: 'Test' });

// ---------------------------------------------------------------------------
section('1. Period key is the calendar month, zero-padded');
{
  const m = load();
  check('July 2026', m.retentionPeriodKey_(JULY), '2026-07');
  check('single-digit month padded', m.retentionPeriodKey_(new Date('2026-03-02T00:00:00')), '2026-03');
  check('December', m.retentionPeriodKey_(new Date('2026-12-31T23:59:00')), '2026-12');
}

// ---------------------------------------------------------------------------
section('2. A member mailed this month is dropped; the rest still send');
{
  const m = load({ rows: [
    [JULY, 'EXPIRING', '600001'],
    [JULY, 'EXPIRING', '600002']
  ]});

  const seen = m.retentionAlreadySentThisPeriod_(JULY);
  check('log was usable', seen.usable, true);

  const due = m.retentionFilterUnsent_('EXPIRING', [member('600001'), member('600002'), member('600003')], seen);
  check('only the unmailed one remains', due.map(x => x.capid), ['600003']);
}

// ---------------------------------------------------------------------------
section('3. Last month does not suppress this month');
{
  const m = load({ rows: [[JUNE, 'EXPIRING', '600001']] });
  const seen = m.retentionAlreadySentThisPeriod_(JULY);
  const due = m.retentionFilterUnsent_('EXPIRING', [member('600001')], seen);

  check('June send does not block a July send', due.length, 1);
}

// ---------------------------------------------------------------------------
section('4. The guard is per email type, not per member');
{
  // A cadet can legitimately turn 18 and expire in the same month.
  const m = load({ rows: [[JULY, 'TURNING_18', '600001']] });
  const seen = m.retentionAlreadySentThisPeriod_(JULY);

  check('the birthday mail is suppressed',
    m.retentionFilterUnsent_('TURNING_18', [member('600001')], seen).length, 0);
  check('the expiration mail still goes',
    m.retentionFilterUnsent_('EXPIRING', [member('600001')], seen).length, 1);
}

// ---------------------------------------------------------------------------
section('5. No Log sheet yet is an empty history, not a failure');
{
  const m = load({ missing: true });
  const seen = m.retentionAlreadySentThisPeriod_(JULY);

  check('usable', seen.usable, true);
  check('nothing recorded', Object.keys(seen.keys).length, 0);
  check('did not warn', m.logCalls.warn.length, 0);
  check('everyone still due',
    m.retentionFilterUnsent_('EXPIRING', [member('600001')], seen).length, 1);
}

// ---------------------------------------------------------------------------
section('6. An unreadable log fails OPEN, and says so');
{
  const m = load({ throws: true });
  const seen = m.retentionAlreadySentThisPeriod_(JULY);

  check('not usable', seen.usable, false);
  check('warned', m.logCalls.warn.some(w => /duplicate protection is OFF/i.test(w.msg)), true);
  check('sends proceed rather than the run doing nothing',
    m.retentionFilterUnsent_('EXPIRING', [member('600001'), member('600002')], seen).length, 2);
}

// ---------------------------------------------------------------------------
section('7. An unconfigured log id also fails open, and says so');
{
  const m = load({ sheetId: '' });
  const seen = m.retentionAlreadySentThisPeriod_(JULY);

  check('not usable', seen.usable, false);
  check('warned', m.logCalls.warn.some(w => /duplicate protection is OFF/i.test(w.msg)), true);
}

// ---------------------------------------------------------------------------
section('8. Junk rows are ignored rather than throwing');
{
  const m = load({ rows: [
    ['', '', ''],                       // blank row
    ['2026-07-01', 'EXPIRING', '600009'], // text timestamp, not a Date
    [JULY, '', '600008'],               // no type
    [JULY, 'EXPIRING', ''],             // no capid
    [JULY, 'EXPIRING', '600001']        // the only real one
  ]});

  const seen = m.retentionAlreadySentThisPeriod_(JULY);
  check('usable', seen.usable, true);
  check('only the well-formed row counted', Object.keys(seen.keys), ['EXPIRING|600001']);
  check('the text-timestamp row did not suppress anyone',
    m.retentionFilterUnsent_('EXPIRING', [member('600009')], seen).length, 1);
}

// ---------------------------------------------------------------------------
section('9. CAPIDs compare as trimmed strings, not by type');
{
  const m = load({ rows: [[JULY, 'EXPIRING', 600001]] }); // numeric cell
  const seen = m.retentionAlreadySentThisPeriod_(JULY);

  check('numeric log cell matches a string capid',
    m.retentionFilterUnsent_('EXPIRING', [member('600001')], seen).length, 0);
  check('whitespace on the member side is tolerated',
    m.retentionFilterUnsent_('EXPIRING', [member(' 600001 ')], seen).length, 0);
}

// ---------------------------------------------------------------------------
section('10. Tenant guard: the module is a no-op where the profile disables it');
{
  const m = load({ runRetention: false, profile: 'cadets' });
  const result = m.sendRetentionEmails();

  check('returns skipped', result, { skipped: true });
  check('did not even clear the cache', m.logCalls.clearCacheCalled, undefined);
  check('said why', m.logCalls.info.some(i => /disabled for this tenant profile/.test(i.msg)), true);
}

done();
