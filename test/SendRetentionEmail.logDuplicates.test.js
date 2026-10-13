/**
 * retentionFindLogDuplicates_() — what counts as a duplicated log row.
 *
 * WHY THIS EXISTS
 *
 * backfillRetentionLogFromSentMail() read its duplicate-protection keys for the
 * CURRENT month while backfilling whichever month it was given (fixed in 1.10.0),
 * so a backfill of a past month found no keys and appended rows the log already
 * held. This scan is how an operator finds what that left behind, and the whole
 * value of it is in the distinctions below: a scan that over-reports sends
 * someone hunting real sends to delete.
 *
 * The key is period|type|capid, matching retentionAlreadySentThisPeriod_()
 * exactly. A second row for one member in one month for one type is a duplicate
 * because the guard would have suppressed it. Everything else is legitimate:
 *
 *   - the same member in DIFFERENT months (that is the normal shape of the log)
 *   - the same member twice in one month for DIFFERENT types (a cadet can turn
 *     18 and expire in the same August)
 *   - the same CAPID under a different type entirely
 *
 * All CAPIDs and dates here are synthetic.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker, Session, Utilities } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'recruiting-and-retention', 'SendRetentionEmail.gs');
const { section, check, done } = makeChecker();

const m = loadModule(MODULE, {
  Logger: makeLogger().logger,
  Session: Session,
  Utilities: Utilities,
  RETENTION_CONFIG: { SUBJECTS: {}, CC_DUTY_TITLES: {} },
  RETENTION_LOG_SPREADSHEET_ID: 'log',
  SpreadsheetApp: {},
  GmailApp: {},
  parseFile: () => [],
  sanitizeEmail: e => e,
  getActiveUsers: () => [],
  console: console
}, ['retentionFindLogDuplicates_']);

const d = iso => new Date(iso + 'T09:00:00');
const row = (iso, type, capid) => [d(iso), type, capid];

// ---------------------------------------------------------------------------
section('1. A member logged twice in one month for one type is a duplicate');
{
  const r = m.retentionFindLogDuplicates_([
    row('2026-08-01', 'EXPIRING', '600001'),
    row('2026-08-14', 'EXPIRING', '600001')
  ]);

  check('one group', r.groups.length, 1);
  check('one row to remove', r.duplicateRows, 1);
  // The original send is the one whose timestamp means anything; the backfill's
  // copy is the fabrication, so the EARLIEST row is the keeper.
  check('keeps the earliest row', r.groups[0].keep.row, 2);
  check('removes the later one', r.groups[0].extras.map(e => e.row), [3]);
  check('named by period, type and CAPID',
    [r.groups[0].period, r.groups[0].type, r.groups[0].capid],
    ['2026-08', 'EXPIRING', '600001']);
}

// ---------------------------------------------------------------------------
section('2. Legitimate repeats are NOT duplicates');
{
  // The normal shape of the log: one row per member per month.
  const months = m.retentionFindLogDuplicates_([
    row('2026-06-02', 'EXPIRING', '600001'),
    row('2026-07-02', 'EXPIRING', '600001'),
    row('2026-08-02', 'EXPIRING', '600001')
  ]);
  check('the same member across months', months.groups.length, 0);

  // A cadet can turn 18 and expire in the same August.
  const types = m.retentionFindLogDuplicates_([
    row('2026-08-02', 'TURNING_18', '600001'),
    row('2026-08-02', 'EXPIRING', '600001')
  ]);
  check('two types in one month for one member', types.groups.length, 0);

  const members = m.retentionFindLogDuplicates_([
    row('2026-08-02', 'EXPIRING', '600001'),
    row('2026-08-02', 'EXPIRING', '600002')
  ]);
  check('two members, same month and type', members.groups.length, 0);

  check('an empty log', m.retentionFindLogDuplicates_([]).groups.length, 0);
  check('a null argument', m.retentionFindLogDuplicates_(null).scanned, 0);
}

// ---------------------------------------------------------------------------
section('3. Rows the duplicate guard cannot read are ignored, not accused');
{
  // These are exactly the rows retentionAlreadySentThisPeriod_() skips. The guard
  // never protected them, so calling them duplicates would be guessing — and
  // would send someone deleting rows on a guess.
  const r = m.retentionFindLogDuplicates_([
    ['2026-08-01', 'EXPIRING', '600001'],        // text date, not a Date
    [d('2026-08-01'), '', '600001'],             // no type
    [d('2026-08-01'), 'EXPIRING', ''],           // no CAPID
    [d('2026-08-01'), 'EXPIRING', '600001'],     // the only readable row
    ['', '', '']                                 // blank trailing row
  ]);

  check('four skipped', r.skipped, 4);
  check('one readable', r.usable, 1);
  check('five scanned', r.scanned, 5);
  check('nothing reported', r.groups.length, 0);
}

// ---------------------------------------------------------------------------
section('4. Three copies report two removals, and groups come back in sheet order');
{
  const r = m.retentionFindLogDuplicates_([
    row('2026-09-05', 'EXPIRING', '600002'),     // row 2  (later group, first in sheet)
    row('2026-08-03', 'EXPIRING', '600001'),     // row 3
    row('2026-08-20', 'EXPIRING', '600001'),     // row 4
    row('2026-09-19', 'EXPIRING', '600002'),     // row 5
    row('2026-08-28', 'EXPIRING', '600001')      // row 6
  ]);

  check('two groups', r.groups.length, 2);
  check('three rows to remove in total', r.duplicateRows, 3);
  // Sheet order by keeper row, so the output reads top-to-bottom like the sheet.
  check('ordered by the row kept', r.groups.map(g => g.keep.row), [2, 3]);
  check('the triple keeps the earliest and removes two',
    r.groups[1].extras.map(e => e.row), [4, 6]);
  check('the pair removes one', r.groups[0].extras.map(e => e.row), [5]);
}

// ---------------------------------------------------------------------------
section('5. Out-of-order timestamps still keep the oldest');
{
  // A backfill appends at the bottom, so the fabricated row is usually LATER in
  // the sheet than the real one. It need not be: rows can be sorted or pasted.
  const r = m.retentionFindLogDuplicates_([
    row('2026-08-28', 'EXPIRING', '600001'),     // row 2, later date
    row('2026-08-03', 'EXPIRING', '600001')      // row 3, the original send
  ]);

  check('keeps row 3, the older date', r.groups[0].keep.row, 3);
  check('removes row 2', r.groups[0].extras.map(e => e.row), [2]);
}

done();
