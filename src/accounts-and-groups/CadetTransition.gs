/**
 * Cadet → senior tenant transition.
 *
 * Version: 1.2.0
 * Date: 2026-07-27
 * Changes: 1.2.0 — runCadetTransitionPipeline() runs all six phases from ONE
 *   trigger, freeing five of the twenty Apps Script allows per script. Safe to
 *   collapse because the phases were never an hour apart for duration — each
 *   acts only on rows the previous made ready, and the three expensive ones
 *   already self-limit and schedule their own continuations. Preserves the two
 *   properties six triggers gave for free: a failing phase does not stop the
 *   others, and order is kept. Parks the phase it stopped at so a short run
 *   cannot starve the late phases. armTransitionPipelineTrigger() swaps the six
 *   for the one; armTransitionTriggers() still installs the six for anyone who
 *   wants them back.
 *   1.1.0 — added the six-trigger scheduler (armTransitionTriggers /
 *   disarm / list) and the daily close reminder wiring; 1.0.0 — initial release.
 *   Detection + state for the cadet→senior lifecycle; the Transitions sheet is
 *   authoritative for who is mid-flight.
 *
 * When a cadet turns 21, or converts voluntarily after 18, CAPWATCH flips their
 * type and they leave the cadet tenant for the senior one. Before this module
 * existed the cadet tenant simply suspended them and deleted the account ~30
 * days later, destroying the mailbox: Archived User licenses are not provisioned
 * on this edition, so deletion is the only way to free a seat and there is no
 * archive to recover from.
 *
 * This module runs on the CADETS tenant only (TRANSITION_CONFIG.ROLE ===
 * 'source'). It owns the lifecycle end to end — detect, migrate, delete, forward
 * — and polls the peer (senior) directory for the destination account rather
 * than having the two tenants signal each other, so there is no shared state to
 * drift. The senior tenant's only involvement is exempting these members from
 * the Level I gate in updateAllMembers(), so the destination mailbox exists to
 * receive mail.
 *
 * Two facts shape the design, both verified rather than assumed:
 *
 *  1. SA impersonation works against SUSPENDED users. The mailbox stays readable
 *     after suspension, so members are suspended on day 0 exactly as before —
 *     preserving the cap discipline that PATRON accounts blew in June 2026 — and
 *     the mail is migrated at leisure inside the hold window. Nothing here ever
 *     unsuspends an account.
 *
 *  2. LICENSE_CONFIG.DAYS_BEFORE_DELETE_INELIGIBLE cannot be reused as the hold
 *     clock: it times a member LAPSING, which is not what a transitioning cadet
 *     is doing. Their old cadet type expires the moment they convert, so that
 *     clock starts running immediately and says nothing about whether their mail
 *     has been carried across yet. The Transitions sheet's DetectedDate is
 *     therefore authoritative for the 90-day hold, and deleteIneligibleSuspendedUsers()
 *     skips anyone holding an open row here.
 *
 * This file covers detection and state. Migration lives in CadetTransitionMigrate.gs.
 *
 * @see TRANSITION_CONFIG in config.gs
 */

// ============================================================================
// SHEET SCHEMA
// ============================================================================

/**
 * Transitions sheet columns, in order.
 *
 * The sheet is state, not a report: it is the authoritative record of who is
 * mid-flight, when their hold expires, and how far their migration got. It is a
 * sheet rather than Script Properties specifically so a human can see a stuck
 * migration and intervene — which is also why FAILED rows are left in place
 * rather than retried forever.
 */
const TRANSITION_COLUMNS_ = [
  'CAPID',            // CAPWATCH member id — the join key across both tenants
  'Name',             // human reference only
  'CadetEmail',       // source mailbox on the cadet tenant
  'SeniorEmail',      // destination mailbox; blank until the senior account appears
  'NewType',          // CAPWATCH type that triggered detection (SENIOR, PATRON, ...)
  'DetectedDate',     // authoritative start of the hold clock
  'DeleteAfter',      // DetectedDate + TRANSITION_CONFIG.HOLD_DAYS
  'MigrationStatus',  // TRANSITION_CONFIG.STATUS.*
  'MigratedDate',
  'MessagesMigrated',
  'LastCursor',       // Gmail pageToken, so a run that hits the 6-minute limit resumes
  'NotifiedDate',     // when the member was told; blank means they have NOT been told
  'DriveMigrated',    // files copied. BLANK = nobody looked; 0 = deliberately nothing to copy
  'DriveCursor',      // Drive pageToken, for resuming a copy across executions
  'ContactsMigrated', // contacts copied. BLANK = nobody looked; 0 = deliberately none
  'ContactsCursor',   // People API pageToken, for resuming across executions
  'ForwardGroupCreated',
  'ForwardGroupExpires',
  'Notes'
];

/**
 * Cached Sheet handle. SpreadsheetApp.openById() is a slow round trip and the
 * write helpers below are called once per row — without this, a 10-row detection
 * run spent most of its ~135s reopening the same spreadsheet, and a larger
 * backlog would have crept into the 6-minute execution limit and died mid-write.
 * Cleared implicitly when the execution ends, which is the only lifetime needed.
 */
let transitionsSheet_ = null;

/**
 * Invalidates the cached Sheet handle and header. Only needed if the sheet is
 * restructured mid-execution, which ensureTransitionColumns_() does on the run
 * that first adds a column.
 */
function resetTransitionsCache_() {
  transitionsSheet_ = null;
  transitionHeader_ = null;
}

/**
 * Resolves the Transitions sheet, creating it with headers on first use.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getTransitionsSheet_() {
  if (transitionsSheet_) return transitionsSheet_;

  const ss = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(TRANSITION_CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(TRANSITION_CONFIG.SHEET_NAME);
    sheet.appendRow(TRANSITION_COLUMNS_);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, TRANSITION_COLUMNS_.length).setFontWeight('bold');
    Logger.info('Transitions sheet created', { name: TRANSITION_CONFIG.SHEET_NAME });
  }

  ensureTransitionColumns_(sheet);

  transitionsSheet_ = sheet;
  return sheet;
}

/**
 * Appends any columns added to TRANSITION_COLUMNS_ since the sheet was created.
 *
 * The sheet is live state with real rows in it, so a new column cannot just
 * appear in the constant and be assumed present — setTransitionField_ resolves
 * positions from TRANSITION_COLUMNS_, and writing to a column the header does
 * not have would silently scribble into whatever is at that index.
 *
 * Appends only. Never reorders or removes: a human may well have added notes or
 * columns of their own, and this is their sheet too.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function ensureTransitionColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (!lastCol) return;

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || '').trim());

  const missing = TRANSITION_COLUMNS_.filter(c => header.indexOf(c) < 0);
  if (!missing.length) return;

  sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  sheet.getRange(1, 1, 1, lastCol + missing.length).setFontWeight('bold');

  // The header just changed under any cached copy.
  transitionHeader_ = null;

  Logger.info('Transitions sheet columns added', { added: missing });
}

/**
 * Reads every transition row, keyed by CAPID.
 *
 * Dates come back as ISO strings rather than Date objects so that a value
 * round-tripped through the sheet compares the same way whether the cell was
 * written by this code or typed by a human.
 *
 * @returns {Object<string, Object>} CAPID -> row object, plus _rowNumber
 */
function readTransitions_() {
  const sheet = getTransitionsSheet_();
  const values = sheet.getDataRange().getValues();
  const byCapid = {};

  if (values.length < 2) return byCapid;

  const header = values[0].map(h => String(h || '').trim());

  for (let r = 1; r < values.length; r++) {
    const capid = String(values[r][header.indexOf('CAPID')] || '').trim();
    if (!capid) continue;

    const row = { _rowNumber: r + 1 };
    for (let c = 0; c < header.length; c++) {
      const value = values[r][c];
      row[header[c]] = value instanceof Date ? value.toISOString() : value;
    }
    byCapid[capid] = row;
  }

  return byCapid;
}

/** Cached sheet header, so every field write does not re-read row 1. */
let transitionHeader_ = null;

/**
 * The sheet's ACTUAL column order.
 *
 * Positions must come from the sheet, never from TRANSITION_COLUMNS_. The two
 * diverge the moment a column is added to the constant, because
 * ensureTransitionColumns_() can only append to a sheet that already has rows,
 * while the constant may declare the new column in the middle. Resolving writes
 * against the constant in that state puts every field after the insertion point
 * one column to the left — silently, into real data.
 *
 * @returns {Array<string>}
 */
function transitionHeader_get_() {
  if (transitionHeader_) return transitionHeader_;

  const sheet = getTransitionsSheet_();
  transitionHeader_ = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => String(h || '').trim());
  return transitionHeader_;
}

/**
 * Writes one field of one row.
 *
 * @param {number} rowNumber - 1-indexed sheet row
 * @param {string} column - column name from TRANSITION_COLUMNS_
 * @param {*} value
 */
function setTransitionField_(rowNumber, column, value) {
  const index = transitionHeader_get_().indexOf(column);
  if (index < 0) {
    throw new Error('Column not present in the Transitions sheet: ' + column +
      ' (ensureTransitionColumns_ should have added it)');
  }
  getTransitionsSheet_().getRange(rowNumber, index + 1).setValue(value);
}

/**
 * Reads one field of one row, fresh from the sheet.
 *
 * The row objects from readTransitions_() are a snapshot; this reads the live
 * cell, needed when appending to a value (like Notes) that earlier writes in the
 * same execution may have changed.
 *
 * @param {number} rowNumber - 1-indexed sheet row
 * @param {string} column
 * @returns {string}
 */
function getTransitionField_(rowNumber, column) {
  const index = transitionHeader_get_().indexOf(column);
  if (index < 0) return '';
  const v = getTransitionsSheet_().getRange(rowNumber, index + 1).getValue();
  return v instanceof Date ? v.toISOString() : String(v == null ? '' : v);
}

/**
 * Runs fn holding a script-wide lock, so only one transition operation touches
 * the Transitions sheet and the mailboxes/Drives at a time.
 *
 * Without it, a scheduled trigger firing while a continuation (or a manual run)
 * is mid-flight would process the same cursor twice and duplicate imports/copies
 * — the cursor discipline guards an interrupted run, not a concurrent one. The
 * late arrival backs off rather than waiting, matching processSendAsNamesBatch:
 * whoever holds the lock is already doing the work and will schedule the next
 * continuation, so a second execution has nothing to add.
 *
 * Acquired ONLY at top-level entry points, never in the per-member workers they
 * call, so a call chain never tries to take the lock twice.
 *
 * @param {function(): T} fn
 * @param {T} bailValue - returned if the lock is held by another execution
 * @returns {T}
 * @template T
 */
function withTransitionLock_(fn, bailValue) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    Logger.warn('Transition operation skipped — another run holds the script lock');
    return bailValue;
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/**
 * True only for a genuinely empty sheet field — '', null, or undefined.
 *
 * Exists to avoid the falsy-zero trap: `!value` and `value || ''` both treat the
 * NUMBER 0 as empty, so a legitimate "0 items migrated" reads as "never handled"
 * and blocks the close permanently. A count field where 0 is a real, deliberate
 * value must test blankness this way, not by truthiness. (0 blocked every member,
 * since all four had 0 personal contacts.)
 *
 * @param {*} v
 * @returns {boolean}
 */
function isBlankField_(v) {
  return v === '' || v === null || v === undefined;
}

/**
 * Durably records a skipped item that must block deletion — written the instant
 * the skip happens, NOT deferred to a completion handler.
 *
 * This exists because the obvious design (collect skips in an array, write the
 * note when the mailbox finishes) silently loses them: a large migration spans
 * many time-limited executions, and a skip in an execution that then pauses is
 * discarded when that execution ends, while the cursor advances past the skipped
 * item so no later run re-encounters it. The skip gets logged but the note that
 * should refuse the close is never written — which is exactly how a genuinely
 * missing message reached a COMPLETE row with no DO NOT DELETE marker.
 *
 * Idempotent via dedupKey (the source item id): a crash mid-page can re-skip the
 * same item on resume, and appending it twice would just be noise.
 *
 * @param {number} rowNumber
 * @param {string} dedupKey - unique id of the skipped item (message/file id)
 * @param {string} description - human-readable, includes the id
 */
function recordSkip_(rowNumber, dedupKey, description) {
  const current = getTransitionField_(rowNumber, 'Notes');
  if (current.indexOf(dedupKey) > -1) return;   // already recorded this item
  const entry = 'DO NOT DELETE — ' + description;
  setTransitionField_(rowNumber, 'Notes', current ? current + ' | ' + entry : entry);
}

/**
 * Appends a new transition row, ordered to match the sheet rather than the
 * constant — see transitionHeader_get_().
 *
 * @param {Object} row - keys matching TRANSITION_COLUMNS_
 */
function appendTransition_(row) {
  getTransitionsSheet_().appendRow(
    transitionHeader_get_().map(c => row[c] === undefined ? '' : row[c])
  );
}

// ============================================================================
// DETECTION
// ============================================================================

/**
 * Finds cadet-tenant accounts whose CAPWATCH type has moved out of the cadet
 * program, and opens a transition row for each.
 *
 * A lapse and a transition look different in Member.txt and must not be
 * confused — a lapse is status != ACTIVE (or no record at all) and should follow
 * the ordinary suspend-and-delete path, while a transition is status ACTIVE with
 * a non-cadet type and needs the mailbox held. Only the latter lands here.
 *
 * Idempotent: an existing row for a CAPID is never re-dated, so the hold clock
 * cannot be restarted by re-running detection. The one thing it does update is
 * NewType, so a PATRON who converts to SENIOR is picked up by the migration pass.
 *
 * Safe to run on any tenant; no-ops unless ROLE is 'source'.
 *
 * @returns {{detected: number, updated: number, existing: number}}
 */
function detectCadetTransitions() {
  if (TRANSITION_CONFIG.ROLE !== 'source') {
    Logger.info('Transition detection skipped — not the source tenant', {
      role: TRANSITION_CONFIG.ROLE || '(off)'
    });
    return { detected: 0, updated: 0, existing: 0 };
  }
  return withTransitionLock_(detectCadetTransitions_, { detected: 0, updated: 0, existing: 0 });
}

function detectCadetTransitions_() {
  const start = new Date();
  Logger.info('Starting cadet transition detection');

  // CAPID -> {status, type} for every CAPWATCH record.
  const memberData = parseFile('Member');
  const infoByCapid = {};
  for (let i = 0; i < memberData.length; i++) {
    const capid = String(memberData[i][0] || '').trim();
    if (capid) {
      // Column order per createMemberObject(): [0] CAPID, [2] last, [3] first,
      // [21] type, [24] status.
      infoByCapid[capid] = {
        status: memberData[i][24],
        type: memberData[i][21],
        name: normalizeName_(memberData[i][3], memberData[i][2])
      };
    }
  }

  const existing = readTransitions_();
  const now = new Date();
  let detected = 0;
  let updated = 0;
  let alreadyOpen = 0;

  // Both active and suspended accounts: by the time detection runs, the member
  // has usually already been suspended by suspendExpiredMembers().
  eachDirectoryUser_(user => {
    const capid = capidOfUser_(user);
    if (!capid) return;

    const info = infoByCapid[capid];
    if (!info) return;                                  // no CAPWATCH record: a lapse, not a transition
    if (info.status !== 'ACTIVE') return;               // lapsed: ordinary suspend-and-delete path
    if (TRANSITION_CONFIG.TRANSITION_TYPES.indexOf(info.type) < 0) return;  // still a cadet

    const open = existing[capid];
    if (open) {
      // Already tracked. Only NewType may change — a PATRON converting to SENIOR
      // is exactly the case the hold window exists to catch.
      if (open.NewType !== info.type) {
        setTransitionField_(open._rowNumber, 'NewType', info.type);
        Logger.info('Transition type changed', {
          capid: capid,
          from: open.NewType,
          to: info.type
        });
        updated++;
      } else {
        alreadyOpen++;
      }
      return;
    }

    const deleteAfter = new Date(now);
    deleteAfter.setDate(deleteAfter.getDate() + TRANSITION_CONFIG.HOLD_DAYS);

    appendTransition_({
      CAPID: capid,
      Name: info.name,
      CadetEmail: user.primaryEmail,
      SeniorEmail: '',
      NewType: info.type,
      DetectedDate: now.toISOString(),
      DeleteAfter: deleteAfter.toISOString(),
      MigrationStatus: TRANSITION_CONFIG.STATUS.PENDING,
      MigratedDate: '',
      MessagesMigrated: '',
      LastCursor: '',
      ForwardGroupCreated: '',
      ForwardGroupExpires: '',
      Notes: ''
    });

    Logger.info('Cadet transition detected', {
      capid: capid,
      name: info.name,
      cadetEmail: user.primaryEmail,
      newType: info.type,
      deleteAfter: deleteAfter.toISOString()
    });
    detected++;
  });

  Logger.info('Cadet transition detection completed', {
    duration: new Date() - start + 'ms',
    detected: detected,
    typeChanged: updated,
    alreadyOpen: alreadyOpen
  });

  return { detected: detected, updated: updated, existing: alreadyOpen };
}

/**
 * Joins first and last into a display name, collapsing internal whitespace.
 *
 * CAPWATCH name fields carry stray padding — a trailing space on a first name
 * yields a doubled space in the join. Cosmetic (the Name column is human
 * reference only), but it makes the sheet look broken.
 *
 * @param {string} first
 * @param {string} last
 * @returns {string}
 */
function normalizeName_(first, last) {
  return `${first || ''} ${last || ''}`.replace(/\s+/g, ' ').trim();
}

/**
 * Rewrites the Name column on existing rows through normalizeName_().
 *
 * One-shot cleanup for rows written before normalization existed. Touches only
 * that column, and only where the value actually changes.
 *
 * @returns {{fixed: number}}
 */
function normalizeTransitionNames() {
  const rows = readTransitions_();
  let fixed = 0;

  for (const capid in rows) {
    const row = rows[capid];
    const current = String(row.Name || '');
    const cleaned = current.replace(/\s+/g, ' ').trim();

    if (cleaned !== current) {
      setTransitionField_(row._rowNumber, 'Name', cleaned);
      Logger.info('Transition name normalized', {
        capid: capid, from: current, to: cleaned
      });
      fixed++;
    }
  }

  Logger.info('Transition name normalization completed', { fixed: fixed });
  return { fixed: fixed };
}

/**
 * Walks every non-admin user in the local directory, suspended included.
 *
 * @param {function(Object): void} callback
 */
function eachDirectoryUser_(callback) {
  let pageToken = '';
  do {
    const page = AdminDirectory.Users.list({
      customer: 'my_customer',
      maxResults: 500,
      query: 'isAdmin=false',
      fields: 'users(primaryEmail,name,suspended,externalIds),nextPageToken',
      pageToken: pageToken
    });
    pageToken = page.nextPageToken;
    (page.users || []).forEach(callback);
  } while (pageToken);
}

/**
 * Human-readable byte count. Shared by the migration (reporting messages too
 * large to move) and the Drive module (reporting what a copy would shift).
 *
 * @param {number} bytes
 * @returns {string}
 */
function formatBytes_(bytes) {
  if (!bytes) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(1024));
  return (bytes / Math.pow(1024, i)).toFixed(i ? 1 : 0) + ' ' + units[i];
}

/**
 * Extracts a CAPID from a directory user.
 *
 * externalIds is the only source. Other call sites fall back to user.employeeId,
 * but that is dead code: employeeId is not a field on the Directory User
 * resource (it belongs to the People API), so addOrUpdateUser() writing it is
 * silently discarded and reading it always yields undefined. Selecting it in a
 * `fields` mask is a hard 400.
 *
 * @param {Object} user
 * @returns {string} CAPID, or '' if absent
 */
function capidOfUser_(user) {
  const ext = (user.externalIds || []).find(id => id.type === 'organization');
  return ext ? String(ext.value || '').trim() : '';
}

/**
 * CAPIDs whose cadet account must not be deleted by the ordinary license
 * lifecycle, because this module owns their deletion instead.
 *
 * Read by deleteIneligibleSuspendedUsers(). A row stops protecting the account
 * once it is COMPLETE or NOT_APPLICABLE and past DeleteAfter — at which point
 * this module deletes it deliberately.
 *
 * @returns {Object<string, boolean>} CAPID -> true
 */
function getHeldTransitionCapids() {
  if (TRANSITION_CONFIG.ROLE !== 'source') return {};

  const held = {};
  try {
    const rows = readTransitions_();
    const now = new Date();

    for (const capid in rows) {
      const row = rows[capid];
      const status = row.MigrationStatus;

      // FAILED holds indefinitely and on purpose: a failed migration followed by
      // an on-schedule deletion is the one outcome that loses mail for good.
      if (status === TRANSITION_CONFIG.STATUS.FAILED) {
        held[capid] = true;
        continue;
      }

      const deleteAfter = row.DeleteAfter ? new Date(row.DeleteAfter) : null;
      if (deleteAfter && now < deleteAfter) held[capid] = true;
    }
  } catch (e) {
    // Fail closed. If the sheet is unreadable we cannot tell who is mid-flight,
    // and deleting a mailbox we should have held is unrecoverable, whereas
    // skipping a deletion costs one suspended seat until the next run.
    Logger.error('Unable to read Transitions sheet — holding all deletions this run', {
      errorMessage: e && e.message ? e.message : String(e)
    });
    throw e;
  }

  return held;
}

// ============================================================================
// PEER DIRECTORY
// ============================================================================

/**
 * CAPID -> primaryEmail for every account on the PEER tenant, suspended ones
 * INCLUDED.
 *
 * Both roles need this, for mirror-image reasons: the destination (senior)
 * tenant asks "does this member already hold a cadet account?" to exempt them
 * from the Level I gate, and the source (cadet) tenant asks "has their senior
 * account appeared yet?" to know where to migrate mail to. Same read, so one
 * function, ungated by role.
 *
 * Deliberately NOT xtPeerWorkspaceEmailByCapid_() from CrossTenantContacts.gs,
 * which skips suspended peers. That is right for publishing live addresses into
 * the GAL and wrong here: a transitioning member's cadet account is already
 * suspended by the time the senior tenant looks, so reusing it would drop
 * precisely the people this feature exists for.
 *
 * Throws on failure. Callers decide what a failure means — see
 * getPeerTenantCapids_(), which swallows it to fail closed.
 *
 * @returns {Object<string, string>} CAPID -> primaryEmail (lowercased)
 */
function peerCapidToEmail_() {
  const cfg = getCrossTenantConfig_();
  const token = xtPeerToken_(
    'https://www.googleapis.com/auth/admin.directory.user.readonly', cfg
  );

  const map = {};
  let pageToken = '';

  do {
    const url = 'https://admin.googleapis.com/admin/directory/v1/users' +
      '?customer=my_customer&maxResults=500&projection=full' +
      '&fields=nextPageToken,users(primaryEmail,externalIds)' +
      (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    const code = resp.getResponseCode();
    if (code < 200 || code >= 300) {
      throw new Error(`Peer directory list failed (${code}): ${resp.getContentText()}`);
    }

    const body = JSON.parse(resp.getContentText() || '{}');
    (body.users || []).forEach(u => {
      const capid = capidOfUser_(u);
      const email = String(u.primaryEmail || '').trim().toLowerCase();
      if (capid && email) map[capid] = email;
    });

    pageToken = body.nextPageToken || '';
  } while (pageToken);

  Logger.info('Peer directory loaded', {
    count: Object.keys(map).length,
    peerDomain: cfg.peerDomain
  });
  return map;
}

/**
 * CAPIDs holding a peer-tenant account. Destination role only.
 *
 * Used by the Level I gate in updateAllMembers() to recognize a transitioning
 * ex-cadet. Returns {} rather than throwing on any failure — a tenant with no
 * peer configured simply has no ex-cadets, and a member update should not die
 * over it.
 *
 * @returns {Object<string, boolean>} CAPID -> true
 */
function getPeerTenantCapids_() {
  if (TRANSITION_CONFIG.ROLE !== 'destination') return {};

  try {
    const cfg = getCrossTenantConfig_();
    if (!cfg.runInbound) return {};

    const capids = {};
    const byCapid = peerCapidToEmail_();
    for (const capid in byCapid) capids[capid] = true;
    return capids;

  } catch (e) {
    // Fail closed: without the peer list we cannot prove anyone is an ex-cadet,
    // so nobody is exempted and the Level I gate holds. That withholds an
    // account that should have been created — visible, and self-corrects on the
    // next run — rather than provisioning senior accounts for members who never
    // completed Level I.
    Logger.error('Peer directory read failed — no Level I exemptions this run', {
      errorMessage: e && e.message ? e.message : String(e)
    });
    return {};
  }
}

/**
 * Fills in SeniorEmail on PENDING rows whose destination account now exists.
 *
 * Separated from detection because the two answer different questions on
 * different schedules: detection notices the CAPWATCH type flip immediately,
 * while the senior account may not appear for days — the member has to be picked
 * up by the senior tenant's own updateAllMembers() run first. This is the poll
 * that closes that gap, and migration only considers rows it has resolved.
 *
 * A PATRON row stays unresolved on purpose: patrons get no senior account, so
 * there is nowhere to migrate to unless and until they convert.
 *
 * @returns {{resolved: number, stillPending: number}}
 */
function resolveTransitionDestinations() {
  if (TRANSITION_CONFIG.ROLE !== 'source') {
    Logger.info('Destination resolution skipped — not the source tenant');
    return { resolved: 0, stillPending: 0 };
  }
  return withTransitionLock_(resolveTransitionDestinations_, { resolved: 0, stillPending: 0 });
}

function resolveTransitionDestinations_() {
  const rows = readTransitions_();
  const peerByCapid = peerCapidToEmail_();
  let resolved = 0;
  let stillPending = 0;

  for (const capid in rows) {
    const row = rows[capid];
    if (row.MigrationStatus !== TRANSITION_CONFIG.STATUS.PENDING) continue;
    if (row.SeniorEmail) continue;

    const peerEmail = peerByCapid[capid];
    if (!peerEmail) {
      stillPending++;
      continue;
    }

    setTransitionField_(row._rowNumber, 'SeniorEmail', peerEmail);
    Logger.info('Transition destination resolved', {
      capid: capid,
      name: row.Name,
      newType: row.NewType,
      seniorEmail: peerEmail
    });
    resolved++;
  }

  Logger.info('Transition destination resolution completed', {
    resolved: resolved,
    stillPending: stillPending
  });
  return { resolved: resolved, stillPending: stillPending };
}

// ============================================================================
// PREVIEW
// ============================================================================

/**
 * Read-only summary of the transition queue. Changes nothing.
 */
function previewCadetTransitions() {
  const rows = readTransitions_();
  const now = new Date();
  const buckets = {};

  console.log('Transition queue — ' + Object.keys(rows).length + ' row(s)');
  console.log('');

  for (const capid in rows) {
    const row = rows[capid];
    buckets[row.MigrationStatus] = (buckets[row.MigrationStatus] || 0) + 1;

    const deleteAfter = row.DeleteAfter ? new Date(row.DeleteAfter) : null;
    const daysLeft = deleteAfter
      ? Math.ceil((deleteAfter - now) / 86400000)
      : null;

    console.log([
      capid,
      row.Name,
      row.NewType,
      row.MigrationStatus,
      row.SeniorEmail || '(no destination yet)',
      daysLeft === null ? '' : daysLeft + 'd until delete'
    ].join(' | '));
  }

  console.log('');
  console.log('By status: ' + JSON.stringify(buckets));
}

// ============================================================================
// TRIGGERS
// ============================================================================

/**
 * Lifecycle handlers that run on a daily schedule. Deliberately excludes the
 * close/delete step — deletion is permanent and stays a human decision.
 */
const TRANSITION_LIFECYCLE_FN_ = 'runCadetTransitionLifecycle';

/** Hour the daily lifecycle pass runs. After the nightly CAPWATCH pull. */
const TRANSITION_LIFECYCLE_HOUR_ = 3;

/**
 * Stop starting new phases past this much wall time, leaving room under the
 * 6-minute hard ceiling for the phase already running to wind down.
 */
const TRANSITION_LIFECYCLE_BUDGET_MS_ = 4 * 60 * 1000;

/**
 * Handlers disarmTransitionTriggers() removes. Includes the per-phase names from
 * the old one-trigger-per-phase scheme so re-arming cleans them up — that scheme
 * blew the 20-trigger-per-script limit.
 */
const TRANSITION_TRIGGER_FUNCTIONS_ = [
  TRANSITION_LIFECYCLE_FN_,
  'detectCadetTransitions',
  'resolveTransitionDestinations',
  'migrateCadetTransitions',
  'migrateAllTransitionDrives',
  'migrateAllTransitionContacts',
  'catchUpTransitionMail',
  'remindPendingTransitionCloses'
];

/**
 * Installs the daily time-driven triggers that run the transition lifecycle
 * hands-off, from detection through migration.
 *
 * ⚠️ RUN THIS AS automation@cawgcadets.org. Apps Script triggers are owned by,
 * and visible only to, the account that creates them, and the completion email's
 * send-as identity is that account. Run it as anyone else and the jobs execute
 * as the wrong identity — the emails fall back to the runner's address, and the
 * triggers won't show up for the automation account.
 *
 * ⚠️ NO close/delete trigger is installed, on purpose. closeCompletedTransitions
 * permanently deletes accounts with no archive and no undo; it stays manual —
 * `closeCompletedTransitions(true)` to review, then `(false)` to act. Same
 * discipline the license-lifecycle reaper landed on: automate the reversible
 * work, keep a human on the irreversible step.
 *
 * Idempotent: clears the lifecycle triggers first, so re-running re-arms cleanly
 * rather than duplicating. Continuation triggers are left alone — they are
 * transient and self-delete.
 *
 * The hours are staggered so each phase feeds the next within a day, and they
 * must sit AFTER the daily CAPWATCH pull, since detection needs a fresh
 * Member.txt. If getCapwatch runs later than ~2 AM, shift these later to match.
 *
 * @returns {{armed: number}}
 */
function armTransitionTriggers() {
  if (TRANSITION_CONFIG.ROLE !== 'source') {
    throw new Error('Transition triggers belong only on the source (cadets) tenant');
  }

  disarmTransitionTriggers();

  // detect -> resolve -> migrate Gmail -> Drive -> Contacts, an hour apart. Each
  // phase only acts on rows the previous one made ready, so a member who does
  // not finish one phase in a day is simply picked up the next day — well inside
  // the 14/90-day windows.
  ScriptApp.newTrigger(TRANSITION_LIFECYCLE_FN_)
    .timeBased().everyDays(1).atHour(TRANSITION_LIFECYCLE_HOUR_).create();
  Logger.info('Transition lifecycle trigger armed', {
    handler: TRANSITION_LIFECYCLE_FN_, atHour: TRANSITION_LIFECYCLE_HOUR_
  });

  console.log('Armed 1 daily trigger (' + TRANSITION_LIFECYCLE_FN_ + ' at ' +
    TRANSITION_LIFECYCLE_HOUR_ + ':00) that runs the whole lifecycle in order.');
  console.log('NO close/delete trigger — the lifecycle only EMAILS when the timer is up;');
  console.log('you still run closeCompletedTransitions(false) by hand.');
  console.log('Owned by whoever ran this — confirm it is automation@cawgcadets.org.');
  return { armed: 1 };
}

/**
 * Runs every lifecycle phase in order, under ONE daily trigger.
 *
 * Apps Script allows only 20 triggers per script, and this project already runs
 * a full core schedule — one trigger per phase (seven of them) exhausted the
 * budget and the arm failed outright. One trigger for one feature is also just
 * the right shape: the phases are strictly sequential and all daily.
 *
 * Each phase is independently time-limited and schedules its own continuation
 * for its own work, so a phase that cannot finish today resumes on its own. The
 * elapsed-time check between phases stops this execution before the 6-minute
 * hard ceiling; whatever is skipped simply runs on tomorrow's pass, because
 * every completed phase returns fast when it has nothing to do.
 *
 * @param {Object} [e] - trigger event (ignored)
 */
function runCadetTransitionLifecycle(e) {
  if (TRANSITION_CONFIG.ROLE !== 'source') {
    Logger.info('Transition lifecycle skipped — not the source tenant');
    return;
  }

  const started = new Date();
  const phases = [
    ['detect',       function () { detectCadetTransitions(); }],
    ['resolve',      function () { resolveTransitionDestinations(); }],
    ['migrate mail', function () { migrateCadetTransitions(); }],
    ['migrate Drive',function () { migrateAllTransitionDrives(); }],
    ['migrate contacts', function () { migrateAllTransitionContacts(); }],
    // Parked accounts stay LIVE and keep receiving, so this sweep is what
    // actually delivers their mail across — the auto-forward is only a
    // best-effort accelerator. Without it a parked mailbox silently piles up.
    ['sweep parked', function () { catchUpTransitionMail(); }],
    ['remind',       function () { remindPendingTransitionCloses(); }]
  ];

  for (let i = 0; i < phases.length; i++) {
    if (new Date() - started > TRANSITION_LIFECYCLE_BUDGET_MS_) {
      Logger.info('Lifecycle budget reached — remaining phases run tomorrow', {
        completed: i, next: phases[i][0]
      });
      break;
    }
    try {
      phases[i][1]();
    } catch (err) {
      // One phase failing must not strand the rest — notably the sweep and the
      // reminder, which are what keep parked accounts working.
      Logger.error('Lifecycle phase failed — continuing', {
        phase: phases[i][0],
        errorMessage: (err && err.message) ? err.message : String(err)
      });
    }
  }

  Logger.info('Transition lifecycle pass finished', {
    duration: new Date() - started + 'ms'
  });
}

/**
 * Removes the lifecycle triggers armTransitionTriggers installed. Leaves the
 * transient continuation triggers alone.
 *
 * @returns {{removed: number}}
 */
function disarmTransitionTriggers() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (TRANSITION_TRIGGER_FUNCTIONS_.indexOf(t.getHandlerFunction()) > -1) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  Logger.info('Transition lifecycle triggers removed', { removed: removed });
  return { removed: removed };
}

// ============================================================================
// ONE-TRIGGER PIPELINE
// ============================================================================

/**
 * The pipeline handler, and where its position is parked between runs.
 *
 * Apps Script allows 20 triggers per script per user, and the six lifecycle
 * phases were six of them. Running them from one driver returns five slots
 * without changing what any phase does.
 */
const TRANSITION_PIPELINE_FN_ = 'runCadetTransitionPipeline';
const TRANSITION_PIPELINE_PHASE_PROP_ = 'TRANSITION_PIPELINE_PHASE';
const TRANSITION_PIPELINE_SAVED_AT_PROP_ = 'TRANSITION_PIPELINE_SAVED_AT';

/**
 * Wall-clock budget for one pipeline run, in minutes. These tenants allow 30;
 * 25 leaves room to park the position and log.
 */
const TRANSITION_PIPELINE_BUDGET_MIN_ = 25;

/**
 * Runs the transition lifecycle phases in order, from one trigger.
 *
 * WHY THIS IS SAFE TO COLLAPSE. The six phases were never an hour apart because
 * each takes an hour — they are a pipeline where every phase acts only on rows
 * the previous one made ready, and the schedule comment has always said a member
 * who misses a phase today is picked up tomorrow, well inside the 14/90-day
 * windows. The three expensive phases (Gmail, Drive, Contacts) additionally
 * self-limit and schedule their OWN continuation triggers a minute out, so none
 * of them needs a private 30-minute window to make progress.
 *
 * THREE BEHAVIOURS OF THE SIX-TRIGGER SETUP THAT ARE PRESERVED DELIBERATELY:
 *
 *   1. A failing phase does not stop the others. Six separate triggers meant a
 *      3 a.m. exception never prevented the 4 a.m. run. Each phase here is
 *      caught and logged, and the pipeline carries on — aborting the rest would
 *      quietly make the system less resilient than the thing it replaced.
 *
 *   2. Order is kept. detect → resolve → Gmail → Drive → Contacts → remind.
 *
 *   3. Continuation triggers are untouched. They belong to the phases, not to
 *      this driver, and are how the heavy work actually drains.
 *
 * AND ONE THAT IS NEW. The run parks the phase it stopped at and resumes there
 * next time, wrapping when the cycle completes. Without that, a day where the
 * budget ran out mid-list would run the same early phases and never reach the
 * late ones — which is exactly how nine squadrons went unvisited for weeks in
 * the group sync. A phase list is short enough that it would have taken far
 * longer to notice.
 *
 * @param {number} [budgetMinutes=25] - Wall-clock budget for THIS execution
 * @returns {Object} Summary of the run
 */
function runCadetTransitionPipeline(budgetMinutes) {
  const summary = { ran: [], failed: [], skipped: [], complete: false, startedAtPhase: 0 };

  if (TRANSITION_CONFIG.ROLE !== 'source') {
    Logger.info('Transition pipeline skipped — not the source tenant', {
      role: TRANSITION_CONFIG.ROLE || '(off)'
    });
    return summary;
  }

  // Each phase is dispatched by name so the parked position survives a deploy:
  // an index into a list of names means the same thing after a push, where a
  // parked function reference would not.
  const phases = {
    detectCadetTransitions: detectCadetTransitions,
    resolveTransitionDestinations: resolveTransitionDestinations,
    // Called with no argument, exactly as a time-driven trigger effectively
    // does: migrateCadetTransitions coerces anything that is not false to
    // "notify on", so this preserves the notification behaviour it had.
    migrateCadetTransitions: migrateCadetTransitions,
    migrateAllTransitionDrives: migrateAllTransitionDrives,
    migrateAllTransitionContacts: migrateAllTransitionContacts,
    remindPendingTransitionCloses: remindPendingTransitionCloses
  };

  const order = TRANSITION_TRIGGER_FUNCTIONS_;
  const budgetMs = Math.max(1, Number(budgetMinutes || TRANSITION_PIPELINE_BUDGET_MIN_)) * 60 * 1000;
  const deadline = Date.now() + budgetMs;

  let index = loadTransitionPipelinePhase_(order.length);
  summary.startedAtPhase = index;

  if (index > 0) {
    Logger.info('Resuming transition pipeline', {
      atPhase: index,
      phase: order[index],
      note: 'A previous run stopped here on time.'
    });
  }

  for (; index < order.length; index++) {
    const name = order[index];

    if (Date.now() >= deadline) {
      Logger.warn('Transition pipeline out of time — parking position', {
        stoppedBefore: name,
        phase: index,
        remaining: order.length - index
      });
      summary.skipped = order.slice(index);
      break;
    }

    try {
      phases[name]();
      summary.ran.push(name);
    } catch (e) {
      // Deliberately not rethrown: see behaviour (1) above.
      Logger.error('Transition phase failed — continuing with the next', {
        phase: name,
        errorMessage: e.message
      });
      summary.failed.push({ phase: name, errorMessage: e.message });
    }
  }

  summary.complete = index >= order.length;
  saveTransitionPipelinePhase_(summary.complete ? 0 : index);

  Logger.info('Transition pipeline run complete', {
    ran: summary.ran.length,
    failed: summary.failed.length,
    skipped: summary.skipped.length,
    complete: summary.complete,
    nextPhase: summary.complete ? 0 : index
  });

  return summary;
}

/**
 * @param {number} phaseCount - Length of the phase list, to reject stale indices
 * @returns {number} Phase to start from
 */
function loadTransitionPipelinePhase_(phaseCount) {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty(TRANSITION_PIPELINE_PHASE_PROP_);
    const index = parseInt(raw, 10);
    if (!isFinite(index) || index <= 0) return 0;

    // A position parked before the phase list changed shape means nothing.
    // Starting over re-runs phases that are idempotent; resuming into the wrong
    // one skips work silently.
    if (index >= phaseCount) {
      Logger.warn('Parked transition phase is past the end of the list — starting over', {
        parked: index,
        phases: phaseCount
      });
      return 0;
    }

    return index;
  } catch (e) {
    Logger.warn('Could not read the parked transition phase; starting over', {
      errorMessage: e.message
    });
    return 0;
  }
}

/**
 * @param {number} index - Phase to resume at; 0 when the cycle completed
 * @returns {void}
 */
function saveTransitionPipelinePhase_(index) {
  try {
    PropertiesService.getScriptProperties().setProperties({
      [TRANSITION_PIPELINE_PHASE_PROP_]: String(index),
      [TRANSITION_PIPELINE_SAVED_AT_PROP_]: new Date().toISOString()
    });
  } catch (e) {
    Logger.warn('Could not park the transition pipeline position', { errorMessage: e.message });
  }
}

/**
 * Installs ONE daily trigger for the whole lifecycle, replacing the six.
 *
 * Deletes before it creates, so the swap needs no free slot on a project at the
 * 20-trigger ceiling — which is the reason to do this at all. Continuation
 * triggers are left alone: they belong to the phases.
 *
 * ⚠ RUN THIS AS THE AUTOMATION ACCOUNT. Triggers are owned by whoever creates
 * them, and the completion mail sends under the automation account's Send-As.
 *
 * @param {number} [hour=3] - Hour of day to run, 0-23, script timezone
 * @returns {{removed: number, installed: string, hour: number}}
 */
function armTransitionPipelineTrigger(hour) {
  if (TRANSITION_CONFIG.ROLE !== 'source') {
    throw new Error('Transition triggers belong only on the source (cadets) tenant');
  }

  const atHour = (hour === undefined || hour === null) ? 3 : Number(hour);
  if (!isFinite(atHour) || atHour < 0 || atHour > 23) {
    throw new Error('hour must be 0-23; got ' + hour);
  }

  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const handler = t.getHandlerFunction();
    if (TRANSITION_TRIGGER_FUNCTIONS_.indexOf(handler) > -1 ||
        handler === TRANSITION_PIPELINE_FN_) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });

  try {
    ScriptApp.newTrigger(TRANSITION_PIPELINE_FN_)
      .timeBased().everyDays(1).atHour(atHour).create();
  } catch (e) {
    Logger.error('Could not install the transition pipeline trigger', {
      errorMessage: e.message,
      removed: removed,
      warning: 'The transition lifecycle now has NO trigger. Free a slot and re-run this.'
    });
    throw e;
  }

  Logger.info('Armed the single transition pipeline trigger', {
    handler: TRANSITION_PIPELINE_FN_,
    atHour: atHour,
    removed: removed,
    note: 'Confirm in the Triggers panel that the owner is the automation account'
  });

  console.log('✅ One daily trigger now runs the whole lifecycle (replaced ' + removed + ').');
  console.log('Freed ' + Math.max(0, removed - 1) + ' trigger slot(s).');
  console.log('Close/delete is still NOT automated — closeCompletedTransitions(false) stays manual.');

  return { removed: removed, installed: TRANSITION_PIPELINE_FN_, hour: atHour };
}

/**
 * Read-only: which phase the next pipeline run will start from.
 * @returns {void}
 */
function checkTransitionPipelineStatus() {
  const props = PropertiesService.getScriptProperties();
  const index = parseInt(props.getProperty(TRANSITION_PIPELINE_PHASE_PROP_), 10) || 0;
  const savedAt = props.getProperty(TRANSITION_PIPELINE_SAVED_AT_PROP_) || '(never)';
  console.log('Next run starts at phase ' + index + ' of ' + TRANSITION_TRIGGER_FUNCTIONS_.length +
    ' (' + TRANSITION_TRIGGER_FUNCTIONS_[index] + ')');
  console.log('Position last written: ' + savedAt);
  if (index === 0) console.log('0 means the last run completed the full cycle.');
}

/**
 * Lists EVERY trigger on the project, not just this feature's, with a count
 * against Apps Script's hard limit of 20 per script.
 *
 * Worth having because that limit is shared across the whole project: this
 * feature adding one trigger per phase exhausted it and the arm failed with
 * "This script has too many triggers." Orphaned continuation triggers (which
 * normally self-delete) also consume slots, so this separates scheduled work
 * from leftovers.
 */
function listAllProjectTriggers() {
  const all = ScriptApp.getProjectTriggers();
  const continuations = ['continueCadetTransitionMigration',
                         'continueCadetTransitionDriveMigration',
                         'continueCadetTransitionContactsMigration'];
  let leftovers = 0;

  console.log('=== ' + all.length + ' of 20 trigger slots used ===');
  all.forEach(function (t) {
    const fn = t.getHandlerFunction();
    const isCont = continuations.indexOf(fn) > -1;
    if (isCont) leftovers++;
    console.log('  ' + fn + (isCont ? '   <-- continuation (transient; stale ones are removable)' : ''));
  });

  console.log('');
  if (leftovers) {
    console.log(leftovers + ' continuation trigger(s) present. If no migration is running,');
    console.log('they are orphans — clear with clearAllTransitionContinuations().');
  }
  if (all.length >= 18) {
    console.log('WARNING: at or near the 20-trigger ceiling. Adding more will fail.');
  }
  return { used: all.length, continuations: leftovers };
}

/** Clears every stale transition continuation trigger (all three kinds). */
function clearAllTransitionContinuations() {
  const a = clearMigrationContinuations();
  const b = clearDriveContinuations();
  const c = clearContactsContinuations();
  const total = (a.removed || 0) + (b.removed || 0) + (c.removed || 0);
  console.log('Cleared ' + total + ' continuation trigger(s).');
  return { removed: total };
}

/**
 * Read-only: lists the transition triggers currently installed and who would
 * own them. Run to confirm state after arming.
 */
function listTransitionTriggers() {
  // Includes the consolidated driver: filtering on the six phase handlers alone
  // would report "none installed" on a project running the pipeline, which is
  // the opposite of true.
  const mine = ScriptApp.getProjectTriggers().filter(function (t) {
    const handler = t.getHandlerFunction();
    return TRANSITION_TRIGGER_FUNCTIONS_.indexOf(handler) > -1 ||
      handler === TRANSITION_PIPELINE_FN_;
  });
  if (!mine.length) {
    console.log('No transition lifecycle triggers installed.');
    return;
  }
  mine.forEach(function (t) {
    console.log(t.getHandlerFunction() + '  (' + t.getEventType() + ')');
  });
  console.log('');
  console.log(mine.length + ' installed. Continuation triggers (transient) are not listed.');
}
