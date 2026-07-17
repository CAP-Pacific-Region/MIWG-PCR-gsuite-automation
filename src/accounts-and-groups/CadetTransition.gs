/**
 * Cadet → senior tenant transition.
 *
 * Version: 1.0.0
 * Date: 2026-07-16
 * Changes: 1.0.0 — initial release. Detection + state for the cadet→senior
 *   lifecycle; the Transitions sheet is authoritative for who is mid-flight.
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
const TRANSITION_TRIGGER_FUNCTIONS_ = [
  'detectCadetTransitions',
  'resolveTransitionDestinations',
  'migrateCadetTransitions',
  'migrateAllTransitionDrives',
  'migrateAllTransitionContacts',
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
  const schedule = [
    ['detectCadetTransitions', 3],
    ['resolveTransitionDestinations', 4],
    ['migrateCadetTransitions', 5],
    ['migrateAllTransitionDrives', 6],
    ['migrateAllTransitionContacts', 7],
    // After migration, so the ready/stuck picture reflects the day's work.
    // Deletion is NOT automated; this only emails IT that the timer is up.
    ['remindPendingTransitionCloses', 8]
  ];

  schedule.forEach(function (pair) {
    ScriptApp.newTrigger(pair[0]).timeBased().everyDays(1).atHour(pair[1]).create();
    Logger.info('Transition trigger armed', { handler: pair[0], atHour: pair[1] });
  });

  console.log('Armed ' + schedule.length + ' daily transition triggers (detect → migrate → remind).');
  console.log('NO close/delete trigger — remindPendingTransitionCloses just EMAILS when the');
  console.log('timer is up; you still run closeCompletedTransitions(false) by hand.');
  console.log('These are owned by whoever ran this — confirm it is automation@cawgcadets.org.');
  return { armed: schedule.length };
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

/**
 * Read-only: lists the transition triggers currently installed and who would
 * own them. Run to confirm state after arming.
 */
function listTransitionTriggers() {
  const mine = ScriptApp.getProjectTriggers().filter(function (t) {
    return TRANSITION_TRIGGER_FUNCTIONS_.indexOf(t.getHandlerFunction()) > -1;
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
