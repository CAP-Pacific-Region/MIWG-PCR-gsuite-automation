/**
 * Cadet → senior transition: Drive.
 *
 * Runs on the CADETS tenant (TRANSITION_CONFIG.ROLE === 'source'). See
 * CadetTransition.gs for the lifecycle this belongs to.
 *
 * WHY COPY RATHER THAN TRANSFER. Deleting a Workspace user destroys their Drive.
 * The Admin console's "transfer files on delete" only works within a domain --
 * Google does not transfer ownership across tenants, and the Data Transfer API
 * is same-tenant only. Sharing does not save anything either: a shared file
 * still dies with its owner. Copying is the only mechanism that survives the
 * deletion, so it must happen BEFORE closeCompletedTransitions() runs.
 *
 * HOW. Share, then server-side copy:
 *   1. impersonate the cadet, grant the senior reader on the file
 *   2. impersonate the senior, files.copy into their own Drive
 * Google moves the bytes internally. The alternative -- download through
 * UrlFetchApp and re-upload -- would push every byte through a 6-minute
 * execution against a 50MB response cap, and would choke on a single large
 * video.
 *
 * SCOPES (SA domain-wide delegation grants, not appsscript.json):
 *   cadets local SA:  https://www.googleapis.com/auth/drive
 *     — full drive, not drive.readonly: granting a permission is a WRITE to the
 *       file's ACL, which drive.readonly forbids.
 *   seniors peer SA:  https://www.googleapis.com/auth/drive
 *     — files.copy writes into the member's Drive.
 *
 * Only files the cadet OWNS are copied. Files merely shared with them belong to
 * someone else, survive the deletion untouched, and copying them would both
 * duplicate other people's documents and silently reassign ownership to the
 * member.
 */

/** Root folder created in the destination Drive. */
const DRIVE_MIGRATE_FOLDER_ = 'Migrated from cadets';

/** Files per page when listing the source Drive. */
const DRIVE_PAGE_SIZE_ = 50;

/** Google Docs-native types cannot be copied with a byte range; they copy whole. */
const DRIVE_FOLDER_MIME_ = 'application/vnd.google-apps.folder';

/** Shortcuts point at other files; copying one copies nothing useful. */
const DRIVE_SKIP_MIMES_ = [
  'application/vnd.google-apps.shortcut',
  'application/vnd.google-apps.fusiontable',
  'application/vnd.google-apps.map'
];

/** Full Drive: needed to write ACLs on the source and create/copy on the dest. */
const DRIVE_SCOPE_ = 'https://www.googleapis.com/auth/drive';

/**
 * appProperties keys stamped on every item this creates in the destination.
 *
 * These ARE the resume mechanism. The destination records what has already been
 * copied, so a run starts by asking the destination "what did earlier runs
 * finish?" rather than trusting a cursor. A lost or stale DriveCursor therefore
 * cannot cause a double-copy — the worst case is re-listing the source and
 * skipping everything already marked. Idempotency lives in the data, not the
 * bookkeeping.
 *
 *   xtMember — the CAPID, so one query returns everything done for this member
 *   xtSrcId  — the source file/folder id this was copied from
 *   xtType   — 'folder' | 'file'
 */
const DRIVE_MARK_MEMBER_ = 'xtMember';
const DRIVE_MARK_SRC_ = 'xtSrcId';
const DRIVE_MARK_TYPE_ = 'xtType';

/** Continuation trigger + its scope property. Separate from the Gmail pair. */
const DRIVE_CONTINUATION_FN_ = 'continueCadetTransitionDriveMigration';
const DRIVE_CONTINUATION_SCOPE_PROP_ = 'DRIVE_CONTINUATION_SCOPE';

// ============================================================================
// ENTRY POINTS
// ============================================================================

/**
 * Copies one member's Drive to their senior account. Resumable, idempotent.
 *
 * @param {string} capid
 * @returns {{folders: number, files: number, skipped: number, complete: boolean}}
 */
function migrateSingleTransitionDrive(capid) {
  if (TRANSITION_CONFIG.ROLE !== 'source') throw new Error('Not the source tenant');

  const rows = readTransitions_();
  const row = rows[String(capid)];
  if (!row) throw new Error('No transition row for CAPID ' + capid);
  if (!row.SeniorEmail) throw new Error('No destination account for CAPID ' + capid);

  return migrateOneDrive_(row, new Date());
}

/**
 * Proves both Drive credentials on ONE member without copying anything.
 *
 * Nothing has copied a real file yet, and the cheapest failure to discover is a
 * missing DWD scope. This mints the source token (needed to grant ACLs) and the
 * destination token impersonating the member (needed to create/copy), and does a
 * harmless read with each, so a scope gap surfaces here rather than partway
 * through 4787 files.
 *
 * @param {string} capid
 */
function previewSingleTransitionDrive(capid) {
  const rows = readTransitions_();
  const row = rows[String(capid)];
  if (!row) { console.log('No row for ' + capid); return; }
  if (!row.SeniorEmail) { console.log(row.Name + ': no destination account'); return; }

  const cfg = getCrossTenantConfig_();

  try {
    const srcToken = getImpersonatedToken_(row.CadetEmail, DRIVE_SCOPE_);
    const s = driveFetch_('https://www.googleapis.com/drive/v3/about?fields=user',
      { method: 'get', token: srcToken }, 'source about');
    console.log('source OK  — impersonated ' + (s.user && s.user.emailAddress));

    const destToken = xtPeerToken_(DRIVE_SCOPE_, cfg, row.SeniorEmail);
    const d = driveFetch_('https://www.googleapis.com/drive/v3/about?fields=user,storageQuota',
      { method: 'get', token: destToken }, 'dest about');
    console.log('dest OK    — impersonated ' + (d.user && d.user.emailAddress));

    // Confirm the member address matches the token, same guard as the Gmail path.
    if (String(d.user.emailAddress).toLowerCase() !== String(row.SeniorEmail).toLowerCase()) {
      console.log('  WARNING: dest token opened ' + d.user.emailAddress + ', not ' + row.SeniorEmail);
    }

    const already = loadDriveProgress_(row.CAPID, destToken);
    console.log('already done: ' + already.folderMap.size + ' folders, ' +
      already.fileCount + ' files' + (already.rootId ? ' (resumable)' : ' (fresh)'));
    console.log('Both credentials work. Nothing was copied.');

  } catch (e) {
    console.log('FAILED: ' + (e && e.message ? e.message : String(e)));
  }
}

/**
 * Copies Drive for every migrated member who has not had it done yet.
 *
 * Gated on MigrationStatus COMPLETE and an empty DriveMigrated, so it only
 * touches members whose mail is already across. One member per execution to keep
 * the picture legible; the continuation carries on.
 *
 * @returns {{processed: number}}
 */
function migrateAllTransitionDrives() {
  if (TRANSITION_CONFIG.ROLE !== 'source') {
    Logger.info('Drive migration skipped — not the source tenant');
    return { processed: 0 };
  }

  const started = new Date();
  const rows = readTransitions_();
  let processed = 0;

  for (const capid in rows) {
    const row = rows[capid];
    if (row.MigrationStatus !== TRANSITION_CONFIG.STATUS.COMPLETE) continue;
    if (String(row.DriveMigrated || '') !== '') continue;   // already handled (0 counts)
    if (!row.SeniorEmail) continue;

    const result = migrateOneDrive_(row, started);
    processed++;

    if (!result.complete) {
      // Out of time on this member. migrateOneDrive_ has already scheduled the
      // continuation scoped to it; just stop starting new members this run.
      break;
    }
  }

  Logger.info('Drive migration pass completed', {
    duration: new Date() - started + 'ms', processed: processed
  });
  return { processed: processed };
}

/** Continuation trigger target. Deletes its own trigger, then resumes. */
function continueCadetTransitionDriveMigration(e) {
  if (e && e.triggerUid) {
    ScriptApp.getProjectTriggers()
      .filter(t => t.getUniqueId() === e.triggerUid)
      .forEach(t => ScriptApp.deleteTrigger(t));
  }

  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(DRIVE_CONTINUATION_SCOPE_PROP_);
  props.deleteProperty(DRIVE_CONTINUATION_SCOPE_PROP_);

  let scope = null;
  try { scope = raw ? JSON.parse(raw) : null; } catch (err) {}

  if (scope && scope.capid) {
    const result = migrateSingleTransitionDrive(scope.capid);
    // If that member finished, look for others — but in a fresh execution, not
    // this one. Starting another member's 4.5-min budget on top of the time
    // already spent here could blow the 6-minute hard cap. If it paused instead,
    // migrateOneDrive_ already scheduled its own continuation; do not stack a
    // second. This is what lets one kickoff drain the whole queue rather than
    // stopping after the member it was scoped to.
    if (result && result.complete) scheduleDriveContinuation_({});
  } else {
    migrateAllTransitionDrives();
  }
}

// ============================================================================
// PER-MEMBER COPY
// ============================================================================

/**
 * Copies one member's owned Drive: folders first (structure), then files.
 *
 * Two phases because a file cannot be copied into a folder that does not exist
 * yet, and folders cannot be server-side copied at all — files.copy refuses a
 * folder, so the tree is recreated with files.create and the files are copied
 * into the recreated folders.
 *
 * @param {Object} row - Transitions row
 * @param {Date} started - pass start, for the shared time budget
 * @returns {{folders: number, files: number, skipped: number, complete: boolean}}
 */
function migrateOneDrive_(row, started) {
  const cfg = getCrossTenantConfig_();
  const srcToken = getImpersonatedToken_(row.CadetEmail, DRIVE_SCOPE_);
  const destToken = xtPeerToken_(DRIVE_SCOPE_, cfg, row.SeniorEmail);

  Logger.info('Drive migration starting', {
    capid: row.CAPID, from: row.CadetEmail, to: row.SeniorEmail
  });

  // What earlier runs already finished — read from the destination itself, so
  // this is correct even if DriveCursor was lost.
  const done = loadDriveProgress_(row.CAPID, destToken);
  const rootId = done.rootId || createMigrationRoot_(row, destToken);

  // Phase 1: folders. Cheap relative to files and needed before any file copy,
  // so always completed before moving on — it holds the whole tree in memory
  // (724 folders for the largest member, trivially).
  const folderResult = replicateFolders_(row, srcToken, destToken, rootId, done, started);
  if (!folderResult.complete) {
    setTransitionField_(row._rowNumber, 'Notes',
      `Drive: folder structure in progress (${done.folderMap.size} folders)`);
    // Reschedule from HERE, not from the caller: this per-member function is what
    // every entry path funnels through (single, all, and each continuation), so
    // owning the continuation here is the only place that keeps a paused run
    // alive regardless of how it was started.
    scheduleDriveContinuation_({ capid: String(row.CAPID) });
    return { folders: done.folderMap.size, files: done.fileCount, skipped: 0, complete: false };
  }

  // Phase 2: files, paginated and time-bounded.
  const fileResult = copyFiles_(row, srcToken, destToken, rootId, done, started);

  if (fileResult.complete) {
    setTransitionField_(row._rowNumber, 'DriveMigrated', done.fileCount + fileResult.copied);
    setTransitionField_(row._rowNumber, 'DriveCursor', '');
    if (fileResult.skipped.length) {
      setTransitionField_(row._rowNumber, 'Notes',
        `Drive copied, but SKIPPED ${fileResult.skipped.length} item(s) — DO NOT DELETE ` +
        `until handled: ` + fileResult.skipped.join(' | '));
      Logger.error('Drive migration finished with skips — deletion blocked', {
        capid: row.CAPID, skipped: fileResult.skipped.length
      });
    } else {
      Logger.info('Drive migration complete', {
        capid: row.CAPID,
        folders: done.folderMap.size,
        files: done.fileCount + fileResult.copied
      });
    }
  } else {
    setTransitionField_(row._rowNumber, 'DriveCursor', fileResult.cursor);
    scheduleDriveContinuation_({ capid: String(row.CAPID) });
  }

  return {
    folders: done.folderMap.size,
    files: done.fileCount + fileResult.copied,
    skipped: fileResult.skipped ? fileResult.skipped.length : 0,
    complete: fileResult.complete
  };
}

/**
 * Reads what earlier runs finished, straight from the destination.
 *
 * One paged query for every item tagged with this member's CAPID returns the
 * root, the source→dest folder map, the set of source ids already copied, and
 * the running file count. This is why a lost cursor cannot double-copy: the
 * destination is the source of truth for "done".
 *
 * @param {string} capid
 * @param {string} destToken
 * @returns {{rootId: string, folderMap: Map, doneIds: Set, fileCount: number}}
 */
function loadDriveProgress_(capid, destToken) {
  const folderMap = new Map();   // source folder id -> dest folder id
  const doneIds = new Set();     // source ids already copied (files and folders)
  let rootId = '';
  let fileCount = 0;
  let pageToken = '';

  const q = `appProperties has { key='${DRIVE_MARK_MEMBER_}' and value='${capid}' } and trashed=false`;

  do {
    const url = 'https://www.googleapis.com/drive/v3/files' +
      '?q=' + encodeURIComponent(q) +
      '&pageSize=1000&fields=nextPageToken,files(id,mimeType,appProperties)' +
      (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

    const body = driveFetch_(url, { method: 'get', token: destToken }, 'load drive progress ' + capid);

    (body.files || []).forEach(f => {
      const ap = f.appProperties || {};
      const srcId = ap[DRIVE_MARK_SRC_];
      if (srcId === 'ROOT') { rootId = f.id; return; }
      if (!srcId) return;

      doneIds.add(srcId);
      if (ap[DRIVE_MARK_TYPE_] === 'folder') {
        folderMap.set(srcId, f.id);
      } else {
        fileCount++;
      }
    });

    pageToken = body.nextPageToken || '';
  } while (pageToken);

  Logger.info('Drive progress loaded', {
    capid: capid, resuming: !!rootId, folders: folderMap.size, files: fileCount
  });
  return { rootId: rootId, folderMap: folderMap, doneIds: doneIds, fileCount: fileCount };
}

/**
 * Creates the destination root folder, marked so future runs recognise it.
 *
 * @param {Object} row
 * @param {string} destToken
 * @returns {string} root folder id
 */
function createMigrationRoot_(row, destToken) {
  const body = driveFetch_(
    'https://www.googleapis.com/drive/v3/files?fields=id',
    {
      method: 'post', token: destToken,
      payload: JSON.stringify({
        name: DRIVE_MIGRATE_FOLDER_,
        mimeType: DRIVE_FOLDER_MIME_,
        appProperties: driveMark_(row.CAPID, 'ROOT', 'folder')
      })
    },
    'create migration root for ' + row.SeniorEmail
  );
  Logger.info('Drive migration root created', { capid: row.CAPID, rootId: body.id });
  return body.id;
}

/** appProperties payload marking a created item. */
function driveMark_(capid, srcId, type) {
  const m = {};
  m[DRIVE_MARK_MEMBER_] = String(capid);
  m[DRIVE_MARK_SRC_] = String(srcId);
  m[DRIVE_MARK_TYPE_] = type;
  return m;
}

// ============================================================================
// PHASE 1 — FOLDERS
// ============================================================================

/**
 * Recreates the owned folder tree in the destination, parents before children.
 *
 * Folders are created a level at a time: each pass creates every folder whose
 * parent already exists in the map (or whose parent is not an owned folder, in
 * which case it hangs off the migration root). A tree resolves in as many passes
 * as it is deep, which is small.
 *
 * @returns {{complete: boolean}}
 */
function replicateFolders_(row, srcToken, destToken, rootId, done, started) {
  const folders = listOwnedFolders_(row.CadetEmail, srcToken);   // all of them, once
  const ownedIds = new Set(folders.map(f => f.id));

  // Folders still needing creation.
  let remaining = folders.filter(f => !done.doneIds.has(f.id));

  while (remaining.length) {
    if (new Date() - started > MIGRATE_SOFT_TIME_LIMIT_MS_) {
      Logger.info('Folder replication paused — will resume', {
        capid: row.CAPID, created: done.folderMap.size, remaining: remaining.length
      });
      return { complete: false };
    }

    // Folders whose parent is ready this pass.
    const ready = remaining.filter(f => {
      const parent = (f.parents && f.parents[0]) || '';
      return !ownedIds.has(parent) || done.folderMap.has(parent);
    });

    if (!ready.length) {
      // No parent resolved — a cycle or an orphan. Attach the rest to root
      // rather than loop forever; structure is best-effort, losing files is not.
      Logger.warn('Folder tree did not fully resolve — attaching remainder to root', {
        capid: row.CAPID, orphans: remaining.length
      });
      ready.push.apply(ready, remaining);
    }

    for (const group of chunk_(ready, MIGRATE_PARALLEL_)) {
      const results = fetchAllWithRetry_(
        group.map(f => {
          const parent = (f.parents && f.parents[0]) || '';
          const destParent = done.folderMap.get(parent) || rootId;
          return {
            url: 'https://www.googleapis.com/drive/v3/files?fields=id',
            method: 'post',
            contentType: 'application/json',
            headers: { Authorization: 'Bearer ' + destToken },
            payload: JSON.stringify({
              name: f.name,
              mimeType: DRIVE_FOLDER_MIME_,
              parents: [destParent],
              appProperties: driveMark_(row.CAPID, f.id, 'folder')
            }),
            muteHttpExceptions: true
          };
        }),
        'create folder'
      );

      results.forEach((res, i) => {
        if (!res.ok) throw new Error(res.error);
        done.folderMap.set(group[i].id, res.body.id);
        done.doneIds.add(group[i].id);
      });
    }

    remaining = remaining.filter(f => !done.doneIds.has(f.id));
  }

  return { complete: true };
}

/**
 * Every folder the cadet owns: id, name, first parent.
 *
 * @param {string} cadetEmail
 * @param {string} srcToken
 * @returns {Array<{id: string, name: string, parents: Array<string>}>}
 */
function listOwnedFolders_(cadetEmail, srcToken) {
  const q = "'me' in owners and trashed = false and mimeType = '" + DRIVE_FOLDER_MIME_ + "'";
  const folders = [];
  let pageToken = '';

  do {
    const url = 'https://www.googleapis.com/drive/v3/files' +
      '?q=' + encodeURIComponent(q) +
      '&pageSize=1000&fields=nextPageToken,files(id,name,parents)' +
      (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

    const body = driveFetch_(url, { method: 'get', token: srcToken }, 'list folders for ' + cadetEmail);
    (body.files || []).forEach(f => folders.push(f));
    pageToken = body.nextPageToken || '';
  } while (pageToken);

  return folders;
}

// ============================================================================
// PHASE 2 — FILES
// ============================================================================

/**
 * Copies owned non-folder files into their recreated parents.
 *
 * Each file is a grant (impersonating the cadet, adding the senior as reader)
 * then a server-side copy (impersonating the senior). Both are batched: all
 * grants in a chunk go in parallel, then all copies. Bytes never touch Apps
 * Script — Google copies internally — so a 1.9GB video is no different from a
 * 10KB doc here.
 *
 * @returns {{complete: boolean, copied: number, skipped: Array<string>, cursor: string}}
 */
function copyFiles_(row, srcToken, destToken, rootId, done, started) {
  let pageToken = row.DriveCursor || '';
  let copied = 0;
  const skipped = [];

  const q = "'me' in owners and trashed = false and mimeType != '" + DRIVE_FOLDER_MIME_ + "'";

  do {
    if (new Date() - started > MIGRATE_SOFT_TIME_LIMIT_MS_) {
      return { complete: false, copied: copied, skipped: skipped, cursor: pageToken };
    }

    const listUrl = 'https://www.googleapis.com/drive/v3/files' +
      '?q=' + encodeURIComponent(q) +
      '&pageSize=' + DRIVE_PAGE_SIZE_ +
      '&fields=nextPageToken,files(id,name,mimeType,parents)' +
      (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

    const page = driveFetch_(listUrl, { method: 'get', token: srcToken }, 'list files for ' + row.CadetEmail);

    // Only files not already copied and not on the skip list of unhelpful types.
    const todo = (page.files || []).filter(f =>
      !done.doneIds.has(f.id) && DRIVE_SKIP_MIMES_.indexOf(f.mimeType) < 0
    );

    for (const group of chunk_(todo, MIGRATE_PARALLEL_)) {
      // Grant the senior read on each source file.
      const grants = fetchAllWithRetry_(
        group.map(f => ({
          url: 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(f.id) +
            '/permissions?fields=id&sendNotificationEmail=false',
          method: 'post',
          contentType: 'application/json',
          headers: { Authorization: 'Bearer ' + srcToken },
          payload: JSON.stringify({ role: 'reader', type: 'user', emailAddress: row.SeniorEmail }),
          muteHttpExceptions: true
        })),
        'grant read'
      );

      // A file we cannot share cannot be copied. Skip it loudly — it will not
      // survive the deletion — rather than failing the whole run.
      const copyable = [];
      grants.forEach((res, i) => {
        if (res.ok) { copyable.push(group[i]); return; }
        skipped.push(`"${group[i].name}" (${group[i].id}): grant failed`);
        Logger.error('Drive file could not be shared — SKIPPED', {
          capid: row.CAPID, fileId: group[i].id, name: group[i].name, error: res.error
        });
      });

      if (!copyable.length) continue;

      const copies = fetchAllWithRetry_(
        copyable.map(f => {
          const parent = (f.parents && f.parents[0]) || '';
          const destParent = done.folderMap.get(parent) || rootId;
          return {
            url: 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(f.id) +
              '/copy?fields=id',
            method: 'post',
            contentType: 'application/json',
            headers: { Authorization: 'Bearer ' + destToken },
            payload: JSON.stringify({
              name: f.name,
              parents: [destParent],
              appProperties: driveMark_(row.CAPID, f.id, 'file')
            }),
            muteHttpExceptions: true
          };
        }),
        'copy file'
      );

      copies.forEach((res, i) => {
        if (!res.ok) {
          // A failed copy is not skippable by default — the file exists in the
          // source and not the destination. But one unrecoverable file must not
          // strand the entire Drive, so record it as a skip that blocks deletion
          // and press on.
          skipped.push(`"${copyable[i].name}" (${copyable[i].id}): copy failed`);
          Logger.error('Drive file copy failed — SKIPPED, will not survive deletion', {
            capid: row.CAPID, fileId: copyable[i].id, name: copyable[i].name, error: res.error
          });
          return;
        }
        done.doneIds.add(copyable[i].id);
        copied++;
      });
    }

    pageToken = page.nextPageToken || '';
    setTransitionField_(row._rowNumber, 'DriveCursor', pageToken);

  } while (pageToken);

  return { complete: true, copied: copied, skipped: skipped, cursor: '' };
}

// ============================================================================
// CONTINUATION
// ============================================================================

/**
 * Schedules a one-shot Drive continuation, guarded against stacking.
 *
 * @param {Object} scope - {capid}
 */
function scheduleDriveContinuation_(scope) {
  const already = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === DRIVE_CONTINUATION_FN_);
  if (already) return;

  PropertiesService.getScriptProperties().setProperty(
    DRIVE_CONTINUATION_SCOPE_PROP_, JSON.stringify(scope || {})
  );
  ScriptApp.newTrigger(DRIVE_CONTINUATION_FN_).timeBased().after(60 * 1000).create();
  Logger.info('Drive continuation scheduled', { scope: scope || {} });
}

/** Removes pending Drive continuations. Manual recovery aid. */
function clearDriveContinuations() {
  let removed = 0;
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === DRIVE_CONTINUATION_FN_)
    .forEach(t => { ScriptApp.deleteTrigger(t); removed++; });
  Logger.info('Drive continuations cleared', { removed: removed });
  return { removed: removed };
}

// ============================================================================
// PREVIEW
// ============================================================================

/**
 * Reports what a Drive migration would move. Copies nothing.
 *
 * Worth running before building anything on top of this: cadet Drive usage
 * varies from nothing to gigabytes, and the answer decides whether this is a
 * five-minute job or an all-day one.
 */
function previewCadetTransitionDrive() {
  if (TRANSITION_CONFIG.ROLE !== 'source') {
    console.log('Not the source tenant.');
    return;
  }

  const rows = readTransitions_();
  let grandFiles = 0;
  let grandBytes = 0;

  console.log('CAPID | name | files | folders | size | largest');
  console.log('------------------------------------------------');

  for (const capid in rows) {
    const row = rows[capid];
    if (!row.SeniorEmail) continue;

    try {
      const stats = driveStats_(row.CadetEmail);
      grandFiles += stats.files;
      grandBytes += stats.bytes;

      console.log([
        capid,
        row.Name,
        stats.files + ' files',
        stats.folders + ' folders',
        formatBytes_(stats.bytes),
        stats.largest ? formatBytes_(stats.largestSize) + ' (' + stats.largest + ')' : '-'
      ].join(' | '));

    } catch (e) {
      console.log(`${capid} | ${row.Name} | ERROR: ${e && e.message ? e.message : String(e)}`);
    }
  }

  console.log('');
  console.log(`Total: ${grandFiles} files, ${formatBytes_(grandBytes)}`);
  console.log('Nothing was copied — this is a preview.');
}

/**
 * Counts owned files in a source Drive.
 *
 * 'me' in owners is the important filter: files merely shared with the cadet
 * belong to someone else and survive the deletion on their own.
 *
 * @param {string} cadetEmail
 * @returns {{files: number, folders: number, bytes: number, largest: string, largestSize: number}}
 */
function driveStats_(cadetEmail) {
  const token = getImpersonatedToken_(cadetEmail, 'https://www.googleapis.com/auth/drive');
  let pageToken = '';
  const out = { files: 0, folders: 0, bytes: 0, largest: '', largestSize: 0 };

  do {
    const url = 'https://www.googleapis.com/drive/v3/files' +
      '?q=' + encodeURIComponent("'me' in owners and trashed = false") +
      '&pageSize=1000&fields=nextPageToken,files(id,name,mimeType,size)' +
      (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

    const body = driveFetch_(url, { method: 'get', token: token }, 'list drive for ' + cadetEmail);

    (body.files || []).forEach(f => {
      if (f.mimeType === DRIVE_FOLDER_MIME_) {
        out.folders++;
        return;
      }
      out.files++;
      const size = Number(f.size) || 0;   // Docs-native files report no size
      out.bytes += size;
      if (size > out.largestSize) {
        out.largestSize = size;
        out.largest = f.name;
      }
    });

    pageToken = body.nextPageToken || '';
  } while (pageToken);

  return out;
}

// ============================================================================
// HTTP
// ============================================================================

/**
 * One Drive REST call with retry on transient failures.
 *
 * Mirrors gmailFetch_ and for the same reasons — the shared executeWithRetry()
 * sleeps 3s before every call and detects transience via e.details?.code, which
 * raw UrlFetchApp calls never set.
 *
 * @param {string} url
 * @param {{method: string, token: string, payload: string=}} opts
 * @param {string} what
 * @returns {Object}
 */
function driveFetch_(url, opts, what) {
  const params = {
    method: opts.method,
    headers: { Authorization: 'Bearer ' + opts.token },
    muteHttpExceptions: true
  };
  if (opts.payload) {
    params.contentType = 'application/json';
    params.payload = opts.payload;
  }

  let lastError = '';

  for (let attempt = 1; attempt <= MIGRATE_MAX_ATTEMPTS_; attempt++) {
    const resp = UrlFetchApp.fetch(url, params);
    const code = resp.getResponseCode();
    const text = resp.getContentText();

    if (code >= 200 && code < 300) return text ? JSON.parse(text) : {};

    lastError = `Drive ${what} failed (${code}): ${text}`;

    const transient = code === 429 || (code >= 500 && code < 600);
    if (!transient || attempt === MIGRATE_MAX_ATTEMPTS_) throw new Error(lastError);

    Utilities.sleep(MIGRATE_BACKOFF_BASE_MS_ * Math.pow(2, attempt - 1));
  }

  throw new Error(lastError);
}

// formatBytes_ lives in CadetTransition.gs — the migration path needs it to
// report oversized messages, and depending on this module for it would be
// backwards.
