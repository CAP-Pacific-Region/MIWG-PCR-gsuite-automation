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
