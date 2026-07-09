/***********************************************
 * File: SyncOrgPaths.gs
 * Description: Detects activated and deactivated CA Wing orgs after each
 *   CAPWATCH download by comparing Organization.txt against OrgPaths.txt.
 *   Auto-adds new entries where the parent path is already mapped, and emails
 *   this tenant's IT mailbox with a full summary of changes and items needing
 *   attention. Called automatically at the end of getCapwatch().
 ***********************************************/

/**
 * Recipient for the OrgPath sync summary — this tenant's IT mailbox.
 * Resolved at call time (not as a top-level const) so it reads the per-tenant
 * ITSUPPORT_EMAIL from config.gs regardless of cross-file load order.
 * @returns {string}
 */
function getOrgPathSyncEmail_() {
  return (typeof ITSUPPORT_EMAIL !== 'undefined' && ITSUPPORT_EMAIL) ||
         (typeof TEST_EMAIL !== 'undefined' && TEST_EMAIL) || '';
}

/**
 * Syncs OrgPaths.txt with active orgs in Organization.txt.
 *
 * - Auto-adds any new UNIT or GROUP whose parent ORGID is already mapped,
 *   deriving the path as parentPath + "/CA-" + unitNumber.
 * - Flags new orgs whose parent isn't mapped yet (needs manual entry).
 * - Flags OrgPath entries with no matching active org in CAPWATCH (possible
 *   deactivation or recharter) for manual cleanup.
 * - Sends an email summary whenever anything changes or needs attention.
 *
 * @returns {{added: number, needsManual: number, deactivated: number}}
 */
function syncOrgPaths() {
  Logger.info('Starting OrgPath sync');

  const orgRows = parseFile('Organization');
  const orgPathRows = parseFile('OrgPaths');

  // Build ORGID → path map from current OrgPaths.txt
  const currentPaths = {};
  orgPathRows.forEach(row => {
    const orgid = String(row[0] || '').trim();
    const path  = String(row[1] || '').trim();
    if (orgid && path) currentPaths[orgid] = path;
  });

  // Build ORGID → metadata map for all active CA Wing orgs from Organization.txt
  // Columns: ORGID, Region, Wing, Unit, NextLevel(parentORGID), Name, Type,
  //          DateChartered, Status, Scope, ...
  const activeCAOrgs = {};
  orgRows.forEach(row => {
    const orgid  = String(row[0] || '').trim();
    const region = String(row[1] || '').trim();
    const wing   = String(row[2] || '').trim();
    const unit   = String(row[3] || '').trim();
    const parent = String(row[4] || '').trim();
    const name   = String(row[5] || '').trim();
    const status = String(row[8] || '').trim();
    const scope  = String(row[9] || '').trim();

    if (region === 'PCR' && wing === CONFIG.WING && status === 'ACTIVE'
        && (scope === 'UNIT' || scope === 'GROUP' || scope === 'WING')) {
      activeCAOrgs[orgid] = { orgid, unit, parent, name, scope };
    }
  });

  const added       = [];
  const needsManual = [];
  const deactivated = [];

  // --- Detect new orgs not yet in OrgPaths ---
  Object.values(activeCAOrgs).forEach(org => {
    if (currentPaths[org.orgid]) return; // already mapped

    const parentPath = currentPaths[org.parent];
    if (parentPath) {
      const newPath = parentPath + '/CA-' + org.unit;
      currentPaths[org.orgid] = newPath;
      added.push({ orgid: org.orgid, path: newPath, name: org.name, scope: org.scope });
      Logger.info('OrgPath auto-added', { orgid: org.orgid, path: newPath, name: org.name });
    } else {
      needsManual.push({
        orgid: org.orgid, unit: org.unit, name: org.name,
        scope: org.scope, parent: org.parent
      });
      Logger.warn('OrgPath cannot be auto-derived — parent not mapped', {
        orgid: org.orgid, name: org.name, parentOrgid: org.parent
      });
    }
  });

  // --- Detect deactivations: in OrgPaths but absent from active CA orgs ---
  // Skip non-CA entries (NHQ, PCR, etc.) — they won't appear in Organization.txt
  Object.entries(currentPaths).forEach(([orgid, path]) => {
    if (!path.startsWith('/CA-001')) return;
    if (!activeCAOrgs[orgid]) {
      deactivated.push({ orgid, path });
      Logger.warn('OrgPath entry has no matching active CA org', { orgid, path });
    }
  });

  // --- Create OUs in Google Workspace and write entries to OrgPaths.txt ---
  if (added.length > 0) {
    added.forEach(e => { e.ouCreated = createOrgUnit_(e.path, e.name); });
    writeOrgPathEntries_(added.map(e => ({ orgid: e.orgid, path: e.path })));
  }

  // --- Email summary if anything needs attention ---
  if (added.length > 0 || needsManual.length > 0 || deactivated.length > 0) {
    sendOrgPathSyncEmail_(added, needsManual, deactivated);
  } else {
    Logger.info('OrgPath sync complete — no changes detected');
  }

  return { added: added.length, needsManual: needsManual.length, deactivated: deactivated.length };
}

/**
 * Appends new ORGID,OrgUnitPath lines to OrgPaths.txt on Drive.
 *
 * @param {Array<{orgid: string, path: string}>} entries
 */
function writeOrgPathEntries_(entries) {
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files  = folder.getFilesByName('OrgPaths.txt');

  if (!files.hasNext()) {
    Logger.error('OrgPaths.txt not found on Drive — cannot write new entries');
    return;
  }

  const file = files.next();
  let content = file.getBlob().getDataAsString();
  if (!content.endsWith('\n')) content += '\n';
  entries.forEach(e => { content += e.orgid + ',' + e.path + '\n'; });
  file.setContent(content);

  Logger.info('OrgPaths.txt updated on Drive', { entriesAdded: entries.length });
}

/**
 * Creates a Google Workspace Organizational Unit at the given path.
 * Derives the parent OU path and OU name from the full path.
 * Returns true if created, false if it already existed or an error occurred.
 *
 * @param {string} fullPath  - e.g. "/CA-001/CA-445/CA-404"
 * @param {string} ouName    - Human-readable name for the OU (e.g. squadron name)
 * @returns {boolean}
 */
function createOrgUnit_(fullPath, ouName) {
  const lastSlash = fullPath.lastIndexOf('/');
  const parentPath = lastSlash > 0 ? fullPath.substring(0, lastSlash) : '/';
  const name = fullPath.substring(lastSlash + 1);

  try {
    AdminDirectory.Orgunits.insert(
      { name: name, parentOrgUnitPath: parentPath, description: ouName },
      'my_customer'
    );
    Logger.info('Org unit created', { path: fullPath, name: ouName });
    return true;
  } catch (e) {
    // 409 = already exists — not an error worth alerting on
    if (e.message && e.message.indexOf('409') !== -1) {
      Logger.info('Org unit already exists', { path: fullPath });
    } else {
      Logger.error('Failed to create org unit', { path: fullPath, errorMessage: e.message });
    }
    return false;
  }
}

/**
 * Sends an HTML summary email to this tenant's IT mailbox (getOrgPathSyncEmail_()).
 *
 * @param {Array} added       - Entries auto-added to OrgPaths.txt
 * @param {Array} needsManual - New orgs whose parent isn't mapped
 * @param {Array} deactivated - OrgPath entries with no active org match
 */
function sendOrgPathSyncEmail_(added, needsManual, deactivated) {
  const date    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy');
  const subject = `[CAWG Automation] OrgPath Sync Report — ${date}`;

  let html = `<p>The daily CAPWATCH sync detected OrgPath changes on <strong>${date}</strong>.</p>`;

  if (added.length > 0) {
    html += `
      <h3 style="color:#1a7a3c">&#x2705; Auto-Added (${added.length})</h3>
      <p>These entries were written to OrgPaths.txt and provisioned in Google Workspace automatically. No action needed.</p>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:monospace;font-size:13px">
        <tr style="background:#f0f0f0"><th>ORGID</th><th>Name</th><th>Scope</th><th>OrgUnitPath</th><th>OU Created</th></tr>
        ${added.map(e =>
          `<tr><td>${e.orgid}</td><td>${e.name}</td><td>${e.scope}</td><td>${e.path}</td><td>${e.ouCreated ? '&#x2705;' : '&#x26A0;&#xFE0F; already existed or error'}</td></tr>`
        ).join('')}
      </table>`;
  }

  if (needsManual.length > 0) {
    html += `
      <h3 style="color:#b85c00">&#x26A0;&#xFE0F; Needs Manual Mapping (${needsManual.length})</h3>
      <p>These active CA orgs are not in OrgPaths.txt and their parent ORGID is not mapped either.
         Add them manually once the parent group OU is created.</p>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:monospace;font-size:13px">
        <tr style="background:#f0f0f0"><th>ORGID</th><th>Unit</th><th>Name</th><th>Scope</th><th>Parent ORGID</th></tr>
        ${needsManual.map(e =>
          `<tr><td>${e.orgid}</td><td>${e.unit}</td><td>${e.name}</td><td>${e.scope}</td><td>${e.parent}</td></tr>`
        ).join('')}
      </table>`;
  }

  if (deactivated.length > 0) {
    html += `
      <h3 style="color:#c0392b">&#x1F534; Possible Deactivations (${deactivated.length})</h3>
      <p>These OrgPath entries have no matching active CA org in CAPWATCH. If the unit has been
         deactivated or rechartered, remove the entry from OrgPaths.txt and clean up any associated
         calendar resources and Google Groups.</p>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:monospace;font-size:13px">
        <tr style="background:#f0f0f0"><th>ORGID</th><th>OrgUnitPath</th></tr>
        ${deactivated.map(e =>
          `<tr><td>${e.orgid}</td><td>${e.path}</td></tr>`
        ).join('')}
      </table>`;
  }

  html += `<p style="color:#888;font-size:11px;margin-top:24px">Sent by CAWG Automation — OrgPath Sync</p>`;

  const syncEmail = getOrgPathSyncEmail_();
  MailApp.sendEmail({ to: syncEmail, subject: subject, htmlBody: html });
  Logger.info('OrgPath sync email sent', { to: syncEmail });
}
