/*******************************************************
 * Group Administration Utilities
 *
 * Filename: groupAdministration.gs
 * Saved: 2026-04-07 18:52 PDT
 *
 * Apps Script equivalents for common GAM group-admin tasks.
 *
 * Run-input usage:
 * - Set GROUP_ADMINISTRATION_RUN_INPUTS.GROUP_EMAIL at the top of this file
 *   to run the group-targeted helpers directly from the Apps Script editor
 *   without passing an argument.
 * - Set GROUP_ADMINISTRATION_RUN_INPUTS.DOMAIN_SHARED_CONTACT_EMAIL for
 *   shared-contact deletion.
 * - Set GROUP_ADMINISTRATION_RUN_INPUTS.DELETE_GROUPS_SHEET or
 *   GROUP_ADMINISTRATION_RUN_INPUTS.DELETE_DOMAIN_SHARED_CONTACTS_SHEET
 *   to control the default worklist tab names for bulk operations.
 *
 * Public functions:
 * - groupAdministration_listGroups()
 *   Lists all Google Groups in the current customer.
 *
 * - groupAdministration_listGroupsNoMembers()
 *   Lists groups whose direct member count is zero.
 *
 * - groupAdministration_writeAllGroupsReport()
 *   Writes a full groups inventory to the "All Groups" tab in the
 *   automation spreadsheet.
 *
 * - groupAdministration_writeNoMemberGroupsReport()
 *   Writes zero-member groups to the "No Member Groups" tab in the
 *   automation spreadsheet.
 *
 * - groupAdministration_previewStaleGroups()
 *   Safe preview helper that generates both reports above for review
 *   before any deletion.
 *
 * - groupAdministration_deleteGroup(groupEmail)
 *   Permanently deletes a single Google Group. If groupEmail is omitted,
 *   uses GROUP_ADMINISTRATION_RUN_INPUTS.GROUP_EMAIL.
 *
 * - groupAdministration_bulkDeleteGroupsFromSheet(sheetName)
 *   Deletes groups listed in a spreadsheet worklist tab. Expected header:
 *   "group" or "email". If sheetName is omitted, uses
 *   GROUP_ADMINISTRATION_RUN_INPUTS.DELETE_GROUPS_SHEET.
 *
 * - groupAdministration_clearGroup(groupEmail)
 *   Removes all direct members from a group but keeps the group itself.
 *   If groupEmail is omitted, uses GROUP_ADMINISTRATION_RUN_INPUTS.GROUP_EMAIL.
 *
 * - groupAdministration_hideGroupFromGal(groupEmail)
 *   Sets includeInGlobalAddressList=false so a group is hidden from the
 *   Gmail directory / Global Address List. If groupEmail is omitted, uses
 *   GROUP_ADMINISTRATION_RUN_INPUTS.GROUP_EMAIL.
 *
 * - groupAdministration_deleteDomainSharedContact(email)
 *   Deletes a single Domain Shared Contact by email address. If email is
 *   omitted, uses GROUP_ADMINISTRATION_RUN_INPUTS.DOMAIN_SHARED_CONTACT_EMAIL.
 *
 * - groupAdministration_bulkDeleteDomainSharedContactsFromSheet(sheetName)
 *   Deletes Domain Shared Contacts listed in a spreadsheet worklist tab.
 *   Expected header: "email". If sheetName is omitted, uses
 *   GROUP_ADMINISTRATION_RUN_INPUTS.DELETE_DOMAIN_SHARED_CONTACTS_SHEET.
 *
 * - groupAdministration_deleteUserContactsForAllUsers_notSupported()
 *   Explicit marker that Apps Script cannot centrally delete personal user
 *   contacts or recent-recipient autocomplete entries for all users.
 *
 * - groupAdministration_previewConfiguredGroups()
 *   Checks the configured stale group email list at the top of this file and
 *   reports which groups still exist in Google Directory.
 *
 * - groupAdministration_deleteConfiguredGroups()
 *   Deletes the configured stale group email list at the top of this file.
 *   Missing groups are logged as already gone instead of throwing.
 *
 * - groupAdministration_auditReceiveListPosting()
 *   Read-only audit of whoCanPostMessage / allowExternalMembers on managed
 *   .cadets/.parents/.all "receive lists". Run it on the tenant that OWNS those
 *   groups (e.g. the cadets tenant) to find sublists that would silently reject
 *   cross-tenant fan-out from a wing .all list.
 *
 * Notes:
 * - Google Groups cannot be restored after deletion.
 * - Apps Script can manage Google Groups and Domain Shared Contacts.
 * - Apps Script cannot centrally delete personal user contacts/recent-recipient
 *   autocomplete entries for all users the way GAM can target user contacts.
 *******************************************************/

/**
 * Enter the target values here when you want to run the no-argument helpers
 * directly from the Apps Script editor.
 *
 * Example:
 *   GROUP_EMAIL: 'ca.dty.group-staff-only@cawgcap.org'
 */
const GROUP_ADMINISTRATION_RUN_INPUTS = {
  GROUP_EMAIL: '',
  DOMAIN_SHARED_CONTACT_EMAIL: '',
  DELETE_GROUPS_SHEET: 'Delete Groups',
  DELETE_DOMAIN_SHARED_CONTACTS_SHEET: 'Delete Domain Shared Contacts'
};

/**
 * Optional convenience list for stale group cleanup.
 * Add full group email addresses here, then run:
 * - groupAdministration_previewConfiguredGroups()
 * - groupAdministration_deleteConfiguredGroups()
 */
const GROUP_ADMINISTRATION_STALE_GROUP_EMAILS = [
  'ca.dty.group-staff-only@cawgcap.org',
  'ca.dty.wing-staff-only@cawgcap.org',
  'ca070.dty.group-staff-only@cawgcap.org',
  'ca188.dty.group-staff-only@cawgcap.org'
];

/**
 * Returns all groups in the current customer.
 *
 * @returns {Array<Object>}
 */
function groupAdministration_listGroups() {
  const out = [];
  let pageToken = '';

  do {
    const res = executeWithRetry(() => AdminDirectory.Groups.list({
      customer: CONFIG.CUSTOMER_ID || 'my_customer',
      maxResults: 200,
      pageToken: pageToken
    }));

    const groups = res.groups || [];
    for (let i = 0; i < groups.length; i++) {
      out.push({
        email: String(groups[i].email || '').toLowerCase(),
        name: String(groups[i].name || ''),
        description: String(groups[i].description || ''),
        directMembersCount: Number(groups[i].directMembersCount || 0),
        adminCreated: String(groups[i].adminCreated || ''),
        id: String(groups[i].id || '')
      });
    }

    pageToken = res.nextPageToken || '';
  } while (pageToken);

  out.sort((a, b) => a.email.localeCompare(b.email));

  Logger.info('Listed groups', { count: out.length });
  return out;
}

/**
 * Returns groups with zero direct members.
 *
 * @returns {Array<Object>}
 */
function groupAdministration_listGroupsNoMembers() {
  const groups = groupAdministration_listGroups().filter(g => g.directMembersCount === 0);
  Logger.info('Listed groups with no members', { count: groups.length });
  return groups;
}

/**
 * Writes a groups report to the automation spreadsheet.
 *
 * @param {boolean} noMembersOnly
 * @param {string=} sheetName
 * @returns {number} number of data rows written
 */
function groupAdministration_writeGroupsReport(noMembersOnly, sheetName) {
  const targetSheetName = String(sheetName || (noMembersOnly ? 'No Member Groups' : 'All Groups')).trim();
  const groups = noMembersOnly ? groupAdministration_listGroupsNoMembers() : groupAdministration_listGroups();
  const ss = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID);
  const sheet = getOrCreateSheet_(ss, targetSheetName);

  const rows = [[
    'Email',
    'Name',
    'Description',
    'Direct Members Count',
    'Admin Created',
    'Group ID'
  ]];

  for (let i = 0; i < groups.length; i++) {
    rows.push([
      groups[i].email,
      groups[i].name,
      groups[i].description,
      groups[i].directMembersCount,
      groups[i].adminCreated,
      groups[i].id
    ]);
  }

  writeTabularData_(sheet, rows);

  Logger.info('Wrote groups report', {
    sheetName: targetSheetName,
    rows: Math.max(rows.length - 1, 0),
    noMembersOnly: !!noMembersOnly
  });

  return Math.max(rows.length - 1, 0);
}

/**
 * Convenience runner: writes every group to "All Groups".
 *
 * @returns {number}
 */
function groupAdministration_writeAllGroupsReport() {
  return groupAdministration_writeGroupsReport(false, 'All Groups');
}

/**
 * Convenience runner: writes no-member groups to "No Member Groups".
 *
 * @returns {number}
 */
function groupAdministration_writeNoMemberGroupsReport() {
  return groupAdministration_writeGroupsReport(true, 'No Member Groups');
}

/**
 * Deletes a single Google Group.
 *
 * @param {string} groupEmail
 * @returns {{group:string,status:string}}
 */
function groupAdministration_deleteGroup(groupEmail) {
  const email = getConfiguredRunEmail_(groupEmail, 'GROUP_EMAIL', 'group email');
  if (!email) throw new Error('Invalid group email');

  executeWithRetry(() => AdminDirectory.Groups.remove(email));
  Logger.warn('Deleted group', { group: email });

  return { group: email, status: 'deleted' };
}

/**
 * Checks whether a group exists in Google Directory.
 *
 * @param {string} groupEmail
 * @returns {{group:string,exists:boolean,name:string}}
 */
function groupAdministration_groupExists(groupEmail) {
  const email = sanitizeEmail(groupEmail);
  if (!email) throw new Error('Invalid group email');

  try {
    const group = executeWithRetry(() => AdminDirectory.Groups.get(email));
    return {
      group: email,
      exists: true,
      name: String(group.name || '')
    };
  } catch (e) {
    if (e.details && e.details.code === 404) {
      return {
        group: email,
        exists: false,
        name: ''
      };
    }
    throw e;
  }
}

/**
 * Deletes groups listed in a sheet. Expected header: "group" or "email".
 *
 * @param {string=} sheetName
 * @returns {{deleted:number,failed:number,groups:Array<Object>}}
 */
function groupAdministration_bulkDeleteGroupsFromSheet(sheetName) {
  const targetSheetName = String(sheetName || GROUP_ADMINISTRATION_RUN_INPUTS.DELETE_GROUPS_SHEET || 'Delete Groups').trim();
  const rows = readSingleColumnSheetObjects_(targetSheetName, ['group', 'email']);
  let deleted = 0;
  let failed = 0;
  const results = [];

  for (let i = 0; i < rows.length; i++) {
    const email = sanitizeEmail(rows[i].value);
    if (!email) continue;

    try {
      groupAdministration_deleteGroup(email);
      deleted++;
      results.push({ group: email, status: 'deleted' });
    } catch (e) {
      failed++;
      results.push({ group: email, status: 'failed', errorMessage: e.message });
      Logger.error('Bulk delete group failed', {
        group: email,
        errorMessage: e.message,
        errorCode: e.details && e.details.code
      });
    }
  }

  Logger.info('Bulk delete groups completed', {
    sheetName: targetSheetName,
    deleted: deleted,
    failed: failed
  });

  return { deleted: deleted, failed: failed, groups: results };
}

/**
 * Removes all direct members from a group but keeps the group.
 *
 * @param {string} groupEmail
 * @returns {{group:string,removed:number}}
 */
function groupAdministration_clearGroup(groupEmail) {
  const email = getConfiguredRunEmail_(groupEmail, 'GROUP_EMAIL', 'group email');
  if (!email) throw new Error('Invalid group email');

  let removed = 0;
  let pageToken = '';

  do {
    const res = executeWithRetry(() => AdminDirectory.Members.list(email, {
      maxResults: 200,
      pageToken: pageToken
    }));

    const members = res.members || [];
    for (let i = 0; i < members.length; i++) {
      const memberKey = String(members[i].id || members[i].email || '').trim();
      if (!memberKey) continue;
      executeWithRetry(() => AdminDirectory.Members.remove(email, memberKey));
      removed++;
    }

    pageToken = res.nextPageToken || '';
  } while (pageToken);

  Logger.warn('Cleared group members', {
    group: email,
    removed: removed
  });

  return { group: email, removed: removed };
}

/**
 * Hides a group from the Gmail Global Address List / directory.
 *
 * @param {string} groupEmail
 * @returns {{group:string,includeInGlobalAddressList:string}}
 */
function groupAdministration_hideGroupFromGal(groupEmail) {
  const email = getConfiguredRunEmail_(groupEmail, 'GROUP_EMAIL', 'group email');
  if (!email) throw new Error('Invalid group email');

  if (typeof AdminGroupsSettings === 'undefined' || !AdminGroupsSettings.Groups || !AdminGroupsSettings.Groups.patch) {
    throw new Error('AdminGroupsSettings advanced service is not enabled');
  }

  executeWithRetry(() => AdminGroupsSettings.Groups.patch({
    includeInGlobalAddressList: 'false'
  }, email));

  Logger.info('Group hidden from GAL', {
    group: email,
    includeInGlobalAddressList: 'false'
  });

  return { group: email, includeInGlobalAddressList: 'false' };
}

/**
 * Deletes a single Domain Shared Contact by email.
 * Uses the helper already implemented in SharedContacts.gs.
 *
 * @param {string} email
 * @returns {string}
 */
function groupAdministration_deleteDomainSharedContact(email) {
  const targetEmail = getConfiguredRunEmail_(email, 'DOMAIN_SHARED_CONTACT_EMAIL', 'domain shared contact email');
  return deleteExternalContactByEmail(targetEmail);
}

/**
 * Deletes Domain Shared Contacts listed in a sheet. Expected header: "email".
 *
 * @param {string=} sheetName
 * @returns {{deleted:number,notFound:number,failed:number,contacts:Array<Object>}}
 */
function groupAdministration_bulkDeleteDomainSharedContactsFromSheet(sheetName) {
  const targetSheetName = String(
    sheetName || GROUP_ADMINISTRATION_RUN_INPUTS.DELETE_DOMAIN_SHARED_CONTACTS_SHEET || 'Delete Domain Shared Contacts'
  ).trim();
  const rows = readSingleColumnSheetObjects_(targetSheetName, ['email']);
  let deleted = 0;
  let notFound = 0;
  let failed = 0;
  const results = [];

  for (let i = 0; i < rows.length; i++) {
    const email = sanitizeEmail(rows[i].value);
    if (!email) continue;

    const status = deleteExternalContactByEmail(email);
    results.push({ email: email, status: status });

    if (status === 'deleted') deleted++;
    else if (status === 'not_found') notFound++;
    else failed++;
  }

  Logger.info('Bulk delete domain shared contacts completed', {
    sheetName: targetSheetName,
    deleted: deleted,
    notFound: notFound,
    failed: failed
  });

  return {
    deleted: deleted,
    notFound: notFound,
    failed: failed,
    contacts: results
  };
}

/**
 * Apps Script does not provide an admin-side equivalent to
 * "gam all users delete contact ...".
 * This function exists to make that limitation explicit in the script UI.
 */
function groupAdministration_deleteUserContactsForAllUsers_notSupported() {
  const message = 'Apps Script cannot centrally delete personal contacts/recent recipients for all users. Use GAM for that workflow.';
  Logger.warn(message, {});
  throw new Error(message);
}

/**
 * Quick preview runner for stale-group cleanup candidates.
 *
 * @returns {{allGroups:number,noMemberGroups:number}}
 */
function groupAdministration_previewStaleGroups() {
  const allGroups = groupAdministration_writeAllGroupsReport();
  const noMemberGroups = groupAdministration_writeNoMemberGroupsReport();
  return {
    allGroups: allGroups,
    noMemberGroups: noMemberGroups
  };
}

/**
 * Checks the configured stale group list and reports which still exist.
 *
 * @returns {{existing:Array<Object>,missing:Array<Object>}}
 */
function groupAdministration_previewConfiguredGroups() {
  const emails = getConfiguredGroupEmailList_();
  const existing = [];
  const missing = [];

  for (let i = 0; i < emails.length; i++) {
    const result = groupAdministration_groupExists(emails[i]);
    if (result.exists) existing.push(result);
    else missing.push(result);
  }

  Logger.info('Configured stale groups preview', {
    existingCount: existing.length,
    missingCount: missing.length,
    existing: existing,
    missing: missing
  });

  return { existing: existing, missing: missing };
}

/**
 * Deletes the configured stale group list.
 * Missing groups are logged and skipped.
 *
 * @returns {{deleted:Array<Object>,missing:Array<Object>,failed:Array<Object>}}
 */
function groupAdministration_deleteConfiguredGroups() {
  const emails = getConfiguredGroupEmailList_();
  const deleted = [];
  const missing = [];
  const failed = [];

  for (let i = 0; i < emails.length; i++) {
    const email = emails[i];
    try {
      const exists = groupAdministration_groupExists(email);
      if (!exists.exists) {
        missing.push({ group: email, status: 'not_found' });
        continue;
      }

      groupAdministration_deleteGroup(email);
      deleted.push({ group: email, status: 'deleted' });
    } catch (e) {
      failed.push({
        group: email,
        status: 'failed',
        errorMessage: e.message
      });
      Logger.error('Configured stale group delete failed', {
        group: email,
        errorMessage: e.message,
        errorCode: e.details && e.details.code
      });
    }
  }

  Logger.info('Configured stale groups delete completed', {
    deletedCount: deleted.length,
    missingCount: missing.length,
    failedCount: failed.length,
    deleted: deleted,
    missing: missing,
    failed: failed
  });

  return { deleted: deleted, missing: missing, failed: failed };
}

/**
 * Read-only audit of "receive-list" posting permissions.
 *
 * Cross-tenant fan-out (wing ca###.all -> cadet ca###.cadets@cawgcadets.org)
 * only delivers if the RECEIVING group accepts mail from the original, external
 * sender. A receiving sublist set to ALL_MEMBERS_CAN_POST or
 * ALL_IN_DOMAIN_CAN_POST silently rejects/holds forwarded wing mail; it
 * generally needs ANYONE_CAN_POST (ideally paired with spam moderation).
 *
 * Run this on the tenant that OWNS the receiving groups (e.g. the cadets
 * tenant). It reads, per managed .cadets/.parents/.all group:
 *   whoCanPostMessage, allowExternalMembers, messageModerationLevel,
 *   spamModerationLevel
 * and flags any whose whoCanPostMessage would block external fan-out.
 * This function only reads settings; it changes nothing.
 *
 * @returns {{checked:number, blocking:Array<Object>, ok:Array<Object>}}
 */
function groupAdministration_auditReceiveListPosting() {
  if (typeof AdminGroupsSettings === 'undefined' || !AdminGroupsSettings.Groups || !AdminGroupsSettings.Groups.get) {
    throw new Error('AdminGroupsSettings advanced service is not enabled');
  }

  const suffixRe = /\.(cadets|parents|all)$/i;
  const groups = groupAdministration_listGroups().filter(g => {
    const local = String(g.email || '').split('@')[0];
    return suffixRe.test(local);
  });

  const blocking = [];
  const ok = [];

  for (let i = 0; i < groups.length; i++) {
    const email = groups[i].email;
    let settings;
    try {
      settings = executeWithRetry(() => AdminGroupsSettings.Groups.get(email));
    } catch (e) {
      Logger.warn('Could not read group settings during receive-list audit', {
        group: email,
        errorMessage: e.message
      });
      continue;
    }

    const row = {
      email: email,
      whoCanPostMessage: String(settings.whoCanPostMessage || ''),
      allowExternalMembers: String(settings.allowExternalMembers || ''),
      messageModerationLevel: String(settings.messageModerationLevel || ''),
      spamModerationLevel: String(settings.spamModerationLevel || '')
    };

    // ANYONE_CAN_POST is the only value that reliably accepts external fan-out.
    if (row.whoCanPostMessage === 'ANYONE_CAN_POST') {
      ok.push(row);
    } else {
      blocking.push(row);
    }
  }

  blocking.sort((a, b) => a.email.localeCompare(b.email));
  ok.sort((a, b) => a.email.localeCompare(b.email));

  Logger.info('Receive-list posting audit complete', {
    checked: groups.length,
    blockingCount: blocking.length,
    okCount: ok.length,
    blocking: blocking.slice(0, 50)
  });

  return { checked: groups.length, blocking: blocking, ok: ok };
}

/**
 * Resolves a caller-supplied email or falls back to the top-of-file run inputs.
 *
 * @param {string=} value
 * @param {string} configKey
 * @param {string} label
 * @returns {string}
 */
function getConfiguredRunEmail_(value, configKey, label) {
  const candidate = String(
    value || (GROUP_ADMINISTRATION_RUN_INPUTS && GROUP_ADMINISTRATION_RUN_INPUTS[configKey]) || ''
  ).trim();
  const email = sanitizeEmail(candidate);
  if (!email) {
    throw new Error('Set ' + configKey + ' at the top of groupAdministration.gs or pass a valid ' + label + ' to the function.');
  }
  return email;
}

/**
 * Returns the validated configured stale group email list.
 *
 * @returns {Array<string>}
 */
function getConfiguredGroupEmailList_() {
  const out = [];
  for (let i = 0; i < GROUP_ADMINISTRATION_STALE_GROUP_EMAILS.length; i++) {
    const email = sanitizeEmail(String(GROUP_ADMINISTRATION_STALE_GROUP_EMAILS[i] || '').trim());
    if (email) out.push(email);
  }

  if (!out.length) {
    throw new Error('Add one or more full group email addresses to GROUP_ADMINISTRATION_STALE_GROUP_EMAILS at the top of groupAdministration.gs.');
  }

  return out;
}

/**
 * Returns or creates a sheet by name.
 *
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @param {string} sheetName
 * @returns {SpreadsheetApp.Sheet}
 */
function getOrCreateSheet_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

/**
 * Replaces a sheet's contents with tabular data.
 *
 * @param {SpreadsheetApp.Sheet} sheet
 * @param {Array<Array<*>>} rows
 * @returns {void}
 */
function writeTabularData_(sheet, rows) {
  sheet.clearContents();
  sheet.clearFormats();

  if (!rows.length) return;

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, rows[0].length);
}

/**
 * Reads a one-column worklist from the automation spreadsheet.
 *
 * @param {string} sheetName
 * @param {Array<string>} allowedHeaders lowercase header names to accept
 * @returns {Array<{row:number,value:string}>}
 */
function readSingleColumnSheetObjects_(sheetName, allowedHeaders) {
  const ss = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Missing sheet: ' + sheetName);
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0].map(h => String(h || '').trim().toLowerCase());
  let col = -1;

  for (let i = 0; i < allowedHeaders.length; i++) {
    col = header.indexOf(String(allowedHeaders[i] || '').trim().toLowerCase());
    if (col > -1) break;
  }

  if (col < 0) {
    throw new Error('Missing required header in sheet ' + sheetName + ': ' + allowedHeaders.join(' or '));
  }

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const value = String(values[r][col] || '').trim();
    if (!value) continue;
    out.push({ row: r + 1, value: value });
  }
  return out;
}
