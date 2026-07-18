/*******************************************************
 * Group Administration Utilities
 *
 * Filename: groupAdministration.gs
 * Saved: 2026-07-17
 * Changes: Added groupAdministration_stageLegacyDlGroups() (read-only bulk scan
 *   for legacy 'DL-CAWG-*' migration groups/aliases -> review sheet) and
 *   groupAdministration_resolveLegacyAddress() (definitive single-address
 *   group/alias/not-a-group check). Neither touches per-user Gmail autocomplete
 *   ("Other contacts"), which is not centrally removable. See PCR_CHANGELOG.md.
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
 * - groupAdministration_stageOrphanedSquadronGroups(sheetName)
 *   Finds existing squadron groups whose list type is now DISABLED for this
 *   tenant (via SQUADRON_DISTRIBUTION_TOGGLES) and writes them to a worklist tab
 *   ("Delete Groups" by default) for review. Does NOT delete — feed the reviewed
 *   sheet to groupAdministration_bulkDeleteGroupsFromSheet().
 *
 * - groupAdministration_stageLegacyDlGroups(prefix, sheetName)
 *   READ-ONLY. Inventories live Groups whose primary address or an alias starts
 *   with a legacy prefix (default "dl-cawg") to a review tab ("Legacy DL
 *   Cleanup"), split into PRIMARY (delete the group) vs ALIAS (remove only the
 *   alias) so the two are not conflated. Does NOT clear per-user autocomplete.
 *
 * - groupAdministration_resolveLegacyAddress(email)
 *   READ-ONLY. Says definitively whether one address is a live group's own
 *   address, an alias on a current group, or not a live directory object at all
 *   (in which case any lingering autocomplete is per-user "Other contacts").
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
 * Stages "orphaned" managed squadron groups for review before deletion.
 *
 * Now that SQUADRON_DISTRIBUTION_TOGGLES is tenant-driven, some list types are no
 * longer managed on a given tenant (e.g. on cadets: .all, .seniors, and the
 * command-staff lists). Disabling a toggle stops managing those lists but does
 * NOT delete the already-created groups, leaving orphans with stale membership.
 *
 * This finds existing groups whose email suffix matches a managed squadron list
 * type whose toggle is currently DISABLED for this tenant, and writes them to the
 * "Delete Groups" worklist tab (first column header "group") for human review.
 * It only reads groups and writes the sheet — it does NOT delete anything. After
 * reviewing/trimming the sheet, run groupAdministration_bulkDeleteGroupsFromSheet().
 *
 * Tenant-aware: it reads THIS tenant's profile toggles (via
 * isSquadronDistributionListEnabled_), so a list type that is still enabled here
 * is treated as managed, not orphaned. Run it on the tenant you want to clean
 * (e.g. cadets). Note: this overwrites the target tab's contents.
 *
 * @param {string=} sheetName Target worklist tab (default "Delete Groups").
 * @returns {{staged:number, sheet:string, groups:Array<Object>}}
 */
function groupAdministration_stageOrphanedSquadronGroups(sheetName) {
  if (typeof isSquadronDistributionListEnabled_ !== 'function') {
    throw new Error('isSquadronDistributionListEnabled_ (SquadronGroups.gs) is required');
  }

  const targetSheetName = String(sheetName || 'Delete Groups').trim();

  // Suffixes the squadron-group automation manages. A group is an orphan
  // candidate when its suffix is one of these AND that suffix is currently
  // disabled by this tenant's toggles.
  const managedSuffixSet = {};
  [
    'all', 'allhands', 'cadets', 'seniors', 'parents',
    'commander', 'deputy-commander', 'deputy-commander-cadets', 'deputy-commander-seniors'
  ].forEach(s => { managedSuffixSet[s] = true; });

  // Managed squadron prefixes: the wing code, optionally + a 3-digit unit
  // (wing-level "ca", group/unit-level "ca###").
  const wing = String((CONFIG && CONFIG.WING) || '').trim().toLowerCase();
  if (!wing) throw new Error('CONFIG.WING is not set');
  const prefixRe = new RegExp('^' + wing.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?:\\d{3})?$');

  const orphans = [];
  const groups = groupAdministration_listGroups();

  for (let i = 0; i < groups.length; i++) {
    const email = String(groups[i].email || '').toLowerCase();
    const local = email.split('@')[0];
    const dot = local.indexOf('.');
    if (dot < 0) continue; // no suffix (bare unit / public-contact) — skip

    const prefix = local.slice(0, dot);
    const suffix = local.slice(dot + 1);

    if (!prefixRe.test(prefix)) continue;
    if (!managedSuffixSet[suffix]) continue;
    if (isSquadronDistributionListEnabled_(suffix)) continue; // still managed → not an orphan

    orphans.push({
      group: email,
      name: String(groups[i].name || ''),
      suffix: suffix,
      directMembersCount: Number(groups[i].directMembersCount || 0),
      reason: 'Squadron list type "' + suffix + '" is disabled for this tenant'
    });
  }

  orphans.sort((a, b) => a.group.localeCompare(b.group));

  const ss = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID);
  const sheet = getOrCreateSheet_(ss, targetSheetName);
  const rows = [['group', 'name', 'suffix', 'directMembersCount', 'reason']];
  for (let i = 0; i < orphans.length; i++) {
    rows.push([
      orphans[i].group,
      orphans[i].name,
      orphans[i].suffix,
      orphans[i].directMembersCount,
      orphans[i].reason
    ]);
  }
  writeTabularData_(sheet, rows);

  Logger.info('Staged orphaned squadron groups for review', {
    sheet: targetSheetName,
    staged: orphans.length,
    sample: orphans.slice(0, 20)
  });

  return { staged: orphans.length, sheet: targetSheetName, groups: orphans };
}

/**
 * READ-ONLY. Inventories live directory Groups whose primary address OR any
 * alias begins with a legacy prefix (default 'dl-cawg'), writing them to a
 * review worklist for a human to triage before any deletion.
 *
 * WHY: after the M365 -> Google migration, distribution lists were recreated
 * with verbose 'DL-CAWG-...' names; the current automation manages the same
 * lists under the modern 'ca###.all' convention. The legacy names linger either
 * as duplicate groups or as aliases on the modern group, cluttering the GAL and
 * re-seeding users' Gmail autocomplete each time someone mails them.
 *
 * WHAT THIS DOES NOT DO: it changes nothing, and it does NOT clear the addresses
 * from anyone's Gmail autocomplete / "Other contacts" — those are per-user and
 * not centrally removable (see groupAdministration_deleteUserContactsForAllUsers_notSupported).
 * Deleting a live directory object stops it re-seeding autocomplete and removes
 * GAL clutter; it does not retroactively scrub existing per-user suggestions.
 *
 * The two match types need DIFFERENT remediation — do not conflate them:
 *   - PRIMARY : the legacy name is the group's OWN address. Safe to delete the
 *               group (feed the primary rows to
 *               groupAdministration_bulkDeleteGroupsFromSheet) — but confirm it
 *               is an unused duplicate first (check members / whether people
 *               still send to it).
 *   - ALIAS   : the legacy name is only an ALIAS on a still-current group (e.g.
 *               dl-cawg-...-110-all aliased onto ca110.all). Do NOT delete the
 *               group; remove just the alias with
 *               AdminDirectory.Groups.Aliases.remove(groupEmail, legacyAddress).
 *
 * Alias detection relies on aliases returned by Groups.list. For a definitive
 * check of one address, use groupAdministration_resolveLegacyAddress().
 *
 * @param {string=} prefix    Legacy local-part prefix to match (default 'dl-cawg').
 * @param {string=} sheetName Target review tab (default 'Legacy DL Cleanup').
 * @returns {{prefix:string, scanned:number, primary:number, alias:number, sheet:string, matches:Array<Object>}}
 */
function groupAdministration_stageLegacyDlGroups(prefix, sheetName) {
  const wantPrefix = String(prefix || 'dl-cawg').trim().toLowerCase();
  if (!wantPrefix) throw new Error('prefix must be a non-empty string');
  const targetSheetName = String(sheetName || 'Legacy DL Cleanup').trim();

  const matches = [];
  let scanned = 0;
  let pageToken = '';

  do {
    const res = executeWithRetry(() => AdminDirectory.Groups.list({
      customer: CONFIG.CUSTOMER_ID || 'my_customer',
      maxResults: 200,
      pageToken: pageToken
    }));

    const groups = res.groups || [];
    for (let i = 0; i < groups.length; i++) {
      scanned++;
      const g = groups[i];
      const email = String(g.email || '').toLowerCase();
      const base = {
        name: String(g.name || ''),
        directMembersCount: Number(g.directMembersCount || 0),
        adminCreated: String(g.adminCreated || ''),
        id: String(g.id || '')
      };

      if (email.split('@')[0].startsWith(wantPrefix)) {
        matches.push(Object.assign({
          matchType: 'PRIMARY',
          legacyAddress: email,
          groupEmail: email
        }, base));
      }

      const nonEditable = (g.nonEditableAliases || []).map(a => String(a).toLowerCase());
      const aliases = [].concat(g.aliases || [], g.nonEditableAliases || []);
      for (let a = 0; a < aliases.length; a++) {
        const alias = String(aliases[a] || '').toLowerCase();
        if (!alias || !alias.split('@')[0].startsWith(wantPrefix)) continue;
        matches.push(Object.assign({
          matchType: nonEditable.indexOf(alias) > -1 ? 'ALIAS (non-editable)' : 'ALIAS',
          legacyAddress: alias,
          groupEmail: email
        }, base));
      }
    }

    pageToken = res.nextPageToken || '';
  } while (pageToken);

  matches.sort((a, b) => a.legacyAddress.localeCompare(b.legacyAddress));

  const ss = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID);
  const sheet = getOrCreateSheet_(ss, targetSheetName);
  const rows = [['legacyAddress', 'matchType', 'groupEmail', 'name', 'directMembersCount', 'adminCreated', 'groupId']];
  for (let i = 0; i < matches.length; i++) {
    const m = matches[i];
    rows.push([m.legacyAddress, m.matchType, m.groupEmail, m.name, m.directMembersCount, m.adminCreated, m.id]);
  }
  writeTabularData_(sheet, rows);

  const primary = matches.filter(m => m.matchType === 'PRIMARY').length;
  const alias = matches.length - primary;

  Logger.info('Legacy DL group scan complete', {
    prefix: wantPrefix,
    scanned: scanned,
    matched: matches.length,
    primary: primary,
    alias: alias,
    sheet: targetSheetName,
    sample: matches.slice(0, 20)
  });

  return { prefix: wantPrefix, scanned: scanned, primary: primary, alias: alias, sheet: targetSheetName, matches: matches };
}

/**
 * READ-ONLY. Definitively resolves ONE legacy address: is it a live group's own
 * address, an alias on a current group, or not a live directory object at all?
 * Uses Groups.get, which resolves both primary emails and aliases, so it is the
 * authoritative spot-check the bulk scan's list-based alias detection is not.
 *
 * A 'NOT_A_GROUP' result means the address is not a live Group or group alias —
 * so if it still autocompletes for a user, it is a per-user contact / "Other
 * contacts" entry, which is not centrally removable.
 *
 * @param {string=} email Address to resolve; falls back to GROUP_EMAIL run input.
 * @returns {Object} resolution incl. kind and remediation guidance.
 */
function groupAdministration_resolveLegacyAddress(email) {
  const target = sanitizeEmail(
    String(email || (GROUP_ADMINISTRATION_RUN_INPUTS && GROUP_ADMINISTRATION_RUN_INPUTS.GROUP_EMAIL) || '').trim()
  );
  if (!target) throw new Error('Pass an email, or set GROUP_ADMINISTRATION_RUN_INPUTS.GROUP_EMAIL.');

  let group;
  try {
    group = executeWithRetry(() => AdminDirectory.Groups.get(target));
  } catch (e) {
    if (e.details && e.details.code === 404) {
      const miss = {
        address: target,
        live: false,
        kind: 'NOT_A_GROUP',
        remediation: 'No live Group or group alias resolves this address. If it still autocompletes, it is a per-user contact / "Other contacts" entry — remove it via the user\'s Gmail autocomplete (hover the suggestion, click the X) or contacts.google.com; it cannot be cleared centrally.'
      };
      Logger.info('Resolved legacy address', miss);
      return miss;
    }
    throw e;
  }

  const primary = String(group.email || '').toLowerCase();
  const nonEditable = (group.nonEditableAliases || []).map(a => String(a).toLowerCase());
  const isPrimary = primary === target;
  const isNonEditable = nonEditable.indexOf(target) > -1;

  const result = {
    address: target,
    live: true,
    kind: isPrimary ? 'GROUP_PRIMARY' : (isNonEditable ? 'GROUP_ALIAS_NONEDITABLE' : 'GROUP_ALIAS'),
    groupEmail: primary,
    groupName: String(group.name || ''),
    directMembersCount: Number(group.directMembersCount || 0),
    remediation: isPrimary
      ? 'Legacy name is the group\'s OWN address. If it is an unused duplicate of the modern list, delete the group (groupAdministration_deleteGroup) — check members/usage first.'
      : (isNonEditable
        ? 'Non-editable alias (derived, e.g. from a secondary domain). Cannot be removed directly; it clears when its source is removed.'
        : 'Legacy name is an ALIAS on a still-current group. Remove ONLY the alias: AdminDirectory.Groups.Aliases.remove("' + primary + '", "' + target + '"). Do NOT delete the group.')
  };

  Logger.info('Resolved legacy address', result);
  return result;
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
