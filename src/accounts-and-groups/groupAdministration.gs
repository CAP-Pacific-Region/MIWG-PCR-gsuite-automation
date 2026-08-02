/*******************************************************
 * Group Administration Utilities
 *
 * Filename: groupAdministration.gs
 * Saved: 2026-07-31
 * Changes: Added the two halves of the cross-tenant fan-out check —
 *   groupAdministration_diagnoseWingAllFanout(groupEmail) on the SENDING (wing)
 *   tenant, for the four prerequisites the squadron sync handles for unit lists
 *   and nobody handles for the wing list, and
 *   groupAdministration_diagnoseReceiveGroup(groupEmail) on the RECEIVING tenant,
 *   for the three neither the wing tenant nor a Workspace API can see from the
 *   other side. Both read-only. See PCR_CHANGELOG.md.
 *   Added groupAdministration_repairReceiveListPosting(dryRun) — the
 *   repair counterpart to the receive-list audit, for the groups no sync owns
 *   (wing all-hands, group-HQ lists, lists outliving their unit). DRY RUN by
 *   default; settings only. See PCR_CHANGELOG.md.
 *   Added groupAdministration_stageLegacyDlGroups() (read-only bulk scan
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
 *   groups (e.g. the cadets tenant) to find lists that would silently reject
 *   a sender on the other tenant — both a person writing across directly and
 *   cross-tenant fan-out from a wing .all list.
 *
 * - groupAdministration_diagnoseWingAllFanout(groupEmail)
 *   READ-ONLY. Run on the WING tenant. Explains why a wing-wide ".all" list is
 *   or is not delivering to the cadet tenant: whether a cadet-tenant group is
 *   nested at all, whether the "User Additions" tab preserves it (the delta pass
 *   removes any plain MEMBER it did not compute), whether the "Groups" row sets
 *   Add EXT (without it, external adds are declined and allowExternalMembers is
 *   written false every run), and the group's live settings. The cadet-side
 *   posting check is groupAdministration_auditReceiveListPosting() on the
 *   CADETS tenant.
 *
 * - groupAdministration_diagnoseReceiveGroup(groupEmail)
 *   READ-ONLY. The receiving half of the check above, run on the tenant that OWNS
 *   the group. Says whether one address is fit to be nested into a list on the
 *   other tenant: it exists, it has members, and whoCanPostMessage is
 *   ANYONE_CAN_POST. None of the three is visible from the sending tenant.
 *
 * - groupAdministration_repairReceiveListPosting(dryRun)
 *   Fixes what that audit flags. DRY RUN by default. For the lists no sync
 *   owns — the wing all-hands, group-headquarters lists, and groups outliving
 *   their unit — which updateAllSquadronGroups() will never reach because it
 *   only iterates UNIT-scope orgs. Settings only; never changes membership.
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
 * The tenants are separate Workspace accounts, so a sender on the other one is
 * external no matter that they are the same wing. A group set to
 * ALL_MEMBERS_CAN_POST or ALL_IN_DOMAIN_CAN_POST therefore rejects/holds both a
 * person writing across directly (a senior on the wing domain mailing
 * ca.all@cawgcadets.org) and cross-tenant fan-out (wing ca###.all -> cadet
 * ca###.cadets@cawgcadets.org), which arrives carrying the original sender.
 * Either way the group needs ANYONE_CAN_POST, paired with spam moderation.
 *
 * SquadronGroups.gs 1.6.0 reconciles both settings on every
 * updateAllSquadronGroups() run, so this is a before/after check on that.
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
 * Repairs the posting permissions the audit above flags. DRY RUN by default.
 *
 * updateAllSquadronGroups() reconciles these settings (SquadronGroups.gs 1.6.0),
 * but only for lists it owns, and it reaches a list only through
 * shouldCreateDistributionLists() — which returns false for any org that is not
 * UNIT scope. So three populations exist that no sync will ever visit, however
 * many times it is run:
 *
 *   1. The wing all-hands (ca.all@...). No CAPWATCH org at all; the sync
 *      iterates orgs, and wing scope is not one of them.
 *   2. Group-headquarters lists (ca006.all, ca006.dty.all, ...). scope=GROUP,
 *      excluded by the UNIT filter in updateAllSquadronGroups().
 *   3. Lists left behind by units that no longer appear in this tenant's
 *      CAPWATCH — the group outlived the org that justified it.
 *
 * Those are exactly the groups a member notices, because the wing all-hands is
 * the one people actually write to. This closes them without inventing
 * membership semantics for orgs the sync does not model: it changes posting
 * policy only, never who is in a group.
 *
 * Scope is whatever groupAdministration_auditReceiveListPosting() flags, so the
 * audit stays the single definition of "should accept outside mail" and the two
 * cannot disagree. Re-running is harmless — a group already correct is not
 * flagged, so there is nothing left to patch.
 *
 *   groupAdministration_repairReceiveListPosting()        // preview, writes nothing
 *   groupAdministration_repairReceiveListPosting(false)   // apply
 *
 * @param {boolean} [dryRun=true] - false to actually write.
 * @returns {{examined:number, repaired:Array<Object>, failed:Array<Object>, dryRun:boolean}}
 */
function groupAdministration_repairReceiveListPosting(dryRun) {
  const isDryRun = (dryRun !== false);

  if (typeof AdminGroupsSettings === 'undefined' || !AdminGroupsSettings.Groups || !AdminGroupsSettings.Groups.patch) {
    throw new Error('AdminGroupsSettings advanced service is not enabled');
  }

  // The same three keys SquadronGroups.applyGroupSettings() manages. Posting is
  // opened to the internet, so moderation travels with it — see that function.
  const desired = {
    whoCanPostMessage: 'ANYONE_CAN_POST',
    allowExternalMembers: 'true',
    spamModerationLevel: 'MODERATE'
  };

  const blocking = groupAdministration_auditReceiveListPosting().blocking;
  const repaired = [];
  const failed = [];

  for (let i = 0; i < blocking.length; i++) {
    const row = blocking[i];
    const patch = {};
    for (const key in desired) {
      if (String(row[key] || '') !== desired[key]) patch[key] = desired[key];
    }

    if (Object.keys(patch).length === 0) continue;

    if (isDryRun) {
      Logger.info('💡 [Dry-Run] Would repair receive-list posting', {
        group: row.email,
        patch: patch
      });
      repaired.push({ email: row.email, patch: patch });
      continue;
    }

    try {
      executeWithRetry(() => AdminGroupsSettings.Groups.patch(patch, row.email));
      repaired.push({ email: row.email, patch: patch });
      Logger.info('Receive-list posting repaired', { group: row.email, patch: patch });
    } catch (e) {
      failed.push({ email: row.email, error: e.message });
      Logger.warn('Failed to repair receive-list posting', {
        group: row.email,
        errorMessage: e.message
      });
    }
  }

  Logger.info('Receive-list posting repair complete', {
    dryRun: isDryRun,
    examined: blocking.length,
    repairedCount: repaired.length,
    failedCount: failed.length
  });

  console.log(isDryRun
    ? `💡 Dry run — ${repaired.length} of ${blocking.length} would be repaired. ` +
      'Run groupAdministration_repairReceiveListPosting(false) to apply.'
    : `✅ Repaired ${repaired.length} of ${blocking.length}; ${failed.length} failed.`);

  return { examined: blocking.length, repaired: repaired, failed: failed, dryRun: isDryRun };
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
 * READ-ONLY. Explains why a wing-wide ".all" list is (or is not) delivering to
 * the cadet tenant. Run it on the WING tenant.
 *
 * The unit lists and the wing list reach cadets by different code, which is why
 * "ca###.all works but ca.all does not" is a normal state rather than a
 * contradiction:
 *
 *   ca###.all  — updateAllSquadronGroups() (SquadronGroups.gs) nests the cadet
 *                group itself and reconciles the group settings that let it.
 *                It only iterates UNIT-scope orgs.
 *   ca.all     — updateEmailGroups() (UpdateGroups.gs) only. No CAPWATCH org is
 *                wing scope, so SquadronGroups never touches it.
 *
 * So every prerequisite the squadron sync handles for a unit list has to be
 * satisfied by the Groups / User Additions tabs for the wing list, and each one
 * fails silently on its own:
 *
 *   1. The nested cadet group has to BE a member.
 *   2. It has to survive the delta pass: updateEmailGroups() removes any plain
 *      MEMBER it did not compute as desired, and the only way a cross-tenant
 *      address becomes desired is a "User Additions" row naming this group.
 *   3. The Groups row has to permit external members ("Add EXT", or "Add Lite"
 *      which implies it). Without it the apply loop declines the add AND
 *      applyManagedGroupSettings_() writes allowExternalMembers=false on every
 *      run, so a hand-fixed flag reverts within a day.
 *   4. The group's live allowExternalMembers has to be true.
 *
 * It checks all four against live Google state plus the sheet that drives the
 * next run, and names which one is failing. Cadet-side posting permission — the
 * fifth prerequisite, and the only one this tenant cannot see — is covered by
 * groupAdministration_auditReceiveListPosting() run on the CADETS tenant.
 *
 * Changes nothing.
 *
 *   groupAdministration_diagnoseWingAllFanout()
 *   groupAdministration_diagnoseWingAllFanout('ca.all@cawgcap.org')
 *
 * @param {string=} groupEmail - Defaults to "<wing>.all" on this tenant's domain.
 * @returns {{group:string, exists:boolean, findings:Array<string>,
 *            peerMembers:Array<Object>, settings:Object, sheet:Object}}
 */
function groupAdministration_diagnoseWingAllFanout(groupEmail) {
  const target = String(
    groupEmail || (String(CONFIG.WING || '').toLowerCase() + '.all' + CONFIG.EMAIL_DOMAIN)
  ).trim().toLowerCase();

  const groupId = target.split('@')[0];
  const baseName = groupId.indexOf('.') > -1
    ? groupId.split('.').slice(1).join('.')
    : groupId;

  const peerDomain = String(
    (typeof getCAWGCadetsTenantDomain_ === 'function' ? getCAWGCadetsTenantDomain_() : '') ||
    CONFIG.CADETS_TENANT_DOMAIN || ''
  ).trim().toLowerCase();

  const findings = [];
  const result = {
    group: target,
    exists: false,
    findings: findings,
    peerMembers: [],
    settings: {},
    sheet: { groupsRows: [], userAdditionsRows: [] }
  };

  console.log('\n' + '='.repeat(80));
  console.log('Cross-tenant fan-out diagnosis: ' + target);
  console.log('Cadet tenant domain: ' + (peerDomain || '(not configured)'));
  console.log('='.repeat(80));

  // --- 1. Does the group exist, and what does Google think of it right now? ---
  let group;
  try {
    group = executeWithRetry(() => AdminDirectory.Groups.get(target));
    result.exists = true;
  } catch (e) {
    findings.push('BLOCKER: group does not exist on this tenant (' + e.message + ').');
    console.log('\n✗ Group not found.');
    Logger.warn('Wing .all fan-out diagnosis: group not found', { group: target, errorMessage: e.message });
    return result;
  }

  console.log('\nGroup: ' + group.name + '   direct members: ' + Number(group.directMembersCount || 0));

  if (typeof AdminGroupsSettings !== 'undefined' && AdminGroupsSettings.Groups && AdminGroupsSettings.Groups.get) {
    try {
      const s = executeWithRetry(() => AdminGroupsSettings.Groups.get(target));
      result.settings = {
        allowExternalMembers: String(s.allowExternalMembers || ''),
        whoCanPostMessage: String(s.whoCanPostMessage || ''),
        messageModerationLevel: String(s.messageModerationLevel || ''),
        spamModerationLevel: String(s.spamModerationLevel || '')
      };
      console.log('Settings: ' + JSON.stringify(result.settings));

      if (result.settings.allowExternalMembers !== 'true') {
        findings.push('BLOCKER: allowExternalMembers is "' + result.settings.allowExternalMembers +
          '". A cadet-tenant group cannot be added while this is false.');
      }
    } catch (e) {
      console.log('Settings: unavailable (' + e.message + ')');
    }
  } else {
    console.log('Settings: AdminGroupsSettings advanced service is not enabled.');
  }

  // --- 2. Is a cadet-tenant group actually nested? ---
  let members = [];
  let pageToken = '';
  do {
    const page = executeWithRetry(() => AdminDirectory.Members.list(target, {
      maxResults: 200,
      pageToken: pageToken
    }));
    if (page.members) members = members.concat(page.members);
    pageToken = page.nextPageToken;
  } while (pageToken);

  const peerMembers = members
    .filter(m => peerDomain && String(m.email || '').toLowerCase().endsWith('@' + peerDomain))
    .map(m => ({
      email: String(m.email || '').toLowerCase(),
      role: String(m.role || 'MEMBER').toUpperCase(),
      type: String(m.type || ''),
      status: String(m.status || ''),
      // A member set to NONE/DISABLED is in the group and still receives nothing.
      deliverySettings: String(m.delivery_settings || m.deliverySettings || '')
    }));
  result.peerMembers = peerMembers;

  console.log('\nMembers on ' + (peerDomain || 'the cadet domain') + ': ' + peerMembers.length);
  peerMembers.forEach(m => console.log('  ' + m.email + '   role=' + m.role +
    ' type=' + m.type + ' status=' + m.status +
    (m.deliverySettings ? ' delivery=' + m.deliverySettings : '')));

  if (!peerMembers.length) {
    findings.push('BLOCKER: no address on the cadet tenant is a member of this group, ' +
      'so nothing sent here can reach a cadet.');
  }
  peerMembers.forEach(m => {
    if (m.status && m.status.toUpperCase() !== 'ACTIVE') {
      findings.push('BLOCKER: ' + m.email + ' is a member but its status is ' + m.status + '.');
    }
    const delivery = m.deliverySettings.toUpperCase();
    if (delivery === 'NONE' || delivery === 'DISABLED') {
      findings.push('BLOCKER: ' + m.email + ' is a member with delivery ' + delivery +
        ' — it is listed but is sent nothing.');
    }
  });

  // --- 3. What will the next updateEmailGroups() run do to it? ---
  const ss = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID);
  const groupsSheet = ss.getSheetByName('Groups');
  let externalAllowedBySheet = false;

  if (!groupsSheet) {
    findings.push('Groups tab not found; cannot predict the next run.');
  } else {
    const rows = groupsSheet.getDataRange().getValues();
    const header = (rows[0] || []).map(h => String(h || '').trim().toLowerCase());
    const idxAddExt = header.indexOf('add ext');
    const idxAddLite = header.indexOf('add lite');
    const truthy = v => ['y', 'yes', 'x', 'true'].indexOf(String(v || '').trim().toLowerCase()) > -1;

    for (let r = 1; r < rows.length; r++) {
      if (String(rows[r][1] || '').trim().toLowerCase() !== baseName.toLowerCase()) continue;
      const addExt = idxAddExt > -1 ? String(rows[r][idxAddExt] || '') : '';
      const addLite = idxAddLite > -1 ? String(rows[r][idxAddLite] || '') : '';
      const allows = truthy(addExt) || truthy(addLite);
      externalAllowedBySheet = externalAllowedBySheet || allows;
      result.sheet.groupsRows.push({
        row: r + 1,
        groupName: String(rows[r][1] || ''),
        attribute: String(rows[r][2] || ''),
        addExt: addExt,
        addLite: addLite,
        allowsExternal: allows
      });
    }

    console.log('\nGroups tab rows named "' + baseName + '": ' + result.sheet.groupsRows.length);
    result.sheet.groupsRows.forEach(r => console.log('  row ' + r.row + '   Attribute=' +
      (r.attribute || '(blank)') + '   Add EXT=' + (r.addExt || '(blank)') +
      '   Add Lite=' + (r.addLite || '(blank)')));

    if (!result.sheet.groupsRows.length) {
      findings.push('No Groups row is named "' + baseName + '", so updateEmailGroups() does not ' +
        'manage this group and a User Additions row for it has nothing to merge into.');
    } else if (!externalAllowedBySheet) {
      findings.push('BLOCKER: no Groups row for "' + baseName + '" sets Add EXT (or Add Lite). ' +
        'Every run therefore declines external adds AND writes allowExternalMembers=false ' +
        'onto this group, so fixing the flag by hand reverts on the next sync. ' +
        'Set Add EXT = Y on the "' + baseName + '" row.');
    }
  }

  // --- 4. Is the nesting preserved, or does the delta pass drop it? ---
  const additionsSheet = ss.getSheetByName('User Additions') ||
    ss.getSheetByName('UserAdditions') || ss.getSheetByName('USER ADDITIONS');
  const preserved = {};

  if (!additionsSheet) {
    findings.push('User Additions tab not found; nothing preserves a cross-tenant member here.');
  } else {
    const rows = additionsSheet.getDataRange().getValues();
    const header = (rows[0] || []).map(h => String(h || '').trim().toLowerCase());
    const idxEmail = header.indexOf('email') > -1 ? header.indexOf('email') : 1;
    const idxGroups = header.indexOf('groups') > -1 ? header.indexOf('groups') : 3;

    for (let r = 1; r < rows.length; r++) {
      const email = String(rows[r][idxEmail] || '').trim().toLowerCase();
      if (!email) continue;
      const tokens = String(rows[r][idxGroups] || '').split(',')
        .map(t => t.trim().toLowerCase())
        .filter(Boolean)
        .map(t => t.endsWith(CONFIG.EMAIL_DOMAIN.toLowerCase())
          ? t.slice(0, -CONFIG.EMAIL_DOMAIN.length)
          : t);
      if (tokens.indexOf(groupId) < 0) continue;

      preserved[email] = true;
      // updateCAWGCadetGroups() rewrites every row whose address matches its own
      // managed pattern. A hand-added one that it does not generate is deleted
      // from the tab on its next run, and the member disappears one sync later.
      const generated = typeof isManagedCAWGCadetGroupEmail_ === 'function' &&
        isManagedCAWGCadetGroupEmail_(email, peerDomain);
      const isGeneratedShape = /\.(cadets|parents)@/.test(email);
      result.sheet.userAdditionsRows.push({
        row: r + 1,
        email: email,
        rewrittenByCadetGroupSync: !!generated && !isGeneratedShape
      });
    }

    console.log('\nUser Additions rows targeting ' + groupId + ': ' + result.sheet.userAdditionsRows.length);
    result.sheet.userAdditionsRows.forEach(r => console.log('  row ' + r.row + '   ' + r.email +
      (r.rewrittenByCadetGroupSync ? '   ⚠ updateCAWGCadetGroups() deletes this row' : '')));

    result.sheet.userAdditionsRows.forEach(r => {
      if (r.rewrittenByCadetGroupSync) {
        findings.push('BLOCKER: the User Additions row for ' + r.email + ' matches the addresses ' +
          'updateCAWGCadetGroups() manages but is not one it generates, so that function deletes ' +
          'the row and the next updateEmailGroups() run then removes the member. It generates ' +
          '.cadets@ and .parents@ addresses only — nest ' +
          String(CONFIG.WING || '').toLowerCase() + '.cadets@' + peerDomain + ' instead.');
      }
    });
  }

  peerMembers.forEach(m => {
    if (m.role !== 'MEMBER') return;   // MANAGER/OWNER are never auto-removed
    if (preserved[m.email]) return;
    findings.push('BLOCKER: ' + m.email + ' is a member but no User Additions row names ' + groupId +
      ', so the next updateEmailGroups() run computes it as unwanted and removes it. ' +
      'Add a row: Email=' + m.email + ', Groups=' + groupId + '.');
  });

  // --- Verdict ---
  console.log('\n' + '-'.repeat(80));
  if (!findings.length) {
    console.log('Nothing on this tenant blocks the fan-out.');
    console.log('Remaining prerequisite is cadet-side: run ' +
      'groupAdministration_auditReceiveListPosting() on the CADETS tenant and confirm the ' +
      'nested group is ANYONE_CAN_POST. Anything else is best proven by Admin console → ' +
      'Reporting → Email Log Search on a real message to ' + target + '.');
  } else {
    findings.forEach((f, i) => console.log((i + 1) + '. ' + f));
  }
  console.log('-'.repeat(80) + '\n');

  Logger.info('Wing .all fan-out diagnosis complete', {
    group: target,
    peerMembers: peerMembers.length,
    findings: findings
  });

  return result;
}

/**
 * READ-ONLY. The receiving half of the fan-out check: is this group fit to be
 * nested into a list on the OTHER tenant? Run it on the tenant that OWNS the
 * group (for ca.cadets@cawgcadets.org, the cadets tenant).
 *
 * groupAdministration_diagnoseWingAllFanout() answers everything visible from the
 * sending tenant and nothing visible from this one, because a Workspace tenant
 * cannot read the other's group settings. Three things decide whether mail that
 * arrives here goes anywhere, and all three are invisible from the wing side:
 *
 *   1. The group exists. Nesting an address that resolves to nothing is accepted
 *      by the sending group and delivers to no one.
 *   2. It has members. An empty receive list is indistinguishable from a working
 *      one until somebody asks a cadet whether they got the mail.
 *   3. whoCanPostMessage is ANYONE_CAN_POST. Fan-out arrives carrying the
 *      ORIGINAL sender, who is on the other tenant and therefore external —
 *      ALL_IN_DOMAIN_CAN_POST and ALL_MEMBERS_CAN_POST both reject it. This is
 *      the one that fails without a bounce anybody notices.
 *
 * The bulk equivalent is groupAdministration_auditReceiveListPosting(), which
 * covers item 3 for every managed list at once. Use this when a specific address
 * is about to be nested, or has been and is not delivering.
 *
 *   groupAdministration_diagnoseReceiveGroup('ca.cadets@cawgcadets.org')
 *
 * @param {string=} groupEmail - Defaults to GROUP_ADMINISTRATION_RUN_INPUTS.GROUP_EMAIL.
 * @returns {{group:string, exists:boolean, directMembersCount:number,
 *            settings:Object, findings:Array<string>}}
 */
function groupAdministration_diagnoseReceiveGroup(groupEmail) {
  const target = getConfiguredRunEmail_(groupEmail, 'GROUP_EMAIL', 'group email');
  const findings = [];
  const result = {
    group: target,
    exists: false,
    directMembersCount: 0,
    settings: {},
    findings: findings
  };

  console.log('\n' + '='.repeat(80));
  console.log('Receive-group check: ' + target);
  console.log('='.repeat(80));

  let group;
  try {
    group = executeWithRetry(() => AdminDirectory.Groups.get(target));
    result.exists = true;
  } catch (e) {
    findings.push('BLOCKER: this group does not exist on this tenant. Nesting it into a ' +
      'list on the other tenant is accepted and delivers to nobody.');
    console.log('\n✗ Not found (' + e.message + ')');
    console.log('-'.repeat(80) + '\n');
    Logger.warn('Receive-group check: group not found', { group: target, errorMessage: e.message });
    return result;
  }

  result.directMembersCount = Number(group.directMembersCount || 0);
  console.log('\nName: ' + group.name);
  console.log('Direct members: ' + result.directMembersCount);

  if (!result.directMembersCount) {
    findings.push('BLOCKER: the group exists but has no members, so mail fanned out to it ' +
      'reaches nobody. An empty receive list looks exactly like a working one from ' +
      'the sending tenant.');
  }

  if (typeof AdminGroupsSettings !== 'undefined' && AdminGroupsSettings.Groups && AdminGroupsSettings.Groups.get) {
    try {
      const s = executeWithRetry(() => AdminGroupsSettings.Groups.get(target));
      result.settings = {
        whoCanPostMessage: String(s.whoCanPostMessage || ''),
        allowExternalMembers: String(s.allowExternalMembers || ''),
        messageModerationLevel: String(s.messageModerationLevel || ''),
        spamModerationLevel: String(s.spamModerationLevel || '')
      };
      console.log('Settings: ' + JSON.stringify(result.settings));

      if (result.settings.whoCanPostMessage !== 'ANYONE_CAN_POST') {
        findings.push('BLOCKER: whoCanPostMessage is ' + result.settings.whoCanPostMessage +
          '. Fan-out from the other tenant carries the original — external — sender, ' +
          'so this group rejects or holds it. Only ANYONE_CAN_POST accepts it; Google has ' +
          'no value meaning "members plus my other tenant". Repair with ' +
          'groupAdministration_repairReceiveListPosting(false) on this tenant.');
      }
      if (result.settings.whoCanPostMessage === 'ANYONE_CAN_POST' &&
          result.settings.spamModerationLevel !== 'MODERATE') {
        findings.push('ANYONE_CAN_POST is open to the internet and spamModerationLevel is "' +
          result.settings.spamModerationLevel + '". The two are meant to travel together.');
      }
    } catch (e) {
      console.log('Settings: unavailable (' + e.message + ')');
      findings.push('Could not read settings, so posting permission is unverified: ' + e.message);
    }
  } else {
    findings.push('AdminGroupsSettings advanced service is not enabled, so posting ' +
      'permission — the prerequisite that fails silently — could not be checked.');
  }

  console.log('\n' + '-'.repeat(80));
  if (!findings.length) {
    console.log('Fit to receive cross-tenant fan-out.');
  } else {
    findings.forEach((f, i) => console.log((i + 1) + '. ' + f));
  }
  console.log('-'.repeat(80) + '\n');

  Logger.info('Receive-group check complete', {
    group: target,
    exists: result.exists,
    directMembersCount: result.directMembersCount,
    findings: findings
  });

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
