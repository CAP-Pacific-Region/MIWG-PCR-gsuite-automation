/**
 * -------------------------------------------------------------------------
 * UpdateRegionGroupChats.gs
 *
 * Creates/updates:
 *  - Google Groups
 *  - Google Chat Spaces
 *
 * Based on Automation spreadsheet "Groups" tab rows where:
 *  Attribute == dutyPositionIdsRegion
 *
 * CAPWATCH source:
 *  Uses CONFIG.REGION_CAPWATCH_DATA_FOLDER_ID (PCR all-members CAPWATCH folder)
 *
 * Membership rule for dutyPositionIdsRegion:
 *  Include members whose DutyPosition Duty matches configured Values AND
 *  DutyPosition ORGID is in scope REGION or WING (from Organization table).
 *
 * Version: 1.0.0
 * Date: 2026-07-09
 * Changes: Folded into the shared src/ from the Pacific Region project; gated
 *   behind PROFILE_.RUN_REGION_GROUP_CHATS (on for pacific, off for wing).
 *   See PCR_CHANGELOG.md.
 * -------------------------------------------------------------------------
 */

function updateRegionGroupChats() {
  if (!PROFILE_.RUN_REGION_GROUP_CHATS) {
    Logger.info('updateRegionGroupChats skipped (RUN_REGION_GROUP_CHATS=false for this tenant profile)');
    return;
  }
  const start = new Date();

  Logger.info('Starting updateRegionGroupChats', {
    regionCapwatchFolderId: CONFIG.REGION_CAPWATCH_DATA_FOLDER_ID,
    automationSpreadsheetId: CONFIG.AUTOMATION_SPREADSHEET_ID
  });

  const orgs = getSquadronsForRegion_(CONFIG.REGION_CAPWATCH_DATA_FOLDER_ID);
  const members = getMembersForRegion_(CONFIG.REGION_CAPWATCH_DATA_FOLDER_ID, orgs);

  // Workspace directory maps (CAPID → userId/email)
  const capidMaps = buildWorkspaceCapidMaps(members);

  // Build a list of group definitions from Groups tab
  const defs = loadDutyPositionIdsRegionDefinitions_();
  if (!defs.length) {
    Logger.info('No dutyPositionIdsRegion definitions found in Groups tab. Nothing to do.');
    return;
  }

  // 1) Groups
  const groupStats = syncRegionDutyGroups_(defs, members, orgs);

  // 2) Chat spaces
  const spaceCache = (typeof DRY_RUN !== 'undefined' && DRY_RUN) ? {} : buildChatSpaceCache();
  if ((typeof DRY_RUN === 'undefined' || !DRY_RUN) && spaceCache === null) {
    Logger.info('Chat space cache failed; skipping Chat Spaces portion.');
  } else {
    const spaceStats = syncRegionDutyChatSpaces_(defs, orgs, capidMaps, spaceCache);
    Logger.info('updateRegionGroupChats ChatSpaces summary', spaceStats || {});
  }

  Logger.info('Finished updateRegionGroupChats', {
    durationMs: new Date() - start,
    groups: groupStats
  });
}

/**
 * Parse CAPWATCH file from a specific folder (REGION_CAPWATCH_DATA_FOLDER_ID).
 * This mirrors your existing parseFile() behavior but points at a different folder.
 */
function parseFileFromFolder_(folderId, baseName) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByName(baseName + '.txt');
  if (!files.hasNext()) throw new Error('Missing CAPWATCH file: ' + baseName + '.txt in folder ' + folderId);

  const file = files.next();
  const content = file.getBlob().getDataAsString();
  // CAPWATCH txt files are typically comma delimited. If yours are tab delimited, adjust.
  return Utilities.parseCsv(content);
}

/**
 * Load Organization rows from the region CAPWATCH folder and build org map.
 * Only needs orgid, scope, wing, region, unit, name.
 */
function getSquadronsForRegion_(folderId) {
  const rows = parseFileFromFolder_(folderId, 'Organization') || [];
  const orgs = {};

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const orgid = String(r[0] || '').trim();
    if (!orgid) continue;

    orgs[orgid] = {
      orgid: orgid,
      region: String(r[1] || '').trim().toUpperCase(),
      wing: String(r[2] || '').trim().toUpperCase(),
      unit: String(r[3] || '').trim(),
      nextLevel: String(r[4] || '').trim(),
      name: String(r[5] || '').trim(),
      scope: String(r[9] || '').trim().toUpperCase()
    };
  }

  return orgs;
}

/**
 * Load Member rows from the region CAPWATCH folder.
 * Keeps ACTIVE members only.
 * Stores orgid and capid.
 */
function getMembersForRegion_(folderId, orgs) {
  const rows = parseFileFromFolder_(folderId, 'Member') || [];
  const members = {};

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const capid = String(r[0] || '').trim();
    const orgid = String(r[11] || '').trim();
    const status = String(r[24] || '').trim().toUpperCase();

    if (!capid) continue;
    if (status !== 'ACTIVE') continue;
    if (!orgid || !orgs[orgid]) continue;

    members[capid] = { capsn: capid, orgid: orgid, email: null };
  }

  return members;
}

/**
 * Pull Groups tab rows where Attribute == dutyPositionIdsRegion
 * Expected columns:
 *  [0]=Category, [1]=Group Name, [2]=Attribute, [3]=Values, [4]=Description
 */
function loadDutyPositionIdsRegionDefinitions_() {
  const ss = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Groups');
  if (!sheet) throw new Error('Missing sheet: Groups');

  const rows = sheet.getDataRange().getValues();
  const defs = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const groupName = String(row[1] || '').trim();
    const attr = String(row[2] || '').trim();
    const valuesRaw = row[3];
    const desc = String(row[4] || '').trim();

    if (!groupName) continue;
    if (attr !== 'dutyPositionIdsRegion') continue;

    const values = parseCsvListRegion_(valuesRaw);
    if (!values.length) continue;

    defs.push({
      groupName: groupName,         // base group name from sheet
      displayName: desc || groupName, // used for Chat space name default
      values: values                // duty titles to match
    });
  }

  return defs;
}

function parseCsvListRegion_(raw) {
  const s = String(raw || '').trim();
  if (!s) return [];
  try {
    const parsed = Utilities.parseCsv(s);
    const row = (parsed && parsed.length) ? parsed[0] : [];
    return (row || []).map(v => String(v || '').trim()).filter(v => v);
  } catch (e) {
    return s.split(',').map(v => String(v || '').trim()).filter(v => v);
  }
}

/**
 * Build CAPID -> duty titles list ONLY for DutyPosition rows where ORG scope is REGION or WING.
 */
function buildRegionDutyIndex_(folderId, orgs) {
  const dutyRows = parseFileFromFolder_(folderId, 'DutyPosition') || [];
  const capidToDutyTitles = {}; // capid -> Set(dutyLower)

  for (let i = 1; i < dutyRows.length; i++) {
    const r = dutyRows[i] || [];
    const capid = String(r[0] || '').trim();
    const duty = String(r[1] || '').trim();
    const orgid = String(r[7] || '').trim();

    if (!capid || !duty || !orgid) continue;

    const org = orgs[orgid];
    if (!org) continue;

    const scope = String(org.scope || '').toUpperCase();
    if (scope !== 'REGION' && scope !== 'WING') continue;

    if (!capidToDutyTitles[capid]) capidToDutyTitles[capid] = new Set();
    capidToDutyTitles[capid].add(duty.toLowerCase());
  }

  return capidToDutyTitles;
}

/**
 * Create/update Google Groups for each definition using REGION+WING duty positions.
 * Group email uses your existing convention: pcr.<groupName>@domain (CONFIG.WING assumed PCR).
 */
function syncRegionDutyGroups_(defs, members, orgs) {
  const start = new Date();

  const folderId = CONFIG.REGION_CAPWATCH_DATA_FOLDER_ID;
  const dutyIndex = buildRegionDutyIndex_(folderId, orgs);

  let created = 0;
  let updated = 0;
  let errors = 0;

  defs.forEach(function(d) {
    const groupId = String(CONFIG.WING || 'pcr').toLowerCase() + '.' + d.groupName;
    const groupEmail = groupId + CONFIG.EMAIL_DOMAIN;

    // Ensure group exists (create if missing)
    try {
      if (typeof DRY_RUN !== 'undefined' && DRY_RUN) {
        Logger.info('[DRY RUN] Would ensure group exists', { groupEmail: groupEmail });
      } else {
        try {
          AdminDirectory.Groups.get(groupEmail);
        } catch (e) {
          if (e.details && e.details.code === 404) {
            AdminDirectory.Groups.insert({
              email: groupEmail,
              name: d.displayName,
              description: d.displayName
            });
            created++;
          } else {
            throw e;
          }
        }
      }
    } catch (e) {
      Logger.error('Failed ensuring group exists', { groupEmail: groupEmail, error: e.message });
      errors++;
      return;
    }

    // Desired members by duty match
    const wanted = new Set(d.values.map(v => String(v).toLowerCase()));
    const desiredEmails = new Set();

    Object.keys(members).forEach(function(capid) {
      const m = members[capid];
      if (!m || !m.email) return;

      const duties = dutyIndex[capid];
      if (!duties || duties.size === 0) return;

      let match = false;
      wanted.forEach(function(w) {
        if (!match && duties.has(w)) match = true;
      });

      if (match) desiredEmails.add(m.email.toLowerCase());
    });

    // Current members
    let current = [];
    try {
      current = getCurrentGroupMembersOnly_(groupEmail);
    } catch (e) {
      Logger.error('Failed reading current group members', { groupEmail: groupEmail, error: e.message });
      errors++;
      return;
    }

    const currentSet = new Set(current);
    const toAdd = [];
    const toRemove = [];

    desiredEmails.forEach(function(em) {
      if (!currentSet.has(em)) toAdd.push(em);
    });

    currentSet.forEach(function(em) {
      if (!desiredEmails.has(em)) toRemove.push(em);
    });

    // Apply deltas
    if (typeof DRY_RUN !== 'undefined' && DRY_RUN) {
      Logger.info('[DRY RUN] Would update group membership', {
        groupEmail: groupEmail,
        add: toAdd.length,
        remove: toRemove.length
      });
      return;
    }

    toAdd.forEach(function(em) {
      try {
        executeWithRetry(function() {
          AdminDirectory.Members.insert({ email: em, role: 'MEMBER' }, groupEmail);
        });
      } catch (e) {
        errors++;
        Logger.error('Group add failed', { groupEmail: groupEmail, email: em, error: e.message });
      }
    });

    toRemove.forEach(function(em) {
      try {
        executeWithRetry(function() {
          AdminDirectory.Members.remove(groupEmail, em);
        });
      } catch (e) {
        errors++;
        Logger.error('Group remove failed', { groupEmail: groupEmail, email: em, error: e.message });
      }
    });

    updated++;
    Logger.info('Region duty group synced', {
      groupEmail: groupEmail,
      desired: desiredEmails.size,
      add: toAdd.length,
      remove: toRemove.length
    });
  });

  return {
    durationMs: new Date() - start,
    created: created,
    updated: updated,
    errors: errors
  };
}

function getCurrentGroupMembersOnly_(groupEmail) {
  let members = [];
  let nextPageToken = '';
  do {
    const page = AdminDirectory.Members.list(groupEmail, {
      roles: 'MEMBER',
      maxResults: 200,
      pageToken: nextPageToken || undefined
    });
    (page.members || []).forEach(function(m) {
      if (m && m.email) members.push(String(m.email).toLowerCase());
    });
    nextPageToken = page.nextPageToken;
  } while (nextPageToken);
  return members;
}

/**
 * Create/update Chat Spaces for each definition.
 * Membership is derived from REGION+WING duty positions (same index used for groups).
 */
function syncRegionDutyChatSpaces_(defs, orgs, capidMaps, existingSpaces) {
  const start = new Date();

  const folderId = CONFIG.REGION_CAPWATCH_DATA_FOLDER_ID;
  const dutyIndex = buildRegionDutyIndex_(folderId, orgs);

  const capidToUserId = capidMaps.capidToUserId || {};
  let spacesTouched = 0;
  let membersAdded = 0;

  defs.forEach(function(d) {
    const displayName = String(d.displayName || '').trim();
    if (!displayName) return;

    let spaceName = existingSpaces[displayName];
    if (!spaceName) {
      spaceName = createChatSpace(displayName, { historyOff: true, restrictPermissions: false });
      if (!spaceName) return;
      existingSpaces[displayName] = spaceName;
    }

    // Description/guidelines (optional)
    try {
      Chat.Spaces.patch(
        {
          spaceDetails: {
            description: 'This Chat space is for ' + displayName + ', and is for official Civil Air Patrol communications only.',
            guidelines: '📋 Chat Space Guidelines\n------------------------\nThis space is for official Civil Air Patrol communications only.'
          }
        },
        spaceName,
        { updateMask: 'spaceDetails' }
      );
    } catch (e) {}

    // Target members by duty match
    const wanted = new Set(d.values.map(v => String(v).toLowerCase()));
    const targetUserIds = new Set();

    Object.keys(dutyIndex).forEach(function(capid) {
      const duties = dutyIndex[capid];
      if (!duties || duties.size === 0) return;

      let match = false;
      wanted.forEach(function(w) {
        if (!match && duties.has(w)) match = true;
      });
      if (!match) return;

      const uid = capidToUserId[capid];
      if (uid) targetUserIds.add(uid);
    });

    const existingUserIds = listChatSpaceMemberUserIds(spaceName);
    const toAdd = [];
    targetUserIds.forEach(function(uid) {
      if (!existingUserIds.has(uid)) toAdd.push(uid);
    });

    toAdd.forEach(function(uid) {
      if (addMemberToChatSpace(spaceName, uid)) membersAdded++;
    });

    spacesTouched++;
    Logger.info('Region duty ChatSpace synced', {
      displayName: displayName,
      target: targetUserIds.size,
      added: toAdd.length
    });
  });

  return {
    durationMs: new Date() - start,
    spacesTouched: spacesTouched,
    membersAdded: membersAdded
  };
}