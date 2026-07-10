/**
 * -------------------------------------------------------------------------
 * Version: 2.0.0
 * Date: 2026-07-09
 * Author: Lt Col Noel Luneau, Pacific Region
 * Changes: Converged as the single shared ChatSpaces module (adopted the
 *   Pacific superset). vs the prior wing version this adds automation-group and
 *   user-additions chat spaces — gated behind PROFILE_.RUN_AUTOMATION_CHAT_SPACES
 *   (off for the wing) — plus an empty-vs-null cache-safety fix. Two corrections
 *   vs the Pacific copy: buildWorkspaceCapidMaps keeps customer:"my_customer"
 *   (not domain:), and getMembersForChatSpaces_ falls back to INDEFINITE not
 *   LIFE. See PCR_CHANGELOG.md.
 *
 * Description: ChatSpaces Module
 *
 * Creates and manages Google Chat Spaces for:
 *  - CAP Committee spaces (from MbrPCommittee)
 *  - Unit spaces (Region, Wing, Group, Squadron, Flight, Senior)
 *
 * Updates:
 * Chatspaces are created with history on
 * Chatspace add manual members defined in User Additions tab of
 * CAPWATCH Automation
 * -------------------------------------------------------------------------
 */

/**
 * Main entrypoint.
 */
const UPDATE_CHATSPACES_DRY_RUN = false; // Set true to dry-run updateChatSpaces only

function isUpdateChatSpacesDryRun_() {
  return !!UPDATE_CHATSPACES_DRY_RUN;
}

function updateChatSpaces() {
  const start = new Date();
  Logger.info(JSON.stringify({
    level: 'INFO',
    message: 'Starting ChatSpaces sync',
    data: { dryRun: isUpdateChatSpacesDryRun_() }
  }));

const squadrons = getSquadronsForChatSpaces_();
const members   = getMembersForChatSpaces_(squadrons);

  const capidMaps      = buildWorkspaceCapidMaps(members);
  const existingSpaces = isUpdateChatSpacesDryRun_() ? {} : buildChatSpaceCache();

  // Safety stop: abort only if cache FAILED (null). Empty cache is valid.
  if (!isUpdateChatSpacesDryRun_() && existingSpaces === null) {
    Logger.info(JSON.stringify({
      level: 'ERROR',
      message: 'Aborting ChatSpaces sync because Chat API/app is not configured or not authorized (space cache failed).',
      data: {}
    }));
    return;
  }

  const committeeStats   = syncCommitteeChatSpaces(members, squadrons, capidMaps, existingSpaces) || {};
  const unitStats        = syncUnitChatSpaces(members, squadrons, capidMaps, existingSpaces) || {};
  // Automation-group + user-additions chat spaces are region features: their loaders
  // fall back to the Groups / User Additions tabs the wing also has, so gate them by
  // profile (PROFILE_.RUN_AUTOMATION_CHAT_SPACES — off for the wing).
  const automationStats  = PROFILE_.RUN_AUTOMATION_CHAT_SPACES
    ? (syncAutomationGroupChatSpaces(members, capidMaps, existingSpaces) || {})
    : {};
  const additionsStats   = PROFILE_.RUN_AUTOMATION_CHAT_SPACES
    ? (syncUserAdditionsChatSpaces_(capidMaps, existingSpaces) || {})
    : {};

  const totalMembersAdded =
    (committeeStats.membersAdded || 0) +
    (unitStats.membersAdded || 0) +
    (automationStats.membersAdded || 0) +
    (additionsStats.membersAdded || 0);

  const totalManagersPromoted =
    (committeeStats.managersPromoted || 0) +
    (unitStats.managersPromoted || 0) +
    (automationStats.managersPromoted || 0) +
    (additionsStats.managersPromoted || 0);

  Logger.info(JSON.stringify({
    level: 'INFO',
    message: 'ChatSpaces overall summary',
    data: {
      totalMembersAdded: totalMembersAdded,
      totalManagersPromoted: totalManagersPromoted,
      durationMs: new Date() - start
    }
  }));

  Logger.info(JSON.stringify({
    level: 'INFO',
    message: 'Finished ChatSpaces sync',
    data: { durationMs: new Date() - start }
  }));
}

function getSquadronsForChatSpaces_() {
  const squadrons = {};
  const orgRows = parseFile('Organization') || [];

  const cfgWing = String(CONFIG.WING || '').trim().toUpperCase();
  const cfgRegion = String(CONFIG.REGION || '').trim().toUpperCase();

  for (let i = 0; i < orgRows.length; i++) {
    const row = orgRows[i] || [];
    const orgid  = String(row[0] || '').trim();
    const region = String(row[1] || '').trim().toUpperCase();
    const wing   = String(row[2] || '').trim().toUpperCase();
    const unit   = String(row[3] || '').trim();
    const next   = String(row[4] || '').trim();
    let name     = String(row[5] || '').trim();
    const scope  = String(row[9] || '').trim().toUpperCase();

    if (!orgid) continue;

    name = name.replace(/Pacific Region Cap/g, 'Pacific Region');

    let include = false;
    if (cfgWing && cfgWing.length === 2) {
      include = (wing === cfgWing);
    } else if (cfgWing === 'PCR') {
      // PCR HQ tenant: ONLY include REGION-scope PCR org rows.
      // This prevents pulling HI/CA/OR/NV wing/group/squadron orgs which also live under the PCR region.
      include = (scope === 'REGION' && region === (cfgRegion || 'PCR'));
    } else {
      include = cfgWing ? (wing === cfgWing) : false;
    }

    if (!include) continue;

    let charter = '';
    try {
      charter = Utilities.formatString('%s-%s-%03d', row[1], row[2], row[3]);
    } catch (ignored) {}

    squadrons[orgid] = {
      orgid: orgid,
      name: name,
      charter: charter,
      unit: unit,
      nextLevel: next,
      scope: scope,
      wing: wing,
      region: region,
      orgPath: ''
    };
  }

  try {
    const paths = parseFile('OrgPaths') || [];
    for (let i = 0; i < paths.length; i++) {
      const oid = String(paths[i][0] || '').trim();
      const path = String(paths[i][1] || '').trim();
      if (oid && squadrons[oid]) squadrons[oid].orgPath = path;
    }
  } catch (e) {}

  return squadrons;
}

function getMembersForChatSpaces_(squadrons) {
  const start = new Date();
  const members = {};

  const types = (CONFIG.MEMBER_TYPES && CONFIG.MEMBER_TYPES.ACTIVE)
    ? CONFIG.MEMBER_TYPES.ACTIVE
    : ['CADET', 'SENIOR', 'FIFTY YEAR', 'INDEFINITE', 'AEM'];

  const rows = parseFile('Member') || [];
  let count = 0;

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || [];
    const status = String(r[24] || '').trim().toUpperCase();
    const orgid  = String(r[11] || '').trim();
    const unitNo = String(r[13] || '').trim();
    const type   = String(r[21] || '').trim().toUpperCase();

    if (status !== 'ACTIVE') continue;
    if (unitNo === '0' || unitNo === '999') continue;
    if (types.indexOf(type) === -1) continue;
    if (!squadrons || !squadrons[orgid]) continue;

    const capid = String(r[0] || '').trim();
    if (!capid) continue;

    members[capid] = { capsn: capid, orgid: orgid, email: null };
    count++;
  }

  Logger.info(JSON.stringify({
    level: 'INFO',
    message: 'ChatSpaces standalone members loaded',
    data: { count: count, durationMs: new Date() - start }
  }));

  return members;
}

function isOrgInTenantScope_(org) {
  if (!org) return false;

  const cfgWing = String(CONFIG.WING || '').trim().toUpperCase();
  const cfgRegion = String(CONFIG.REGION || '').trim().toUpperCase();

  const scope = String(org.scope || '').trim().toUpperCase();
  const wing  = String(org.wing || '').trim().toUpperCase();
  const region = String(org.region || '').trim().toUpperCase();

  if (cfgWing && cfgWing.length === 2) return wing === cfgWing;

  if (cfgWing === 'PCR') {
    // PCR HQ tenant: only REGION spaces.
    return (scope === 'REGION' && region === (cfgRegion || 'PCR'));
  }

  return cfgWing ? (wing === cfgWing) : false;
}

/**
 * Build CAPID → Workspace userId/email maps from AdminDirectory.Users.
 */
function buildWorkspaceCapidMaps(members) {
  const capidToUserId = {};
  const capidToEmail  = {};
  const emailToUserId = {};
  let pageToken       = '';

  try {
    do {
      const resp = AdminDirectory.Users.list({
        customer: "my_customer",
        maxResults: 500,
        projection: 'full',
        fields: 'users(id,primaryEmail,externalIds),nextPageToken',
        pageToken: pageToken || undefined
      });

      (resp.users || []).forEach(function(u) {
        const primaryEmail = (u.primaryEmail || '').toLowerCase();
        if (primaryEmail && u.id) {
          emailToUserId[primaryEmail] = u.id;
        }
        const externalIds  = u.externalIds || [];
        const eid          = externalIds.find(function(e) {
          return e && e.type === 'organization' && e.value;
        });

        if (eid && eid.value) {
          const capid = String(eid.value);
          capidToUserId[capid] = u.id;
          capidToEmail[capid]  = primaryEmail;
          if (members[capid]) {
            members[capid].email = primaryEmail;
          }
        }
      });

      pageToken = resp.nextPageToken;
    } while (pageToken);

    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'Built Workspace CAPID maps',
      data: { count: Object.keys(capidToUserId).length }
    }));
  } catch (e) {
    Logger.info(JSON.stringify({
      level: 'ERROR',
      message: 'Failed to build Workspace CAPID maps',
      data: { error: e.message }
    }));
  }

  return { capidToUserId: capidToUserId, capidToEmail: capidToEmail, emailToUserId: emailToUserId };
}

/**
 * Cache existing Chat Spaces: displayName → space.name
 */
function buildChatSpaceCache() {
  const cache = {};
  let pageToken = '';

  try {
    do {
      const resp = Chat.Spaces.list({
        pageSize: 100,
        pageToken: pageToken || undefined
      });

      (resp.spaces || []).forEach(function(space) {
        if (space.spaceType === 'SPACE' && space.displayName) {
          cache[space.displayName] = space.name;
        }
      });

      pageToken = resp.nextPageToken;
    } while (pageToken);

    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'Cached Chat Spaces',
      data: { count: Object.keys(cache).length }
    }));

    // Empty cache is valid (e.g., brand new Chat app not yet in any spaces)
    return cache;
  } catch (e) {
    Logger.info(JSON.stringify({
      level: 'ERROR',
      message: 'Failed to cache Chat Spaces',
      data: { error: e.message }
    }));

    // Return null on failure so callers can tell “failed” vs “empty”
    return null;
  }
}

/**
 * Parse a CSV list from a single spreadsheet cell.
 * Supports quoted values containing commas, e.g. "DCS, Communications".
 * Returns an array of trimmed, non empty strings.
 */
function parseCsvList_(raw) {
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
 * For PCR tenant only: map CAPWATCH committee name (Automation sheet "Values") -> desired ChatSpace name (Automation sheet "Description").
 * This aligns committee ChatSpace naming with UpdateGroups (uses Description column).
 */
function buildCommitteeDisplayNameMapFromAutomation_() {
  const map = {};

  // Only applies to PCR tenant
  if (String(CONFIG.WING || '').toUpperCase() !== 'PCR') return map;

  const ssId = (CONFIG && CONFIG.AUTOMATION_SPREADSHEET_ID)
    ? String(CONFIG.AUTOMATION_SPREADSHEET_ID).trim()
    : '';
  if (!ssId) return map;

  let ss;
  try {
    ss = SpreadsheetApp.openById(ssId);
  } catch (e) {
    return map;
  }
  if (!ss) return map;

  const sheet = ss.getSheetByName('Automation') || ss.getSheetByName('Groups') || ss.getSheets()[0];
  if (!sheet) return map;

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return map;

  // UpdateGroups-style columns:
  // [0]=Category, [1]=Group Name, [2]=Attribute, [3]=Values, [4]=Description
  for (let r = 1; r < values.length; r++) {
    const row = values[r] || [];
    const category = String(row[0] || '').trim().toLowerCase();
    if (category !== 'committee') continue;

    const committeeName = String(row[3] || '').trim();
    const description   = String(row[4] || '').trim();
    if (!committeeName || !description) continue;

    // Map normalized committee key to the desired display name
    map[committeeName.toLowerCase()] = description;
  }

  return map;
}

/**
 * Committee Chat Spaces
 * Mirrors Python sync_chatspaces logic using MbrCommittee + Organization.
 */
function syncCommitteeChatSpaces(members, squadrons, capidMaps, existingSpaces) {
  const start = new Date();
  const rows  = parseFile('MbrCommittee') || [];
  const committeeNameMap = buildCommitteeDisplayNameMapFromAutomation_();

  if (!rows.length) {
    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'No MbrCommittee rows; skipping committee ChatSpaces'
    }));
    return;
  }

  // Index committees by "committee|orgid"
  const index = {};
    rows.forEach(function(row) {
      // CAPWATCH MbrCommittee columns:
      // [0]=CAPID, [1]=Committee, [2]=Chair, [3]=ORGID, [4]=DateAssigned
      const capid        = String(row[0] || '').trim();
      const committeeRaw = (row[1] || '');
      const chairFlagRaw = (row[2] || '');
      const orgid        = String(row[3] || '').trim();

    if (!committeeRaw || !orgid) return;

    const committee = String(committeeRaw).trim();
    if (!committee) return;

    const org = squadrons[orgid];
    if (!org) return;

    if (!isOrgInTenantScope_(org)) return;

    const key = committee.toLowerCase() + '|' + orgid;
    if (!index[key]) {
      index[key] = {
        committee: committee,
        orgid: orgid,
        org: org,
        members: [],
        chairs: []
      };
    }

    index[key].members.push(capid);

    // Chair logic (full)
    const chairValue = String(chairFlagRaw).trim();
    if (['Y','Yes','TRUE','1','y','yes','true',1].indexOf(chairValue) !== -1) {
      index[key].chairs.push(capid);
    }
  });

  const keys = Object.keys(index);
  Logger.info(JSON.stringify({
    level: 'INFO',
    message: 'Committee instances indexed for ChatSpaces',
    data: { count: keys.length }
  }));

  if (isUpdateChatSpacesDryRun_()) {
    keys.forEach(function(k) {
      const entry = index[k];
      const name  = buildCommitteeSpaceName(entry.org, entry.committee);
      Logger.info(JSON.stringify({
        level: 'INFO',
        message: '[DRY RUN] Would create/sync committee ChatSpace',
        data: { displayName: name, orgid: entry.orgid }
      }));
    });
    return;
  }

  const capidToUserId = capidMaps.capidToUserId;
  let totalAdded = 0;

  keys.forEach(function(k) {
    const entry = index[k];
    const org   = entry.org;
    // For PCR tenant, align committee space names to Automation sheet Description values (UpdateGroups naming)
    // Example: committee="Finance Committee" -> displayName="Finance Committee" (not "PCR - Finance Committee")
    let displayName = buildCommitteeSpaceName(org, entry.committee);
    if (String(CONFIG.WING || '').toUpperCase() === 'PCR') {
      const mapped = committeeNameMap[String(entry.committee || '').trim().toLowerCase()];
      if (mapped) {
        displayName = mapped;
      } else {
        // Fallback: use the raw committee name without PCR prefix
        displayName = String(entry.committee || '').trim();
      }
    }

    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'Processing committee ChatSpace',
      data: { displayName: displayName, orgid: entry.orgid }
    }));

    // Ensure space exists
    let spaceName = existingSpaces[displayName];

    // PCR tenant: if an older prefixed space exists, reuse it and rename it to the desired Automation name
    if (!spaceName && String(CONFIG.WING || '').toUpperCase() === 'PCR') {
      const legacyName = 'PCR - ' + String(entry.committee || '').trim();
      if (existingSpaces[legacyName]) {
        spaceName = existingSpaces[legacyName];
        try {
          Chat.Spaces.patch(
            { displayName: displayName },
            spaceName,
            { updateMask: 'displayName' }
          );
          // Update local cache keys
          delete existingSpaces[legacyName];
          existingSpaces[displayName] = spaceName;
          Logger.info(JSON.stringify({
            level: 'INFO',
            message: 'Renamed legacy committee ChatSpace to Automation display name',
            data: { from: legacyName, to: displayName, spaceName: spaceName }
          }));
        } catch (e) {
          // If rename not allowed, keep using the legacy name to avoid duplicates
          existingSpaces[legacyName] = spaceName;
          displayName = legacyName;
          Logger.info(JSON.stringify({
            level: 'WARN',
            message: 'Could not rename legacy committee ChatSpace; using legacy name to avoid duplicate creation',
            data: { legacyName: legacyName, desiredName: displayName, error: e.message }
          }));
        }
      }
    }

    if (!spaceName) {
      spaceName = createChatSpace(displayName, { historyOff: true, restrictPermissions: false });
      if (!spaceName) {
        Logger.info(JSON.stringify({
          level: 'ERROR',
          message: 'Failed to create committee ChatSpace',
          data: { displayName: displayName }
        }));
        return;
      }
      existingSpaces[displayName] = spaceName;
    }

    // Description & guidelines (same as unit spaces)
    const desc =
      'This Chat space is for ' + displayName +
      ', and is for official Civil Air Patrol communications only.';
    const guidelines =
      '📋 Chat Space Guidelines\n' +
      '------------------------\n' +
      'This space is for official Civil Air Patrol communications only.';

    try {
      Chat.Spaces.patch(
        { spaceDetails: { description: desc, guidelines: guidelines } },
        spaceName,
        { updateMask: 'spaceDetails' }
      );
    } catch (e) {
      Logger.info(JSON.stringify({
        level: 'WARN',
        message: 'Could not set description/guidelines for committee ChatSpace',
        data: { displayName: displayName, error: e.message }
      }));
    }

    // Target members (userIds)
    const targetUserIds = new Set();
    entry.members.forEach(function(capid) {
      const uid = capidToUserId[capid];
      if (uid) targetUserIds.add(uid);
    });

    const existingUserIds = listChatSpaceMemberUserIds(spaceName);
    const toAdd = [];
    targetUserIds.forEach(function(uid) {
      if (!existingUserIds.has(uid)) toAdd.push(uid);
    });

    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'Committee membership delta',
      data: { displayName: displayName, toAdd: toAdd.length }
    }));

    let addedHere = 0;
    toAdd.forEach(function(uid) {
      if (addMemberToChatSpace(spaceName, uid)) {
        addedHere++;
        totalAdded++;
      }
    });

    // Promote chair to manager (if any)
    if (entry.chairs.length) {
      const chairCapid = entry.chairs[0];
      const chairUserId = capidToUserId[chairCapid];
      if (chairUserId) {
        promoteChatSpaceManager(spaceName, chairUserId, {
          displayName: displayName,
          capid: chairCapid,
          role: 'Chair'
        });
      } else {
        Logger.info(JSON.stringify({
          level: 'WARN',
          message: 'Chair CAPID not found in directory',
          data: { displayName: displayName, capid: chairCapid }
        }));
      }
    }

    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'Committee ChatSpace sync complete',
      data: { displayName: displayName, added: addedHere }
    }));
  });

  Logger.info(JSON.stringify({
    level: 'INFO',
    message: 'Committee ChatSpaces summary',
    data: { totalSpaces: keys.length, membersAdded: totalAdded, durationMs: new Date() - start }
  }));
  return {
    membersAdded: totalAdded,
    managersPromoted: 0
  };
}

/**
 * Build committee ChatSpace name with same prefixes as Python:
 *  REGION   → "PCR - "
 *  WING     → "{WING}WG - "
 *  GROUP    → "{WING}WG-GP# - "
 *  SQUADRON → "{WING}{UNIT} - "
 */
function buildCommitteeSpaceName(org, committeeName) {
  const orgScope = (org.scope || '').toUpperCase();
  const orgWing  = (org.wing || CONFIG.WING || '').toUpperCase();
  const orgUnit  = (org.unit || '').toString().padStart(3, '0');
  const orgName  = org.name || '';

  let prefix = '';
  if (orgScope === 'REGION') {
    // Same as Python: hardcoded PCR prefix
    prefix = 'PCR - ';
  } else if (orgScope === 'WING') {
    prefix = orgWing + 'WG - ';
  } else if (orgScope === 'GROUP') {
    const match = orgName.match(/GROUP\s*(\d+)/i);
    const groupStr = match ? 'GP' + match[1] : '';
    prefix = orgWing + 'WG' + (groupStr ? '-' + groupStr : '') + ' - ';
  } else if (orgScope === 'SQUADRON') {
    prefix = orgWing + orgUnit + ' - ';
  }

  return prefix + String(committeeName).trim();
}

/**
 * Unit Chat Spaces
 * One space per Region/Wing/Group/Squadron/Flight/Senior
 * Display name = full unit name in sentence case.
 */
function syncUnitChatSpaces(members, squadrons, capidMaps, existingSpaces) {
  const start = new Date();

  // PCR HQ tenant: do NOT create a REGION unit space (we use Automation sheet spaces such as "All Members" instead).
  // Wings (2-letter CONFIG.WING) continue to create REGION/WING/GROUP/SQUADRON/etc spaces as before.
  if (String(CONFIG.WING || '').trim().toUpperCase() === 'PCR') {
    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'Skipping unit ChatSpaces for PCR tenant (use Automation sheet spaces such as All Members instead).',
      data: { wing: 'PCR' }
    }));
    return;
  }

  const cfgWing = String(CONFIG.WING || '').trim().toUpperCase();
  const validScopes = (cfgWing === 'PCR')
    ? ['REGION']
    : ['REGION', 'WING', 'GROUP', 'SQUADRON', 'FLIGHT', 'SENIOR'];
  const orgs = Object.keys(squadrons)
    .map(function(k) { return squadrons[k]; })
    .filter(function(o) {
      if (!o || !o.scope) return false;
      const scope = String(o.scope).toUpperCase();
      if (validScopes.indexOf(scope) === -1) return false;
      return isOrgInTenantScope_(o);
    });

  Logger.info(JSON.stringify({
    level: 'INFO',
    message: 'Unit orgs for ChatSpaces',
    data: { count: orgs.length }
  }));

  if (isUpdateChatSpacesDryRun_()) {
    orgs.forEach(function(o) {
      Logger.info(JSON.stringify({
        level: 'INFO',
        message: '[DRY RUN] Would create/sync unit ChatSpace',
        data: { displayName: formatUnitDisplayName(o.name || ''), orgid: o.orgid }
      }));
    });
    return;
  }

  const capidToUserId = capidMaps.capidToUserId;
  let totalSpaces = 0;
  let totalAdded = 0;
  let totalManagers = 0;

  const dutyRows = parseFile('DutyPosition') || [];

  orgs.forEach(function(org) {
    const orgid  = org.orgid;
    const scope  = (org.scope || '').toUpperCase();
    const name   = org.name || '';
    const displayName = formatUnitDisplayName(name);

    if (!displayName) return;

    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'Processing unit ChatSpace',
      data: { displayName: displayName, orgid: orgid, scope: scope }
    }));

    let spaceName = existingSpaces[displayName];
    if (!spaceName) {
      spaceName = createChatSpace(displayName, { historyOff: true, restrictPermissions: true });
      if (!spaceName) {
        Logger.info(JSON.stringify({
          level: 'ERROR',
          message: 'Failed to create unit ChatSpace',
          data: { displayName: displayName }
        }));
        return;
      }
      existingSpaces[displayName] = spaceName;
      totalSpaces++;
    }

    // Add all members with this ORGID
    const existingUserIds = listChatSpaceMemberUserIds(spaceName);
    let addedHere = 0;

    Object.keys(members).forEach(function(capid) {
      const m = members[capid];
      if (!m) return;
      if (String(m.orgid) !== String(orgid)) return;

      const userId = capidToUserId[capid];
      if (!userId) return;
      if (existingUserIds.has(userId)) return;

      if (addMemberToChatSpace(spaceName, userId)) {
        addedHere++;
        totalAdded++;
      }
    });

    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'Unit ChatSpace members ensured',
      data: { displayName: displayName, added: addedHere }
    }));

    // Promote key duty positions to managers
    const includeRoles = [
      'commander',
      'deputy commander',
      'deputy commander for seniors',
      'deputy commander for cadets',
      'chief of staff',
      'public affairs',
      'information technology officer',
      'information technologies officer',
      'director of it'
    ];
    const excludeRoles = ['advisor to the commander'];

    let managersHere = 0;
    dutyRows.forEach(function(row) {
      // CAPWATCH DutyPosition: CAPID, Duty, FunctArea, Lvl, Asst, UsrID, DateMod, ORGID
      const capid = String(row[0]);
      const duty  = String(row[1] || '');
      const rowOrgid = row[7];

      if (String(rowOrgid) !== String(orgid)) return;

      const title = duty.toLowerCase();
      const include = includeRoles.some(function(r) { return title.indexOf(r) !== -1; });
      const exclude = excludeRoles.some(function(r) { return title.indexOf(r) !== -1; });
      if (!include || exclude) return;

      const userId = capidToUserId[capid];
      if (!userId) return;

      if (promoteChatSpaceManager(spaceName, userId, {
        displayName: displayName,
        capid: capid,
        duty: duty
      })) {
        managersHere++;
        totalManagers++;
      }
    });

    // Description & guidelines
    const desc =
      'This Chat space is for ' + displayName +
      ', and is for official Civil Air Patrol communications only.';
    const guidelines =
      '📋 Chat Space Guidelines\n' +
      '------------------------\n' +
      'This space is for official Civil Air Patrol communications only.';

    try {
      Chat.Spaces.patch(
        { spaceDetails: { description: desc, guidelines: guidelines } },
        spaceName,
        { updateMask: 'spaceDetails' }
      );
    } catch (e) {
      Logger.info(JSON.stringify({
        level: 'WARN',
        message: 'Could not set description/guidelines for unit ChatSpace',
        data: { displayName: displayName, error: e.message }
      }));
    }

    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'Unit ChatSpace processed',
      data: { displayName: displayName, managersPromoted: managersHere }
    }));
  });

  Logger.info(JSON.stringify({
    level: 'INFO',
    message: 'Unit ChatSpaces summary',
    data: {
      newSpaces: totalSpaces,
      membersAdded: totalAdded,
      managersPromoted: totalManagers,
      durationMs: new Date() - start
    }
  }));
  return {
    membersAdded: totalAdded,
    managersPromoted: totalManagers
  };
}

/**
 * Create a Chat Space with optional history + permission settings.
 */
function createChatSpace(displayName, options) {
  options = options || {};
  const historyOff        = !!options.historyOff;
  const restrictPerms     = !!options.restrictPermissions;
  let spaceName = null;

  try {
    const payload = {
      displayName: displayName,
      spaceType: 'SPACE'
    };

    // Allow external users to join (immutable, must be set at creation time)
    if (options && options.externalUserAllowed === true) {
      payload.externalUserAllowed = true;
    }

    const space = Chat.Spaces.create(payload);
    spaceName = space.name;

    if (historyOff === true) {
      try {
        Chat.Spaces.patch(
          { spaceHistoryState: 'HISTORY_OFF' },
          spaceName,
          { updateMask: 'spaceHistoryState' }
        );
      } catch (e) {
        Logger.info(JSON.stringify({
          level: 'WARN',
          message: 'Could not set history off for ChatSpace',
          data: { displayName: displayName, error: e.message }
        }));
      }
    } else {
      try {
        Chat.Spaces.patch(
          { spaceHistoryState: 'HISTORY_ON' },
          spaceName,
          { updateMask: 'spaceHistoryState' }
        );
      } catch (e) {
        Logger.info(JSON.stringify({
          level: 'WARN',
          message: 'Could not set history on for ChatSpace',
          data: { displayName: displayName, error: e.message }
        }));
      }
    }

    if (restrictPerms) {
      try {
        Chat.Spaces.patch(
          {
            permissionSettings: {
              manage_members_and_groups: 'OWNERS',
              modify_space_details: 'OWNERS',
              toggle_history: 'OWNERS',
              use_at_mention_all: 'OWNERS',
              manage_apps: 'OWNERS',
              manage_webhooks: 'OWNERS'
            }
          },
          spaceName,
          { updateMask: 'permissionSettings' }
        );
      } catch (e) {
        Logger.info(JSON.stringify({
          level: 'WARN',
          message: 'Could not update permissionSettings for ChatSpace',
          data: { displayName: displayName, error: e.message }
        }));
      }
    }

    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'ChatSpace created',
      data: { displayName: displayName, spaceName: spaceName }
    }));
  } catch (e) {
    Logger.info(JSON.stringify({
      level: 'ERROR',
      message: 'Failed to create ChatSpace',
      data: { displayName: displayName, error: e.message }
    }));
  }

  return spaceName;
}

/**
 * Return set of Workspace userIds present in a space.
 */
function listChatSpaceMemberUserIds(spaceName) {
  const ids = new Set();
  let pageToken = '';

  try {
    do {
      const resp = Chat.Spaces.Members.list(spaceName, {
        pageSize: 100,
        pageToken: pageToken || undefined
      });

      (resp.memberships || []).forEach(function(m) {
        const memberName = m.member && m.member.name || '';
        if (memberName.indexOf('users/') === 0) {
          const parts = memberName.split('/');
          const id = parts[parts.length - 1];
          if (id) ids.add(id);
        }
      });

      pageToken = resp.nextPageToken;
    } while (pageToken);
  } catch (e) {
    Logger.info(JSON.stringify({
      level: 'ERROR',
      message: 'Failed to list ChatSpace members',
      data: { spaceName: spaceName, error: e.message }
    }));
  }

  return ids;
}

/**
 * Add one user to a Chat Space.
 */
function addMemberToChatSpace(spaceName, userId) {
  try {
    Chat.Spaces.Members.create(
      { member: { name: 'users/' + userId, type: 'HUMAN' } },
      spaceName
    );
    return true;
  } catch (e) {
    // If already exists, API may throw; we just log and continue.
    Logger.info(JSON.stringify({
      level: 'WARN',
      message: 'Failed to add member to ChatSpace',
      data: { spaceName: spaceName, userId: userId, error: e.message }
    }));
    return false;
  }
}

/**
 * Promote a member to ROLE_MANAGER in a Chat Space.
 */
function promoteChatSpaceManager(spaceName, userId, context) {
  context = context || {};
  const membershipName = spaceName + '/members/' + userId;

  try {
    Chat.Spaces.Members.patch(
      { role: 'ROLE_MANAGER' },
      membershipName,
      { updateMask: 'role' }
    );
    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'Promoted ChatSpace member to manager',
      data: { spaceName: spaceName, userId: userId, context: context }
    }));
    return true;
  } catch (e) {
    Logger.info(JSON.stringify({
      level: 'WARN',
      message: 'Failed to promote ChatSpace member to manager',
      data: { spaceName: spaceName, userId: userId, context: context, error: e.message }
    }));
    return false;
  }
}

/**
 * Full unit name in sentence case.
 */
function formatUnitDisplayName(name) {
  if (!name) return '';
  const lower = String(name).toLowerCase();
  return lower.charAt(0).toUpperCase() + lower.slice(1);
}

/**
 * Automation-defined ChatSpaces (duty-based)
 * Creates one Chat space per definition in the Automation spreadsheet.
 * Membership is computed from CAPWATCH duty assignments (DutyAssignment/DutyPosition),
 * NOT from Google Groups.
 */
function syncAutomationGroupChatSpaces(members, capidMaps, existingSpaces) {
  const start = new Date();

  // Allow this function to be run standalone.
  // If existingSpaces wasn't passed in, build the cache here.
  if (!existingSpaces) {
    existingSpaces = isUpdateChatSpacesDryRun_() ? {} : buildChatSpaceCache();
  }

  // If cache FAILED (null), abort to avoid TypeErrors and mass creation attempts.
  if (!isUpdateChatSpacesDryRun_() && existingSpaces === null) {
    Logger.info(JSON.stringify({
      level: 'WARN',
      message: 'Skipping automation-defined ChatSpaces because space cache failed (Chat API/app likely not configured).',
      data: {}
    }));
    return;
  }

  // Avoid mass creation if Chat API not ready (existingSpaces is null when cache fails)
  if (!isUpdateChatSpacesDryRun_() && existingSpaces === null) {
    Logger.info(JSON.stringify({
      level: 'WARN',
      message: 'Skipping automation-defined ChatSpaces because space cache failed (Chat API/app likely not configured).',
      data: {}
    }));
    return;
  }

  const defs = loadAutomationChatSpaceDefinitions_();
  Logger.info(JSON.stringify({
    level: 'INFO',
    message: 'Automation ChatSpace definitions loaded',
    data: {
      count: defs.length,
      automationSpreadsheetId: String(CONFIG.AUTOMATION_SPREADSHEET_ID || '')
    }
  }));

  if (!defs.length) return;

  if (isUpdateChatSpacesDryRun_()) {
    defs.forEach(function(d) {
      Logger.info(JSON.stringify({
        level: 'INFO',
        message: '[DRY RUN] Would create/sync automation-defined ChatSpace',
        data: { displayName: d.displayName, matchType: d.matchType, valuesCount: d.values.length }
      }));
    });
    return;
  }

  const capidToUserId = (capidMaps && capidMaps.capidToUserId) ? capidMaps.capidToUserId : {};

  // PCR tenant: only DutyPosition.txt exists. Also, parseFile() may throw if a file is missing.
  let dutyRows = [];
  try {
    dutyRows = parseFile('DutyPosition') || [];
  } catch (e) {
    Logger.info(JSON.stringify({
      level: 'ERROR',
      message: 'Failed to load DutyPosition for automation-defined ChatSpaces',
      data: { error: e.message }
    }));
    dutyRows = [];
  }

  let spacesTouched = 0;
  let membersAdded = 0;

  defs.forEach(function(d) {
    const displayName = String(d.displayName || '').trim();
    if (!displayName) return;

    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'Processing automation-defined ChatSpace',
      data: { displayName: displayName, matchType: d.matchType, valuesCount: d.values.length }
    }));

    // Ensure space exists
    let spaceName = existingSpaces[displayName];
    if (!spaceName) {
      spaceName = createChatSpace(displayName, {
        historyOff: true,
        restrictPermissions: false,
        externalUserAllowed: d.externalUserAllowed === true
      });
      if (!spaceName) {
        Logger.info(JSON.stringify({
          level: 'ERROR',
          message: 'Failed to create automation-defined ChatSpace',
          data: { displayName: displayName }
        }));
        return;
      }
      existingSpaces[displayName] = spaceName;
    }

    // Description & guidelines (same as other spaces)
    const desc =
      'This Chat space is for ' + displayName +
      ', and is for official Civil Air Patrol communications only.';
    const guidelines =
      '📋 Chat Space Guidelines\n' +
      '------------------------\n' +
      'This space is for official Civil Air Patrol communications only.';

    try {
      Chat.Spaces.patch(
        { spaceDetails: { description: desc, guidelines: guidelines } },
        spaceName,
        { updateMask: 'spaceDetails' }
      );
    } catch (e) {
      Logger.info(JSON.stringify({
        level: 'WARN',
        message: 'Could not set description/guidelines for automation-defined ChatSpace',
        data: { displayName: displayName, error: e.message }
      }));
    }

    // Build target CAPIDs from duty assignments within tenant scope.
    const wanted = new Set((d.values || []).map(v => String(v || '').trim().toLowerCase()).filter(v => v));
    const targetUserIds = new Set();

    // DutyPosition layout (CAPWATCH):
    // [0]=CAPID, [1]=Duty, ... , [7]=ORGID
    // DutyAssignment may differ, but we still treat [0]=CAPID and [1]=Duty.
    for (let i = 0; i < dutyRows.length; i++) {
      const row = dutyRows[i] || [];
      const capid = String(row[0] || '').trim();
      const duty  = String(row[1] || '').trim();

      if (!capid || !duty) continue;

      // PCR tenant: include IAOD by scoping on DutyPosition ORGID (PCR HQ orgid 434),
      // not on primary membership orgid (which excludes IAOD).
      if (String(CONFIG.WING || '').trim().toUpperCase() === 'PCR') {
        const dutyOrgid = String(row[7] || '').trim(); // DutyPosition ORGID
        if (dutyOrgid !== '434') continue;
      } else {
        // Non-PCR tenants: limit to members already in this tenant scope (pre-filtered by getMembersForChatSpaces_).
        if (!members || !members[capid]) continue;
      }

      const dutyKey = duty.toLowerCase();

      let isMatch = false;
      if (d.matchType === 'duty-exact') {
        isMatch = wanted.has(dutyKey);
      } else if (d.matchType === 'duty-contains') {
        wanted.forEach(function(w) {
          if (!isMatch && dutyKey.indexOf(w) !== -1) isMatch = true;
        });
      } else {
        isMatch = wanted.has(dutyKey);
      }

      if (!isMatch) continue;

      const uid = capidToUserId[capid];
      if (uid) targetUserIds.add(uid);
    }

    const existingUserIds = listChatSpaceMemberUserIds(spaceName);
    const toAdd = [];
    targetUserIds.forEach(function(uid) {
      if (!existingUserIds.has(uid)) toAdd.push(uid);
    });

    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'Automation-defined ChatSpace membership delta',
      data: { displayName: displayName, target: targetUserIds.size, toAdd: toAdd.length }
    }));

    let addedHere = 0;
    toAdd.forEach(function(uid) {
      if (addMemberToChatSpace(spaceName, uid)) {
        addedHere++;
        membersAdded++;
      }
    });

    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'Automation-defined ChatSpace members ensured',
      data: { displayName: displayName, added: addedHere, targetCount: targetUserIds.size }
    }));

    spacesTouched++;
  });

  Logger.info(JSON.stringify({
    level: 'INFO',
    message: 'Automation-defined ChatSpaces summary',
    data: { spacesTouched: spacesTouched, membersAdded: membersAdded, durationMs: new Date() - start }
  }));
  return {
    membersAdded: membersAdded,
    managersPromoted: 0
  };
}

/**
 * User Additions → Chat Spaces (add-only)
 *
 * Reads the Automation spreadsheet "User Additions" tab and adds users to existing
 * Chat spaces that correspond to group definitions in the "Groups" tab.
 *
 * Rules:
 *  - Add only, never remove.
 *  - User Additions may NOT reference spaces/groups that are not present in the Groups tab.
 *  - Assumes the Chat spaces already exist. If a space is missing, log a WARN and skip.
 */
function syncUserAdditionsChatSpaces_(capidMaps, existingSpaces) {
  const start = new Date();

  if (isUpdateChatSpacesDryRun_()) {
    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'UPDATE_CHATSPACES_DRY_RUN enabled; skipping User Additions → ChatSpaces sync',
      data: {}
    }));
    return;
  }

  // existingSpaces === null means the Chat API/app is not configured/authorized.
  if (existingSpaces === null) {
    Logger.info(JSON.stringify({
      level: 'WARN',
      message: 'Skipping User Additions → ChatSpaces because Chat space cache failed (Chat API/app likely not configured).',
      data: {}
    }));
    return;
  }

  const ssId = (CONFIG && CONFIG.AUTOMATION_SPREADSHEET_ID)
    ? String(CONFIG.AUTOMATION_SPREADSHEET_ID).trim()
    : '';
  if (!ssId) {
    Logger.info(JSON.stringify({
      level: 'WARN',
      message: 'Automation spreadsheet ID not configured; skipping User Additions → ChatSpaces.',
      data: {}
    }));
    return;
  }

  // Build Groups tab lookup: baseGroupName -> desired ChatSpace displayName
  const groupToSpaceName = buildGroupNameToChatSpaceDisplayNameMap_();

  // Load User Additions rows
  const additions = loadUserAdditionsRows_();
  if (!additions.length) {
    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'No User Additions rows; skipping User Additions → ChatSpaces.',
      data: { durationMs: new Date() - start }
    }));
    return;
  }

  const emailToUserId = (capidMaps && capidMaps.emailToUserId) ? capidMaps.emailToUserId : {};

  // Group additions by ChatSpace displayName so we can list membership once per space.
  const planBySpace = {}; // displayName -> [{email,userId,desiredRole,sourceGroup}]

  additions.forEach(function(a) {
    const email = String(a.email || '').trim().toLowerCase();
    if (!email) return;

    const userId = emailToUserId[email];
    if (!userId) {
      Logger.info(JSON.stringify({
        level: 'WARN',
        message: 'User Additions email not found in Workspace directory; skipping for ChatSpaces',
        data: { email: email }
      }));
      return;
    }

    const roleRaw = String(a.role || '').trim().toUpperCase();
    const desiredRole = (roleRaw === 'MANAGER' || roleRaw === 'OWNER') ? 'ROLE_MANAGER' : 'ROLE_MEMBER';

    const groups = parseCsvList_(a.groupsRaw);
    if (!groups.length) return;

    groups.forEach(function(g) {
      const token = String(g || '').trim();
      if (!token) return;

      // Normalize: allow either "it" or "pcr.it"; base group name is everything after first dot.
      const fullGroupName = token.indexOf('.') === -1 ? ('pcr.' + token) : token;
      const baseGroupName = fullGroupName.split('.').slice(1).join('.');

      // Enforce: User Additions may only reference groups present in Groups tab.
      const displayName = groupToSpaceName[String(baseGroupName || '').toLowerCase()];
      if (!displayName) {
        Logger.info(JSON.stringify({
          level: 'WARN',
          message: 'User Additions references a group not found in Groups tab; skipping for ChatSpaces',
          data: { email: email, groupToken: token, baseGroupName: baseGroupName }
        }));
        return;
      }

      if (!planBySpace[displayName]) planBySpace[displayName] = [];
      planBySpace[displayName].push({
        email: email,
        userId: userId,
        desiredRole: desiredRole,
        sourceGroup: fullGroupName
      });
    });
  });

  const spaceNames = Object.keys(planBySpace);
  if (!spaceNames.length) {
    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'No eligible User Additions → ChatSpaces actions after validation; nothing to do.',
      data: { durationMs: new Date() - start }
    }));
    return;
  }

  let totalAdded = 0;
  let totalPromoted = 0;
  let totalMissingSpaces = 0;

  spaceNames.forEach(function(displayName) {
    const spaceName = existingSpaces[displayName];

    // Requirement: do NOT create spaces. Warn and skip if missing.
    if (!spaceName) {
      totalMissingSpaces++;
      Logger.info(JSON.stringify({
        level: 'WARN',
        message: 'ChatSpace does not exist for User Additions target (will not create).',
        data: { displayName: displayName }
      }));
      return;
    }

    const existingUserIds = listChatSpaceMemberUserIds(spaceName);
    const rows = planBySpace[displayName] || [];

    // De-dup by userId per space (keep the strongest role)
    const byUser = {}; // userId -> {desiredRole,email,sourceGroups:[]}
    rows.forEach(function(r) {
      if (!r || !r.userId) return;
      if (!byUser[r.userId]) {
        byUser[r.userId] = { desiredRole: r.desiredRole, email: r.email, sourceGroups: [r.sourceGroup] };
      } else {
        byUser[r.userId].sourceGroups.push(r.sourceGroup);
        // If any row requests manager, keep manager.
        if (r.desiredRole === 'ROLE_MANAGER') byUser[r.userId].desiredRole = 'ROLE_MANAGER';
      }
    });

    const userIds = Object.keys(byUser);

    let addedHere = 0;
    let promotedHere = 0;

    userIds.forEach(function(uid) {
      const info = byUser[uid];
      if (!info) return;

      if (!existingUserIds.has(uid)) {
        if (addMemberToChatSpace(spaceName, uid)) {
          addedHere++;
          totalAdded++;
          existingUserIds.add(uid);
        }
      }

      // Promotion is best-effort; if user isn't a member yet (add failed), promotion will fail and log WARN.
      if (info.desiredRole === 'ROLE_MANAGER') {
        const ok = promoteChatSpaceManager(spaceName, uid, {
          displayName: displayName,
          email: info.email,
          sourceGroups: info.sourceGroups
        });
        if (ok) {
          promotedHere++;
          totalPromoted++;
        }
      }
    });

    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'User Additions → ChatSpace processed',
      data: { displayName: displayName, added: addedHere, promoted: promotedHere }
    }));
  });

  Logger.info(JSON.stringify({
    level: 'INFO',
    message: 'User Additions → ChatSpaces summary',
    data: {
      spacesConsidered: spaceNames.length,
      spacesMissing: totalMissingSpaces,
      membersAdded: totalAdded,
      managersPromoted: totalPromoted,
      durationMs: new Date() - start
    }
  }));
  return {
    membersAdded: totalAdded,
    managersPromoted: totalPromoted
  };
}

/**
 * Build a lookup of base Groups-tab "Group Name" -> ChatSpace displayName.
 * For PCR tenant naming alignment, we prefer the Groups tab "Description" column when present.
 */
function buildGroupNameToChatSpaceDisplayNameMap_() {
  const map = {};

  const ssId = (CONFIG && CONFIG.AUTOMATION_SPREADSHEET_ID)
    ? String(CONFIG.AUTOMATION_SPREADSHEET_ID).trim()
    : '';
  if (!ssId) return map;

  let ss;
  try {
    ss = SpreadsheetApp.openById(ssId);
  } catch (e) {
    return map;
  }
  if (!ss) return map;

  const sheet = ss.getSheetByName('Groups') || ss.getSheetByName('Automation');
  if (!sheet) return map;

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return map;

  // Attempt to detect header names; otherwise fall back to UpdateGroups fixed columns.
  const header = (values[0] || []).map(h => String(h || '').trim().toLowerCase());

  function colIndex_(names) {
    for (let i = 0; i < names.length; i++) {
      const idx = header.indexOf(names[i]);
      if (idx >= 0) return idx;
    }
    return -1;
  }

  let idxGroupName = colIndex_(['group name', 'group', 'name']);
  let idxDesc      = colIndex_(['description', 'display name', 'display_name', 'title']);

  // Fixed columns fallback: [1]=Group Name, [4]=Description
  if (idxGroupName < 0) idxGroupName = 1;
  if (idxDesc < 0) idxDesc = 4;

  for (let r = 1; r < values.length; r++) {
    const row = values[r] || [];
    const groupName = String(row[idxGroupName] || '').trim();
    if (!groupName) continue;

    const desc = String(row[idxDesc] || '').trim();

    // For ChatSpaces we use the same rule as automation-defined spaces: displayName := Description when present.
    // If missing, fall back to the raw group name.
    const displayName = (desc || groupName).replace(/\s+/g, ' ').trim();

    map[groupName.toLowerCase()] = displayName;
  }

  return map;
}

/**
 * Load rows from Automation spreadsheet "User Additions".
 * Expected columns:
 *   A Name, B Email, C Role, D Groups
 */
function loadUserAdditionsRows_() {
  const out = [];

  const ssId = (CONFIG && CONFIG.AUTOMATION_SPREADSHEET_ID)
    ? String(CONFIG.AUTOMATION_SPREADSHEET_ID).trim()
    : '';
  if (!ssId) return out;

  let ss;
  try {
    ss = SpreadsheetApp.openById(ssId);
  } catch (e) {
    Logger.info(JSON.stringify({
      level: 'WARN',
      message: 'Could not open automation spreadsheet for User Additions',
      data: { error: e.message }
    }));
    return out;
  }
  if (!ss) return out;

  const sheet = ss.getSheetByName('User Additions');
  if (!sheet) return out;

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return out;

  // Header-aware (case-insensitive) with fixed fallback.
  const header = (values[0] || []).map(h => String(h || '').trim().toLowerCase());
  const idxName  = header.indexOf('name') >= 0 ? header.indexOf('name') : 0;
  const idxEmail = header.indexOf('email') >= 0 ? header.indexOf('email') : 1;
  const idxRole  = header.indexOf('role') >= 0 ? header.indexOf('role') : 2;
  const idxGroups = header.indexOf('groups') >= 0 ? header.indexOf('groups') : 3;

  for (let r = 1; r < values.length; r++) {
    const row = values[r] || [];
    const email = String(row[idxEmail] || '').trim();
    const groupsRaw = row[idxGroups];

    if (!email) continue;

    out.push({
      name: String(row[idxName] || '').trim(),
      email: email,
      role: String(row[idxRole] || '').trim(),
      groupsRaw: groupsRaw
    });
  }

  return out;
}

/**
 * Read ChatSpace definitions from the Automation spreadsheet.
 * Uses UpdateGroups-style columns:
 *   [0]=Category, [1]=Group Name, [2]=Attribute, [3]=Values, [4]=Description
 * For ChatSpaces we use:
 *   - displayName := Description (row[4])
 *   - values      := parsed CSV list from Values (row[3])
 *   - matchType   := based on Category/Attribute (default duty-exact)
 */
function loadAutomationChatSpaceDefinitions_() {
  const ssId = (CONFIG && CONFIG.AUTOMATION_SPREADSHEET_ID)
    ? String(CONFIG.AUTOMATION_SPREADSHEET_ID).trim()
    : '';

  if (!ssId) {
    Logger.info(JSON.stringify({
      level: 'WARN',
      message: 'Automation spreadsheet ID not configured; skipping automation-defined ChatSpaces.',
      data: {}
    }));
    return [];
  }

  const ss = SpreadsheetApp.openById(ssId);
  if (!ss) return [];

  const tabCandidates = [
    'Groups',
    'groups',
    'Automation',
    'automation',
    'Update Groups',
    'Group Automation',
    'Group Definitions'
  ];

  let sheet = null;
  for (let i = 0; i < tabCandidates.length; i++) {
    const s = ss.getSheetByName(tabCandidates[i]);
    if (s) { sheet = s; break; }
  }

  // Case-insensitive fallback, first sheet containing "group"
  if (!sheet) {
    const all = ss.getSheets();
    for (let i = 0; i < all.length; i++) {
      const nm = String(all[i].getName() || '').toLowerCase();
      if (nm.indexOf('group') !== -1) { sheet = all[i]; break; }
    }
  }

  if (!sheet) {
    Logger.info(JSON.stringify({
      level: 'WARN',
      message: 'No Groups-style tab found in automation spreadsheet.',
      data: { sheets: ss.getSheets().map(s => s.getName()) }
    }));
    return [];
  }

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  const header = (values[0] || []).map(h => String(h || '').trim().toLowerCase());

  function colIndex_(names) {
    for (let i = 0; i < names.length; i++) {
      const idx = header.indexOf(names[i]);
      if (idx >= 0) return idx;
    }
    return -1;
  }

  // Try headers first; if not present assume fixed columns.
  let idxCategory  = colIndex_(['category', 'type']);
  let idxAttr      = colIndex_(['attribute', 'attr']);
  let idxValues    = colIndex_(['values', 'value']);
  let idxDesc      = colIndex_(['description', 'display name', 'display_name', 'title']);
  let idxEnabled   = colIndex_(['enabled', 'active', 'create', 'sync']);
  let idxAddExt    = colIndex_(['add ext', 'addext', 'allow external', 'allow external users', 'external users']);

  const headerLooksLikeUpdateGroups =
    header.indexOf('category') !== -1 &&
    (header.indexOf('values') !== -1 || header.indexOf('value') !== -1) &&
    (header.indexOf('description') !== -1 || header.indexOf('display name') !== -1);

  if (!headerLooksLikeUpdateGroups) {
    idxCategory = 0;
    idxAttr     = 2;
    idxValues   = 3;
    idxDesc     = 4;
    if (idxAddExt < 0) idxAddExt = 5;
  }

  const defs = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r] || [];

    // Optional enabled gate
    if (idxEnabled >= 0) {
      const v = String(row[idxEnabled] || '').trim().toLowerCase();
      if (v === 'false' || v === 'no' || v === '0' || v === 'off') continue;
    }

    const category = String(row[idxCategory] || '').trim().toLowerCase();
    const attr     = String(row[idxAttr] || '').trim().toLowerCase();
    const rawVals  = (idxValues >= 0) ? row[idxValues] : '';
    const display  = String(row[idxDesc] || '').trim();
    const addExtRaw = (idxAddExt >= 0) ? row[idxAddExt] : '';

    const addExt = (function(v) {
      const t = String(v || '').trim().toLowerCase();
      return t === 'y' || t === 'yes' || t === 'true' || t === 'x';
    })(addExtRaw);

    // For ChatSpaces, we ONLY care about duty-based definitions.
    // Accept either:
    //  - category == 'duty-position'
    //  - or attr includes 'duty' (future-proof)
    const dutyBased = (category === 'duty-position') || (attr.indexOf('duty') !== -1);
    if (!dutyBased) continue;

    if (!display) continue;

    const list = parseCsvList_(rawVals);
    if (!list.length) continue;

    // Default match behavior: exact duty title match.
    // If you ever want contains-matching, set category to 'duty-position-contains' in the sheet.
    let matchType = 'duty-exact';
    if (category === 'duty-position-contains') matchType = 'duty-contains';

    defs.push({
      displayName: display.replace(/\s+/g, ' ').trim(),
      values: list,
      matchType: matchType,

      // External access gate is driven by Groups tab "Add EXT" column.
      // NOTE: externalUserAllowed is immutable and must be set at creation time.
      externalUserAllowed: addExt
    });
  }

  // De-dup by displayName
  const seen = {};
  const out = [];
  defs.forEach(d => {
    const k = String(d.displayName || '').toLowerCase();
    if (!k || seen[k]) return;
    seen[k] = true;
    out.push(d);
  });

  Logger.info(JSON.stringify({
    level: 'INFO',
    message: 'Automation ChatSpace sheet selected',
    data: { sheetName: sheet.getName(), defs: out.length }
  }));

  return out;
}

/**
 * List member emails for a Google Group.
 * Returns primaryEmail values (lowercased).
 */
function listGroupMemberEmails_(groupEmail) {
  const emails = [];
  let pageToken = '';

  try {
    do {
      const resp = AdminDirectory.Members.list(groupEmail, {
        maxResults: 200,
        pageToken: pageToken || undefined
      });

      (resp.members || []).forEach(function(m) {
        if (!m || !m.email) return;
        // Only include active members
        const st = String(m.status || '').toUpperCase();
        if (st && st !== 'ACTIVE') return;
        emails.push(String(m.email).toLowerCase());
      });

      pageToken = resp.nextPageToken;
    } while (pageToken);
  } catch (e) {
    Logger.info(JSON.stringify({
      level: 'WARN',
      message: 'Failed to list group members',
      data: { groupEmail: groupEmail, error: e.message }
    }));
  }

  return emails;
}
