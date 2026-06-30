
/**
 * PCR-GW-SCRIPTS: ChatSpaces Module (Apps Script Version)
 *
 * Creates and manages Google Chat Spaces for:
 *  - CAP Committee spaces (from MbrCommittee)
 *  - Unit spaces (Region, Wing, Group, Squadron, Flight, Senior)
 *
 * Requirements:
 *  - Advanced Services: AdminDirectory, Chat
 *  - CONFIG and DRY_RUN globals (as in other modules)
 *  - getMembers(), getSquadrons(), parseFile() available (same as updateGroups)
 */

/**
 * Main entrypoint.
 */
function updateChatSpaces() {
  const start = new Date();
  Logger.info(JSON.stringify({
    level: 'INFO',
    message: 'Starting ChatSpaces sync',
    data: { dryRun: typeof DRY_RUN !== 'undefined' ? DRY_RUN : false }
  }));

const squadrons = getSquadronsForChatSpaces_();
const members   = getMembersForChatSpaces_(squadrons);

  const capidMaps      = buildWorkspaceCapidMaps(members);
  const existingSpaces = (typeof DRY_RUN !== 'undefined' && DRY_RUN) ? {} : buildChatSpaceCache();

  // Safety stop: if Chat app/API is not configured, buildChatSpaceCache() logs an error and returns {}.
  // In that situation, do NOT attempt to create hundreds of spaces.
  if ((typeof DRY_RUN === 'undefined' || !DRY_RUN) && Object.keys(existingSpaces).length === 0) {
    Logger.info(JSON.stringify({
      level: 'ERROR',
      message: 'Aborting ChatSpaces sync because Chat API/app is not configured (space cache empty).',
      data: {}
    }));
    return;
  }

  syncCommitteeChatSpaces(members, squadrons, capidMaps, existingSpaces);
  syncUnitChatSpaces(members, squadrons, capidMaps, existingSpaces);

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
    : ['CADET', 'SENIOR', 'FIFTY YEAR', 'LIFE', 'AEM'];

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
  let pageToken       = '';

  try {
    do {
      const resp = AdminDirectory.Users.list({
        domain: CONFIG.DOMAIN,
        maxResults: 500,
        projection: 'full',
        fields: 'users(id,primaryEmail,externalIds),nextPageToken',
        pageToken: pageToken || undefined
      });

      (resp.users || []).forEach(function(u) {
        const primaryEmail = (u.primaryEmail || '').toLowerCase();
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

  return { capidToUserId: capidToUserId, capidToEmail: capidToEmail };
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
  } catch (e) {
    Logger.info(JSON.stringify({
      level: 'ERROR',
      message: 'Failed to cache Chat Spaces',
      data: { error: e.message }
    }));
  }

  return cache;
}

/**
 * Committee Chat Spaces
 * Mirrors Python sync_chatspaces logic using MbrCommittee + Organization.
 */
function syncCommitteeChatSpaces(members, squadrons, capidMaps, existingSpaces) {
  const start = new Date();
  const rows  = parseFile('MbrCommittee') || [];

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

  if (typeof DRY_RUN !== 'undefined' && DRY_RUN) {
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
    const displayName = buildCommitteeSpaceName(org, entry.committee);

    Logger.info(JSON.stringify({
      level: 'INFO',
      message: 'Processing committee ChatSpace',
      data: { displayName: displayName, orgid: entry.orgid }
    }));

    // Ensure space exists
    let spaceName = existingSpaces[displayName];
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

  if (typeof DRY_RUN !== 'undefined' && DRY_RUN) {
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
    const space = Chat.Spaces.create({
      displayName: displayName,
      spaceType: 'SPACE'
    });
    spaceName = space.name;

    if (historyOff) {
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