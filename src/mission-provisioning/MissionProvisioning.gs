/**
 * Mission workspace provisioning webhook for FileMaker.
 * Version: 1.1
 * Filename: MissionProvisioning.gs
 * Saved: 2026-04-4 10:00 PDT
 *
 * FileMaker can POST JSON to the deployed Apps Script web app using
 * Insert from URL + cURL. Apps Script web app requests do not expose
 * custom HTTP headers reliably in `doPost(e)`, so this webhook expects
 * the shared secret in the JSON body as `secret`. The secret must NOT be
 * passed in the query string, since URLs are logged by servers/proxies.
 *
 * Required payload fields:
 *   - missionId
 *
 * Recommended payload fields:
 *   - missionName
 *   - missionCode
 *   - ownerEmail
 *   - parentFolderId
 *
 * Optional override fields:
 *   - groupEmailLocalPart
 *   - groupName
 *   - groupDescription
 *   - chatDisplayName
 *   - chatDescription
 *   - folderName
 *   - owners: string[]
 *   - managers: string[]
 *   - members: string[]
 *   - allowExternalMembers: boolean
 *   - shareFolderWithGroup: boolean
 */

const MISSION_PROVISIONING_SHEET_NAME = 'Mission Provisioning';
const MISSION_PARENT_FOLDER_PROPERTY = 'MISSION_PARENT_FOLDER_ID';
const MISSION_WEBHOOK_SECRET_PROPERTY = 'MISSION_WEBHOOK_SECRET';


function doGet(e) {
  // Anonymous, unauthenticated endpoint: expose only a static usage hint.
  // Do not reflect request parameters or disclose server-side configuration
  // state (whether the secret/parent folder are set), which would aid probing.
  const output = {
    ok: true,
    message: 'Mission provisioning web app is deployed. Send a POST request with JSON to provision a mission workspace.',
    expects: {
      method: 'POST',
      contentType: 'application/json',
      requiredFields: ['missionId'],
      secretLocation: 'json.secret'
    }
  };

  return missionProvisioningJsonResponse_(200, output);
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const payload = parseMissionProvisioningRequest_(e);
    validateMissionProvisioningSecret_(payload, e);

    const result = provisionMissionWorkspace_(payload);
    return missionProvisioningJsonResponse_(200, { ok: true, result: result });
  } catch (err) {
    Logger.error('Mission provisioning webhook failed', {
      errorMessage: err && err.message ? err.message : String(err)
    });
    return missionProvisioningJsonResponse_(500, {
      ok: false,
      error: err && err.message ? err.message : 'Unknown error'
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (ignored) {}
  }
}

function provisionMissionWorkspace_(payload) {
  const mission = buildMissionProvisioningSpec_(payload);
  const sheet = ensureMissionProvisioningSheet_();
  const existing = findMissionProvisioningRow_(sheet, mission.missionId);

  if (existing && String(existing.record.status || '').toUpperCase() === 'SUCCESS') {
    Logger.info('Mission already provisioned; returning existing resources', {
      missionId: mission.missionId
    });
    return existing.record;
  }

  const result = {
    missionId: mission.missionId,
    missionName: mission.missionName,
    missionCode: mission.missionCode,
    groupEmail: '',
    groupUrl: '',
    chatDisplayName: '',
    chatSpaceName: '',
    chatUrl: '',
    folderId: '',
    folderUrl: '',
    status: 'PENDING',
    processedAt: new Date().toISOString()
  };

  try {
    const group = ensureMissionGroup_(mission);
    result.groupEmail = group.groupEmail || '';
    result.groupUrl = group.groupUrl || '';

    const chat = ensureMissionChatSpace_(mission);
    result.chatDisplayName = chat.displayName || '';
    result.chatSpaceName = chat.spaceName || '';
    result.chatUrl = chat.chatUrl || '';

    const folder = ensureMissionFolder_(mission, result.groupEmail);
    result.folderId = folder.folderId || '';
    result.folderUrl = folder.folderUrl || '';

    result.status = 'SUCCESS';
    result.processedAt = new Date().toISOString();

    upsertMissionProvisioningRow_(sheet, mission, payload, result, existing ? existing.row : 0, '');
    Logger.info('Mission provisioning completed', {
      missionId: mission.missionId,
      groupEmail: result.groupEmail,
      chatSpaceName: result.chatSpaceName,
      folderId: result.folderId
    });

    return result;
  } catch (err) {
    result.status = 'ERROR';
    result.processedAt = new Date().toISOString();
    upsertMissionProvisioningRow_(
      sheet,
      mission,
      payload,
      result,
      existing ? existing.row : 0,
      err && err.message ? err.message : String(err)
    );
    throw err;
  }
}

function parseMissionProvisioningRequest_(e) {
  const raw = e && e.postData && e.postData.contents ? String(e.postData.contents) : '';
  if (!raw) throw new Error('Missing JSON request body');

  let payload;
  try {
    payload = JSON.parse(raw);
  } catch (err) {
    throw new Error('Request body must be valid JSON');
  }

  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
    throw new Error('JSON body must be an object');
  }

  return payload;
}

function validateMissionProvisioningSecret_(payload, e) {
  const expected = String(
    PropertiesService.getScriptProperties().getProperty(MISSION_WEBHOOK_SECRET_PROPERTY) || ''
  ).trim();

  if (!expected) {
    throw new Error(
      'Script property ' + MISSION_WEBHOOK_SECRET_PROPERTY + ' is not configured'
    );
  }

  // Secret must be supplied in the JSON body only. Accepting it in the query
  // string ('?key=' / '?secret=') would leak it into web-server, proxy, and
  // browser-history logs, so that path is intentionally not supported.
  const provided = String((payload && payload.secret) || '').trim();

  if (!provided || !constantTimeEquals_(provided, expected)) {
    throw new Error('Unauthorized');
  }
}

/**
 * Compares two strings in constant time to avoid leaking the secret's length
 * or content through timing differences.
 *
 * @param {string} a
 * @param {string} b
 * @returns {boolean} True if the strings are equal.
 */
function constantTimeEquals_(a, b) {
  a = String(a);
  b = String(b);
  // Fold the length difference into the comparison so mismatched lengths still
  // run the full loop and always return false.
  let diff = a.length ^ b.length;
  for (let i = 0; i < a.length; i++) {
    diff |= a.charCodeAt(i) ^ b.charCodeAt(i % b.length);
  }
  return diff === 0;
}

function buildMissionProvisioningSpec_(payload) {
  const missionId = String(payload.missionId || '').trim();
  if (!missionId) throw new Error('missionId is required');

  const missionName = String(payload.missionName || payload.name || missionId).trim();
  const missionCode = String(payload.missionCode || payload.code || '').trim();
  const defaultBase = missionCode || missionId || missionName;
  const slug = buildMissionSlug_(defaultBase);
  if (!slug) throw new Error('Could not derive a safe mission slug from missionId/missionCode/missionName');

  const ownerEmail = normalizeMissionEmail_(payload.ownerEmail);
  const owners = normalizeMissionEmailList_(payload.owners);
  const managers = normalizeMissionEmailList_(payload.managers);
  const members = normalizeMissionEmailList_(payload.members);

  if (ownerEmail && owners.indexOf(ownerEmail) === -1) {
    owners.unshift(ownerEmail);
  }

  const scriptProps = PropertiesService.getScriptProperties();
  const parentFolderId = String(
    payload.parentFolderId || scriptProps.getProperty(MISSION_PARENT_FOLDER_PROPERTY) || ''
  ).trim();

  if (!parentFolderId) {
    throw new Error(
      'parentFolderId is required, or set script property ' + MISSION_PARENT_FOLDER_PROPERTY
    );
  }

  const groupEmailLocalPart = normalizeGroupLocalPart_(
    payload.groupEmailLocalPart || ('mission.' + slug)
  );
  if (!groupEmailLocalPart) {
    throw new Error('Could not derive a valid group email local part');
  }

  const groupName = String(payload.groupName || ('Mission - ' + missionName)).trim();
  const groupDescription = String(
    payload.groupDescription ||
    ('Workspace group for mission ' + missionName + (missionCode ? ' (' + missionCode + ')' : ''))
  ).trim();

  const chatDisplayName = String(payload.chatDisplayName || groupName).trim();
  const chatDescription = String(
    payload.chatDescription ||
    ('This Chat space is for mission ' + missionName + ' and official Civil Air Patrol communications only.')
  ).trim();

  const folderName = String(payload.folderName || groupName).trim();

  return {
    missionId: missionId,
    missionName: missionName,
    missionCode: missionCode,
    groupEmailLocalPart: groupEmailLocalPart,
    groupName: groupName,
    groupDescription: groupDescription,
    chatDisplayName: chatDisplayName,
    chatDescription: chatDescription,
    folderName: folderName,
    parentFolderId: parentFolderId,
    ownerEmail: ownerEmail,
    owners: owners,
    managers: managers,
    members: members,
    allowExternalMembers: payload.allowExternalMembers === true,
    shareFolderWithGroup: payload.shareFolderWithGroup !== false
  };
}

function ensureMissionGroup_(mission) {
  const domain = String(CONFIG.EMAIL_DOMAIN || '').trim();
  if (!domain) throw new Error('CONFIG.EMAIL_DOMAIN is not configured');

  const groupEmail = mission.groupEmailLocalPart + domain;
  let group;

  try {
    group = AdminDirectory.Groups.get(groupEmail);
  } catch (err) {
    if (!(err && err.details && err.details.code === 404)) throw err;

    group = AdminDirectory.Groups.insert({
      email: groupEmail,
      name: mission.groupName,
      description: mission.groupDescription
    });

    Logger.info('Mission group created', {
      missionId: mission.missionId,
      groupEmail: groupEmail
    });
  }

  applyGroupMeta_(groupEmail, {
    name: mission.groupName,
    description: mission.groupDescription
  });

  applyManagedGroupSettings_(groupEmail, {
    allowExternalMembers: mission.allowExternalMembers,
    whoCanViewMembership: 'ALL_IN_DOMAIN_CAN_VIEW',
    whoCanPostMessage: 'ALL_IN_DOMAIN_CAN_POST'
  });

  ensureGroupMembers_(groupEmail, mission.owners, 'OWNER');
  ensureGroupMembers_(groupEmail, mission.managers, 'MANAGER');
  ensureGroupMembers_(groupEmail, mission.members, 'MEMBER');

  return {
    groupEmail: groupEmail,
    groupUrl: 'https://groups.google.com/a/' + CONFIG.DOMAIN + '/g/' + mission.groupEmailLocalPart
  };
}

function ensureMissionChatSpace_(mission) {
  let existingSpaces = buildChatSpaceCache();
  if (existingSpaces === null) {
    throw new Error('Chat API is not configured or authorized for this Apps Script project');
  }

  let spaceName = existingSpaces[mission.chatDisplayName];
  if (!spaceName) {
    spaceName = createChatSpace(mission.chatDisplayName, {
      historyOff: false,
      restrictPermissions: true,
      externalUserAllowed: mission.allowExternalMembers
    });
    if (!spaceName) {
      throw new Error('Failed to create Chat space for ' + mission.chatDisplayName);
    }
    existingSpaces[mission.chatDisplayName] = spaceName;
  }

  try {
    Chat.Spaces.patch(
      {
        spaceDetails: {
          description: mission.chatDescription,
          guidelines: 'This space is for official Civil Air Patrol communications only.'
        }
      },
      spaceName,
      { updateMask: 'spaceDetails' }
    );
  } catch (err) {
    Logger.warn('Could not update mission Chat space details', {
      missionId: mission.missionId,
      spaceName: spaceName,
      errorMessage: err.message
    });
  }

  const memberEmails = uniqueMissionEmails_(
    []
      .concat(mission.owners || [])
      .concat(mission.managers || [])
      .concat(mission.members || [])
  );

  const existingUserIds = listChatSpaceMemberUserIds(spaceName);
  const ownerUserIds = [];
  const managerUserIds = [];

  memberEmails.forEach(function(email) {
    const userId = getWorkspaceUserIdByEmail_(email);
    if (!userId) return;

    if (!existingUserIds.has(userId)) {
      addMemberToChatSpace(spaceName, userId);
    }

    if ((mission.owners || []).indexOf(email) > -1) ownerUserIds.push(userId);
    if ((mission.managers || []).indexOf(email) > -1) managerUserIds.push(userId);
  });

  ownerUserIds.concat(managerUserIds).forEach(function(userId) {
    promoteChatSpaceManager(spaceName, userId, {
      missionId: mission.missionId,
      displayName: mission.chatDisplayName
    });
  });

  return {
    displayName: mission.chatDisplayName,
    spaceName: spaceName,
    chatUrl: buildChatSpaceUrl_(spaceName)
  };
}

function ensureMissionFolder_(mission, groupEmail) {
  const folder = findOrCreateMissionFolder_(mission.parentFolderId, mission.folderName);

  if (mission.shareFolderWithGroup && groupEmail) {
    ensureDrivePermissionForGroup_(folder.id, groupEmail, 'writer');
  }

  return {
    folderId: folder.id,
    folderUrl: folder.webViewLink || ('https://drive.google.com/drive/folders/' + folder.id)
  };
}

function findOrCreateMissionFolder_(parentFolderId, folderName) {
  const escapedName = String(folderName || '').replace(/\\/g, '\\\\').replace(/'/g, "\\'");
  const query =
    "'" + parentFolderId + "' in parents and trashed=false and " +
    "mimeType='application/vnd.google-apps.folder' and name='" + escapedName + "'";

  const existing = Drive.Files.list({
    q: query,
    includeItemsFromAllDrives: true,
    supportsAllDrives: true,
    fields: 'files(id,name,webViewLink)'
  });

  if (existing && existing.files && existing.files.length) {
    return existing.files[0];
  }

  const folder = Drive.Files.create(
    {
      name: folderName,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [parentFolderId]
    },
    null,
    {
      supportsAllDrives: true,
      fields: 'id,name,webViewLink'
    }
  );

  Logger.info('Mission folder created', {
    folderId: folder.id,
    folderName: folderName,
    parentFolderId: parentFolderId
  });

  return folder;
}

function ensureDrivePermissionForGroup_(fileId, groupEmail, role) {
  role = String(role || 'writer').trim() || 'writer';

  try {
    const existing = Drive.Permissions.list(fileId, {
      supportsAllDrives: true,
      fields: 'permissions(id,type,emailAddress,role)'
    });

    const match = (existing.permissions || []).find(function(permission) {
      return permission &&
        permission.type === 'group' &&
        String(permission.emailAddress || '').toLowerCase() === String(groupEmail || '').toLowerCase();
    });

    if (match) {
      if (String(match.role || '') !== role) {
        Drive.Permissions.update(
          { role: role },
          fileId,
          match.id,
          { supportsAllDrives: true, fields: 'id,role' }
        );
      }
      return;
    }

    Drive.Permissions.create(
      {
        type: 'group',
        role: role,
        emailAddress: groupEmail
      },
      fileId,
      {
        supportsAllDrives: true,
        sendNotificationEmail: false,
        fields: 'id'
      }
    );
  } catch (err) {
    Logger.warn('Failed to apply Drive permission for mission group', {
      fileId: fileId,
      groupEmail: groupEmail,
      role: role,
      errorMessage: err.message
    });
  }
}

function ensureGroupMembers_(groupEmail, emails, role) {
  const normalized = uniqueMissionEmails_(emails || []);
  normalized.forEach(function(email) {
    try {
      AdminDirectory.Members.insert({
        email: email,
        role: role
      }, groupEmail);
    } catch (err) {
      if (err && err.details && err.details.code === 409) return;
      if (err && err.details && err.details.code === 404) {
        Logger.warn('Group member not found for mission provisioning', {
          groupEmail: groupEmail,
          email: email,
          role: role
        });
        return;
      }
      throw err;
    }
  });
}

function getWorkspaceUserIdByEmail_(email) {
  const normalized = normalizeMissionEmail_(email);
  if (!normalized) return '';

  try {
    const user = AdminDirectory.Users.get(normalized, {
      projection: 'basic',
      fields: 'id,primaryEmail'
    });
    return user && user.id ? String(user.id) : '';
  } catch (err) {
    Logger.warn('Could not resolve Workspace user for mission provisioning', {
      email: normalized,
      errorMessage: err.message
    });
    return '';
  }
}

function ensureMissionProvisioningSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(MISSION_PROVISIONING_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(MISSION_PROVISIONING_SHEET_NAME);
  }

  const headers = [[
    'Mission ID',
    'Mission Name',
    'Mission Code',
    'Status',
    'Group Email',
    'Chat Display Name',
    'Chat Space Name',
    'Folder ID',
    'Folder URL',
    'Processed At',
    'Error',
    'Request JSON',
    'Response JSON'
  ]];

  const current = sheet.getRange(1, 1, 1, headers[0].length).getValues()[0];
  const needsHeader = headers[0].some(function(header, idx) {
    return current[idx] !== header;
  });

  if (needsHeader) {
    sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    sheet.getRange(1, 1, 1, headers[0].length)
      .setFontWeight('bold')
      .setBackground('#1a73e8')
      .setFontColor('#ffffff');
  }

  return sheet;
}

function findMissionProvisioningRow_(sheet, missionId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const finder = sheet.getRange(2, 1, lastRow - 1, 1)
    .createTextFinder(String(missionId))
    .matchEntireCell(true)
    .findNext();

  if (!finder) return null;

  const row = finder.getRow();
  const values = sheet.getRange(row, 1, 1, 13).getValues()[0];
  return {
    row: row,
    record: {
      missionId: values[0] || '',
      missionName: values[1] || '',
      missionCode: values[2] || '',
      status: values[3] || '',
      groupEmail: values[4] || '',
      groupUrl: values[4] ? ('https://groups.google.com/a/' + CONFIG.DOMAIN + '/g/' + String(values[4]).split('@')[0]) : '',
      chatDisplayName: values[5] || '',
      chatSpaceName: values[6] || '',
      chatUrl: values[6] ? buildChatSpaceUrl_(values[6]) : '',
      folderId: values[7] || '',
      folderUrl: values[8] || '',
      processedAt: values[9] || '',
      error: values[10] || ''
    }
  };
}

function upsertMissionProvisioningRow_(sheet, mission, payload, result, row, errorMessage) {
  const sanitizedPayload = sanitizeMissionPayloadForLog_(payload);
  const values = [[
    mission.missionId,
    mission.missionName,
    mission.missionCode,
    result.status,
    result.groupEmail || '',
    result.chatDisplayName || '',
    result.chatSpaceName || '',
    result.folderId || '',
    result.folderUrl || '',
    result.processedAt || new Date().toISOString(),
    errorMessage || '',
    JSON.stringify(sanitizedPayload || {}),
    JSON.stringify(result || {})
  ]];

  if (row && row > 1) {
    sheet.getRange(row, 1, 1, values[0].length).setValues(values);
  } else {
    sheet.appendRow(values[0]);
  }
}

function missionProvisioningJsonResponse_(statusCode, payload) {
  const output = ContentService
    .createTextOutput(JSON.stringify(Object.assign({ statusCode: statusCode }, payload)))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

function sanitizeMissionPayloadForLog_(payload) {
  const clone = {};
  Object.keys(payload || {}).forEach(function(key) {
    if (key === 'secret') return;
    clone[key] = payload[key];
  });
  return clone;
}

function buildMissionSlug_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 48);
}

function normalizeGroupLocalPart_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/@.*$/, '')
    .replace(/[^a-z0-9._-]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function normalizeMissionEmail_(value) {
  return sanitizeEmail(String(value || '').trim()) || '';
}

function normalizeMissionEmailList_(value) {
  if (!value) return [];
  if (Array.isArray(value)) {
    return uniqueMissionEmails_(value);
  }
  return uniqueMissionEmails_(String(value).split(','));
}

function uniqueMissionEmails_(values) {
  const seen = {};
  const out = [];

  (values || []).forEach(function(value) {
    const email = normalizeMissionEmail_(value);
    if (!email || seen[email]) return;
    seen[email] = true;
    out.push(email);
  });

  return out;
}

function buildChatSpaceUrl_(spaceName) {
  const id = String(spaceName || '').split('/').pop();
  return id ? ('https://mail.google.com/chat/u/0/#chat/space/' + id) : '';
}

function testMissionProvisioningPayload_() {
  return provisionMissionWorkspace_({
    missionId: 'TEST-001',
    missionName: 'Mission Provisioning Test',
    missionCode: 'MISSION-TEST',
    ownerEmail: 'automation@pcr.cap.gov',
    secret: 'set via doPost only',
    shareFolderWithGroup: true
  });
}
