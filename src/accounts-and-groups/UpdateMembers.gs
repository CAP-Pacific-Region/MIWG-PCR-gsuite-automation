/**
 * -------------------------------------------------------------------------
 * Version: 1.4.5
 * Date: 2026-07-09
 * Authors: Michigan Wing (MIWG) — Extended and Maintained by Lt Col Noel Luneau
 * Changes: AdminDirectory.Users.list standardized to customer:"my_customer";
 *   testImpersonationToken uses console.log (Logger is overridden). See PCR_CHANGELOG.md.
 *
 * Description:
 * - Added CADET_LITE filtering logic and configuration controls.
 * - Unified cadet duty position handling with senior duty logic, including
 *   matching field structure, assistant indicators, dutyPositionIds, and
 *   dutyPositionIdsAndLevel.
 * - Updated "Member" default organization title to "No Duty Assignment" for
 *   both seniors and cadets.
 * - Standardized Directory and Gmail Send-As displayName generation.
 * - Improved duty position parsing consistency between cadet and senior flows.
 * - Multiple reliability improvements to user updates, including better
 *   rank handling, fullName consistency, and org metadata population.
 * - Fixed Manual Members to process via an OU added to OrgPaths.txt like
 *   434, PCR-001.
 * - Updated manager email assignment to follow the CAPWATCH command hierarchy
 *   so commanders report to the nearest parent-org commander instead of themselves.
 * - Added managerEmail to member change detection so manager relation updates
 *   are included in normal Workspace sync runs.
 * - Updated recovery contact assignment to use CAPWATCH secondary email, then
 *   cadet parent email, and recovery phone precedence of member cell phone,
 *   then cadet parent phone.
 * - Updated Workspace other email to use CAPWATCH EMAIL SECONDARY instead of
 *   EMAIL PRIMARY.
 * - Allowed recovery email sources to use DoNotContact rows while keeping
 *   Workspace other email restricted to contactable EMAIL SECONDARY rows.
 * - Updated temporary password generation to use WING + script generation date
 *   + CAPID so non-CAPWATCH accounts do not produce invalid date passwords.
 * - Updated manager email assignment to resolve commanders by CAPID using the
 *   Workspace primary email map before falling back to generated email.
 * -------------------------------------------------------------------------
 */

/**
 * Member Synchronization Module
 *
 * Manages synchronization between CAPWATCH data and Google Workspace:
 * - Retrieves and parses CAPWATCH member data
 * - Creates/updates Google Workspace user accounts
 * - Manages email aliases
 * - Suspends expired members
 * - Reactivates renewed members (including archived)
 * - Tracks changes for efficient updates
 */

/**
 * Gets all squadrons for the configured wing from CAPWATCH data
 * Includes both regular squadrons and special units (e.g., AEM)
 *
 * @returns {Object} Squadron data indexed by orgid with properties:
 *   - orgid: Organization ID
 *   - name: Squadron name
 *   - charter: Charter number (e.g., "NER-MI-100")
 *   - unit: Unit number
 *   - nextLevel: Parent organization ID
 *   - scope: Organization scope (UNIT, GROUP, WING)
 *   - wing: Wing abbreviation
 *   - orgPath: Google Workspace organizational unit path
 */

 // GLOBAL — CAPID → existing Workspace primary email
 let workspaceEmailByCapid = {};

function getSquadrons() {
  let squadrons = {};
  let squadronData = parseFile('Organization');

  for (let i = 0; i < squadronData.length; i++) {
    if (squadronData[i][2] === CONFIG.WING) {
      squadrons[squadronData[i][0]] = {
        orgid: squadronData[i][0],
        name: squadronData[i][5],
        charter: Utilities.formatString("%s-%s-%03d", squadronData[i][1], squadronData[i][2], squadronData[i][3]),
        unit: squadronData[i][3],
        nextLevel: squadronData[i][4],
        scope: squadronData[i][9],
        wing: squadronData[i][2],
        orgPath: ''
      }
    }
  }

  // Create artificial AEM Unit using MIWG as template
  squadrons[CONFIG.SPECIAL_ORGS.AEM_UNIT] = Object.assign(
    {},
    squadrons[CONFIG.CAPWATCH_ORGID],
    {
      orgid: CONFIG.SPECIAL_ORGS.AEM_UNIT,
      name: "Aerospace Education Members"
    }
  );

  // Add organizational unit paths from OrgPaths file
  let orgPaths = parseFile('OrgPaths');
  for (let i = 0; i < orgPaths.length; i++) {
    if (squadrons[orgPaths[i][0]]) {
      squadrons[orgPaths[i][0]].orgPath = orgPaths[i][1];
    }
  }

  return squadrons;
}

/**
 * Retrieves and processes member data from CAPWATCH files
 *
 * This is the main data retrieval function that:
 * 1. Parses member data from CAPWATCH files
 * 2. Filters by member type and status
 * 3. Validates member data
 * 4. Adds contact information
 * 5. Optionally adds duty positions
 *
 * @param {string[]} types - Member types to include (default: all active types)
 * @param {boolean} includeDutyPositions - Whether to parse duty positions (default: true)
 * @returns {Object} Members object indexed by CAPID
 */

/**
 * Loads manual members from the 'ManualMembers' sheet and returns an object indexed by CAPID.
 * Each object is formatted to match the structure of CAPWATCH members.
 *
 * @param {Object} squadrons - Squadron lookup object
 * @returns {Object} Manual members indexed by CAPID
 */
function loadManualMembers(squadrons) {
  // Accept either tab name to avoid silent mismatches
  const ss = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Manual Members') || ss.getSheetByName('ManualMembers');
  if (!sheet) {
    Logger.warn("Manual members sheet not found", {
      spreadsheetId: CONFIG.AUTOMATION_SPREADSHEET_ID,
      expectedTabs: ['Manual Members', 'ManualMembers']
    });
    return {};
  }

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) {
    Logger.info('Manual members sheet is empty (no data rows)', {
      tabName: sheet.getName(),
      rowCount: values ? values.length : 0
    });
    return {};
  }

  // Normalize headers to avoid issues with trailing spaces / different casing
  const rawHeader = values[0];
  const header = rawHeader.map(h => String(h || '').trim());

  // Build OrgID → orgPath lookup from OrgPaths.txt so manual members can reference
  // orgs outside this wing (e.g., PCR, NHQ) as long as an orgPath mapping exists.
  const orgPathsMap = {};
  try {
    const orgPathsRows = parseFile('OrgPaths');
    for (let i = 0; i < orgPathsRows.length; i++) {
      const oid = String(orgPathsRows[i][0] || '').trim();
      const path = String(orgPathsRows[i][1] || '').trim();
      if (oid && path) orgPathsMap[oid] = path;
    }
  } catch (e) {
    Logger.warn('Unable to load OrgPaths for manual member fallback', {
      errorMessage: e && e.message ? e.message : String(e)
    });
  }

  // Build OrgID → { name, charter } lookup from Organization.txt
  // This allows manual members in external orgs (e.g., PCR, NHQ) to get correct orgName/charter
  // even when `getSquadrons()` is wing-scoped.
  const orgMetaMap = {};
  try {
    const orgRows = parseFile('Organization');
    for (let i = 0; i < orgRows.length; i++) {
      const oid = String(orgRows[i][0] || '').trim();
      if (!oid) continue;

      // Organization.txt columns used elsewhere in this file:
      // [1]=Region, [2]=Wing, [3]=Unit number, [5]=Org Name
      const region = String(orgRows[i][1] || '').trim();
      const wing = String(orgRows[i][2] || '').trim();
      const unitNum = orgRows[i][3];
      const name = String(orgRows[i][5] || '').trim();

      let charter = '';
      try {
        charter = Utilities.formatString('%s-%s-%03d', region, wing, unitNum);
      } catch (ignored) {
        charter = '';
      }

      orgMetaMap[oid] = { name: name, charter: charter };
    }
  } catch (e) {
    Logger.warn('Unable to load Organization metadata for manual member fallback', {
      errorMessage: e && e.message ? e.message : String(e)
    });
  }

  const members = {};
  let added = 0;
  let skippedNoCapid = 0;
  let skippedNoOrg = 0;

  for (let i = 1; i < values.length; i++) {
    const row = {};
    for (let j = 0; j < header.length; j++) {
      const key = header[j];
      if (!key) continue;
      row[key] = values[i][j];
    }

    const capid = String(row['CAPID'] || '').trim();
    if (!capid) {
      skippedNoCapid++;
      continue;
    }

    // Prefer wing-scoped squadrons lookup, but allow fallback to OrgPaths mapping
    // so manual members can be placed into external org OUs (e.g., PCR, NHQ).
    const orgid = String(row['OrgID'] || '').trim();
    let squadron = squadrons[orgid];

    if (!squadron) {
      const fallbackOrgPath = orgPathsMap[orgid];
      if (fallbackOrgPath) {
        // Best-effort metadata for downstream display.
        // orgPath is the only required field for provisioning placement.
        const meta = orgMetaMap[orgid] || {};
        squadron = {
          charter: String(meta.charter || row['Charter'] || row['charter'] || 'External'),
          name: String(meta.name || row['OrgName'] || row['OrganizationName'] || row['UnitName'] || 'External Organization'),
          orgPath: fallbackOrgPath
        };

        Logger.info('Manual member using OrgPaths fallback', {
          capid: capid,
          orgid: orgid,
          orgPath: fallbackOrgPath
        });
      } else {
        skippedNoOrg++;
        Logger.warn('Manual member skipped (OrgID not found in squadrons lookup and no OrgPaths mapping)', {
          capid: capid,
          orgid: orgid,
          lastName: row['LastName'] || '',
          firstName: row['FirstName'] || ''
        });
        continue;
      }
    }

    const dutyId = String(row['DutyID'] || '').trim();
    const isAssistant = String(row['Assistant'] || '').trim() === '1';

    // Manual email handling
    // Email           -> member.secondaryEmail (welcome + template only)
    // Secondary Email -> member.email (Workspace Email (Other))
    const secondaryEmail = String(row['Email'] || '').trim();
    const workspaceEmail = String(row['Secondary Email'] || '').trim();

    // Manual phone handling
    const primaryPhoneRaw = String(row['Primary Phone'] || '').trim();

    // Normalize phone to E.164 +1XXXXXXXXXX if possible
    let primaryPhone = '';
    if (primaryPhoneRaw) {
      const digits = primaryPhoneRaw.replace(/\D/g, '');
      if (digits.length >= 10) {
        primaryPhone = `+1${digits.slice(-10)}`;
      }
    }

    members[capid] = {
      capsn: capid,
      lastName: row['LastName'] || '',
      firstName: row['FirstName'] || '',
      orgid: orgid,
      group: (squadrons[orgid] && typeof calculateGroup === 'function') ? calculateGroup(orgid, squadrons) : '',
      charter: squadron.charter,
      orgName: squadron.name,
      rank: String(row['Rank'] || '').trim(),
      type: row['Type'] || 'SENIOR',
      status: row['Status'] || 'ACTIVE',
      middleName: row['MiddleName'] || '',
      suffix: row['Suffix'] || '',
      modified: new Date().toISOString(),
      orgPath: (squadron && squadron.orgPath) ? squadron.orgPath : (orgPathsMap[orgid] || ''),
      email: workspaceEmail || null,
      secondaryEmail: secondaryEmail || null,
      phone: primaryPhone || null,

      dutyPositions: dutyId ? [{ id: dutyId, assistant: isAssistant }] : [],
      dutyPositionIds: dutyId ? [dutyId] : [],
      // Manual members do not have a CAPWATCH level. Keep a placeholder level for parity.
      dutyPositionIdsAndLevel: dutyId ? [dutyId + '_4'] : []
    };

    added++;
  }

  Logger.info('Manual members loaded from sheet', {
    tabName: sheet.getName(),
    added: added,
    skippedNoCapid: skippedNoCapid,
    skippedNoOrg: skippedNoOrg
  });

  return members;
}

function getMembers(types = CONFIG.MEMBER_TYPES.ACTIVE, includeDutyPositions = true) {
  const start = new Date();
  const members = {};
  const squadrons = getSquadrons();

  Logger.info('Starting member data retrieval', { types: types });

  // Build member objects from Member.txt
  const memberData = parseFile('Member');
  let processedCount = 0;

  for (let i = 0; i < memberData.length; i++) {
    if (shouldProcessMember(memberData[i], types)) {
      const member = createMemberObject(memberData[i], squadrons);

      // Validate before adding
      const validation = validateMember(member);
      if (validation.isValid) {
        members[memberData[i][0]] = member;
        processedCount++;
      } else {
        Logger.warn('Invalid member data', {
          capsn: memberData[i][0],
          errors: validation.errors
        });
      }
    }
  }

  Logger.info('Members parsed', {
    count: processedCount,
    duration: new Date() - start + 'ms'
  });

  // Add contact information from MbrContact.txt
  const contactStart = new Date();
  addContactInfo(members, parseFile('MbrContact'));
  Logger.info('Contact info added', {
    duration: new Date() - contactStart + 'ms'
  });

  // Add duty positions if requested
  if (includeDutyPositions) {
    const dutyStart = new Date();
    addDutyPositions(members, parseFile('DutyPosition'), squadrons);
    addCadetDutyPositions(members, parseFile('CadetDutyPositions'), squadrons);
    Logger.info('Duty positions added', {
      duration: new Date() - dutyStart + 'ms'
    });
  }

  // Merge in manual members from the ManualMembers sheet
  const manualMembers = loadManualMembers(squadrons);
  Object.assign(members, manualMembers);
  Logger.info('Manual members merged', { count: Object.keys(manualMembers).length });

  Logger.info('Member retrieval completed', {
    totalMembers: Object.keys(members).length,
    totalDuration: new Date() - start + 'ms'
  });

  return members;
}

/**
 * Determines if a member should be processed based on status and type
 *
 * Excludes members in CONFIG.EXCLUDED_ORG_IDS (CA's Unit 000/999 holding
 * squadrons) by ORGID rather than by raw Unit number, so members are never
 * eligible for accounts based on which physical org they're parked in —
 * regardless of member type (including FIFTY YEAR/INDEFINITE, which never
 * expire but are still excluded if held in one of those orgs).
 *
 * @param {Array} memberRow - Raw member data row from CSV
 * @param {string[]} types - Valid member types to include
 * @returns {boolean} True if member should be processed
 */
function shouldProcessMember(memberRow, types) {

// ----- CADET-LITE FILTER (NEW + LOGGING) -----
if (CONFIG.CADET_LITE === true) {
  const rank = (memberRow[14] || '').trim();
  if (CONFIG.CADET_LITE_EXCLUDED_GRADES.indexOf(rank) > -1) {

    // Log exclusion
    Logger.info("Cadet-Lite: excluded low-grade cadet", {
      capid: memberRow[0],
      name: memberRow[3] + " " + memberRow[2],
      rank: rank,
      orgid: memberRow[11]
    });

    // Track counter for summary
    if (!globalThis.cadetLiteExcludedCount) {
      globalThis.cadetLiteExcludedCount = 0;
    }
    globalThis.cadetLiteExcludedCount++;

    return false;
  }
}
// ---------------------------------------------

  return memberRow[24] === 'ACTIVE' &&
         CONFIG.EXCLUDED_ORG_IDS.indexOf(String(memberRow[11])) === -1 &&
         types.indexOf(memberRow[21]) > -1;
}

/**
 * Creates a structured member object from raw CAPWATCH data
 *
 * @param {Array} memberRow - Raw member data row from CSV
 * @param {Object} squadrons - Squadron lookup object
 * @returns {Object} Formatted member object with all required fields
 */
function createMemberObject(memberRow, squadrons) {
  return {
    capsn: memberRow[0],
    lastName: memberRow[2],
    firstName: memberRow[3],
    middleName: memberRow[4] || '',
    suffix: memberRow[5] || '',
    orgid: memberRow[11],
    group: calculateGroup(memberRow[11], squadrons),
    charter: squadrons[memberRow[11]].charter,
    orgName: squadrons[memberRow[11]].name,
    rank: memberRow[14],
    type: memberRow[21],
    status: memberRow[24],
    joined: memberRow[15],
    modified: memberRow[19],
    orgPath: squadrons[memberRow[11]].orgPath,
    email: null,
    dutyPositions: [],
    dutyPositionIds: [],
    dutyPositionIdsAndLevel: []
  };
}

function addContactInfo(members, contactData) {
  for (let i = 0; i < contactData.length; i++) {
    const capid = contactData[i][0];
    const type = contactData[i][1]?.toUpperCase() || '';
    const priority = contactData[i][2]?.toUpperCase() || '';
    const contact = contactData[i][3]?.trim() || '';
    const doNotContact = contactData[i][6]?.toUpperCase() === 'TRUE';

    if (!members[capid]) continue;

    if (type === 'EMAIL') {
      const sanitized = sanitizeEmail(contact);
      if (!sanitized) continue;

      if (priority === 'PRIMARY') {
        if (!doNotContact) members[capid].email = sanitized;
      } else if (priority === 'SECONDARY') {
        members[capid].secondaryEmail = sanitized;
        if (!doNotContact) members[capid].otherEmail = sanitized;
      }
    } else if (members[capid].type === 'CADET' && type === 'CADET PARENT EMAIL') {
      const sanitized = sanitizeEmail(contact);
      if (!sanitized) continue;
      if (!members[capid].parentEmail || priority === 'PRIMARY') {
        members[capid].parentEmail = sanitized;
      }
    }

    const digits = contact.replace(/\D/g, '');
    if (digits.length < 10) continue;
    if (doNotContact) continue;

    const normalizedPhone = `+1${digits.slice(-10)}`;
    const isMemberCellPhone = !type.includes('PARENT') &&
      (type === 'CELL PHONE' || type === 'MOBILE PHONE' || type.includes('CELL'));
    const isCadetParentPhone = members[capid].type === 'CADET' && type === 'CADET PARENT PHONE';

    if (isMemberCellPhone && (
      members[capid].phoneSource !== 'MEMBER_CELL' ||
      priority === 'PRIMARY'
    )) {
      members[capid].phone = normalizedPhone;
      members[capid].phoneSource = 'MEMBER_CELL';
    } else if (isCadetParentPhone && !members[capid].phone) {
      members[capid].phone = normalizedPhone;
      members[capid].phoneSource = 'CADET_PARENT';
    }
  }

  Logger.info('Contact info added (email + phone)', {
    totalMembers: Object.keys(members).length
  });
}

/**
 * Adds senior member duty positions to member objects
 *
 * @param {Object} members - Members object indexed by CAPID
 * @param {Array} dutyPositionData - Parsed duty position data
 * @param {Object} squadrons - Squadron lookup object
 * @returns {void}
 */

function addDutyPositions(members, dutyPositionData, squadrons) {
  for (let i = 0; i < dutyPositionData.length; i++) {
    if (members[dutyPositionData[i][0]]) {
      let dutyPositionID = dutyPositionData[i][1].trim();
      const indicator = dutyPositionData[i][4] == '1' ? ' (A)' : ' (P)';
      const charter = squadrons[dutyPositionData[i][7]] ? squadrons[dutyPositionData[i][7]].charter : 'Unknown';
      members[dutyPositionData[i][0]].dutyPositions.push({
        value: dutyPositionID + indicator + ' (' + charter + ')',
        id: dutyPositionID,
        level: dutyPositionData[i][3],
        assistant: dutyPositionData[i][4] == '1'
      });
      members[dutyPositionData[i][0]].dutyPositionIds.push(dutyPositionID);
      members[dutyPositionData[i][0]].dutyPositionIdsAndLevel.push(
        dutyPositionID + '_' + dutyPositionData[i][3]
      );
    }
  }
}

/**
 * Assigns manager email for each member based on the CAPWATCH command hierarchy.
 * Members report to their own org commander. Commanders report to the nearest
 * parent org commander instead of reporting to themselves.
 *
 * @param {Object} members - Members object indexed by CAPID
 */
function assignManagerEmails(members) {
  const commandersData = parseFile('Commanders');
  const organizationData = parseFile('Organization');
  const commanderByOrg = {};
  const parentOrgByOrg = {};

  for (let i = 0; i < organizationData.length; i++) {
    const orgid = String(organizationData[i][0] || '').trim();
    const parentOrgid = String(organizationData[i][4] || '').trim();
    if (orgid) parentOrgByOrg[orgid] = parentOrgid;
  }

  // Commanders.txt: ORGID = col 1, commander CAPID = col 5
  for (let i = 0; i < commandersData.length; i++) {
    const orgid = String(commandersData[i][0] || '').trim();
    const commanderCAPID = String(commandersData[i][4] || '').trim();
    if (orgid && commanderCAPID && members[commanderCAPID]) {
      commanderByOrg[orgid] = commanderCAPID;
    }
  }

  function getWorkspaceEmailForCapid(capid) {
    const mappedEmail = workspaceEmailByCapid[String(capid)];
    if (mappedEmail) return String(mappedEmail).toLowerCase();

    const member = members[capid];
    if (!member) return '';

    return [
      String(member.firstName || '').toLowerCase().replace(/\s+/g, ''),
      '.',
      String(member.lastName || '').toLowerCase().replace(/\s+/g, ''),
      CONFIG.EMAIL_DOMAIN
    ].join('');
  }

  function findManagerCapid(memberCapid, orgid) {
    let currentOrgid = String(orgid || '').trim();
    const visitedOrgids = {};

    while (currentOrgid && !visitedOrgids[currentOrgid]) {
      visitedOrgids[currentOrgid] = true;

      const commanderCAPID = commanderByOrg[currentOrgid];
      if (commanderCAPID && commanderCAPID !== memberCapid) {
        return commanderCAPID;
      }

      currentOrgid = parentOrgByOrg[currentOrgid] || '';
    }

    return '';
  }

  let assignedCount = 0;
  for (const capid in members) {
    const member = members[capid];
    const managerCapid = findManagerCapid(String(capid), member.orgid);
    member.managerEmail = managerCapid ? getWorkspaceEmailForCapid(managerCapid) : '';
    if (member.managerEmail) assignedCount++;
  }

  Logger.info('Manager emails assigned', {
    membersAssigned: assignedCount,
    commandersLoaded: Object.keys(commanderByOrg).length
  });
}

/**
 * Adds cadet duty positions to member objects
 *
 * @param {Object} members - Members object indexed by CAPID
 * @param {Array} cadetDutyPositionData - Parsed cadet duty position data
 * @param {Object} squadrons - Squadron lookup object
 * @returns {void}
 */
function addCadetDutyPositions(members, cadetDutyPositionData, squadrons) {
  for (let i = 0; i < cadetDutyPositionData.length; i++) {
    const capid = cadetDutyPositionData[i][0];
    if (!members[capid]) continue;

    const dutyId   = (cadetDutyPositionData[i][1] || '').trim(); // e.g. "Cadet IT Officer"
    const level    = cadetDutyPositionData[i][3] || '';          // e.g. "UNIT", "WING"
    const isAsst   = cadetDutyPositionData[i][4] == '1';         // "1" = assistant
    const orgid    = cadetDutyPositionData[i][7];                // org key for charter lookup
    const squadron = squadrons[orgid];
    const charter  = squadron ? squadron.charter : 'Unknown';

    // Build the same style "value" string as seniors
    const indicator = isAsst ? ' (A)' : ' (P)';
    const value = dutyId + indicator + ' (' + charter + ')';

    members[capid].dutyPositions.push({
      value: value,
      id: dutyId,
      level: level,
      assistant: isAsst
    });

    // Keep these in sync with seniors too
    members[capid].dutyPositionIds.push(dutyId);
    members[capid].dutyPositionIdsAndLevel.push(dutyId + '_' + level);
  }
}

/**
 * Retrieves Aerospace Education Members only
 * Convenience function that calls getMembers with AEM filter
 *
 * @returns {Object} AEM members object indexed by CAPID
 */
function getAEMembers() {
  return getMembers(CONFIG.MEMBER_TYPES.AEM_ONLY, false);
}

/**
 * Retrieves previously saved member data from Drive
 * Used to detect changes and avoid unnecessary API calls
 *
 * @returns {Object} Previously saved member data or empty object
 */
function getCurrentMemberData() {
  let folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  let files = folder.getFilesByName('CurrentMembers.txt');

  if (files.hasNext()) {
    let content = files.next().getBlob().getDataAsString();
    if (content) {
      try {
        return JSON.parse(content);
      } catch (e) {
        Logger.error('Failed to parse CurrentMembers.txt', { errorMessage: e.message });
        return {};
      }
    }
  }

  Logger.warn('CurrentMembers.txt not found or empty');
  return {};
}

/**
 * Saves current member data to Drive for change detection
 *
 * @param {Object} currentMembers - Current member data to save
 * @returns {void}
 */
function saveCurrentMemberData(currentMembers) {
  let folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  let files = folder.getFilesByName('CurrentMembers.txt');

  if (files.hasNext()) {
    let file = files.next();
    let content = JSON.stringify(currentMembers);
    file.setContent(content);
    Logger.info('Current member data saved', {
      memberCount: Object.keys(currentMembers).length,
      fileName: 'CurrentMembers.txt'
    });
  } else {
    Logger.warn('CurrentMembers.txt not found - cannot save', {
      folderId: CONFIG.CAPWATCH_DATA_FOLDER_ID
    });
  }
}

// Modified by Lt Col Noel Luneau on 2025-11-22 – added secondaryEmail and mobilePhone tracking
/**
 * Checks if a member's data has changed since last update
 * Compares rank, charter, duty positions, status, email, phone, and manager email
 *
 * @param {Object} newMember - New member data
 * @param {Object} previousMember - Previously saved member data
 * @returns {boolean} True if member data has changed or is new
 */
function memberUpdated(newMember, previousMember) {
  return (!newMember || !previousMember) ||
         (newMember.rank !== previousMember.rank ||
          newMember.charter !== previousMember.charter ||
          JSON.stringify(newMember.dutyPositions) !== JSON.stringify(previousMember.dutyPositions) ||
          newMember.status !== previousMember.status ||
          newMember.email !== previousMember.email ||
          newMember.secondaryEmail !== previousMember.secondaryEmail ||
          newMember.otherEmail !== previousMember.otherEmail ||
          newMember.parentEmail !== previousMember.parentEmail ||
          newMember.phone !== previousMember.phone ||
          newMember.managerEmail !== previousMember.managerEmail);
}

function getPrimaryOrgForLog_(user) {
  const organizations = user && user.organizations ? user.organizations : [];
  return organizations.find(org => org.primary) || organizations[0] || {};
}

function getRelationForLog_(user, type) {
  const relations = user && user.relations ? user.relations : [];
  const relation = relations.find(item => item.type === type);
  return relation ? String(relation.value || '') : '';
}

function getPhoneForLog_(user, type) {
  const phones = user && user.phones ? user.phones : [];
  const phone = phones.find(item => item.type === type);
  return phone ? String(phone.value || '') : '';
}

function getEmailForLog_(user, type) {
  const emails = user && user.emails ? user.emails : [];
  const email = emails.find(item => item.type === type);
  return email ? String(email.address || '') : '';
}

function getGradeForLog_(user) {
  return String(
    user &&
    user.customSchemas &&
    user.customSchemas.CAPWATCH &&
    user.customSchemas.CAPWATCH.Grade
      ? user.customSchemas.CAPWATCH.Grade
      : ''
  );
}

function addChangeForLog_(changes, field, currentValue, desiredValue) {
  const current = String(currentValue || '');
  const desired = String(desiredValue || '');
  if (current !== desired) {
    changes.push({
      field: field,
      current: current,
      desired: desired
    });
  }
}

function logWorkspaceUserUpdateDiff_(primaryEmail, member, updates) {
  let existing;
  try {
    existing = AdminDirectory.Users.get(primaryEmail, { projection: 'full' });
  } catch (e) {
    Logger.warn('Unable to read current user before update diff logging', {
      email: primaryEmail,
      capsn: member.capsn,
      errorMessage: e.message
    });
    return;
  }

  const existingOrg = getPrimaryOrgForLog_(existing);
  const desiredOrg = updates.organizations && updates.organizations.length
    ? updates.organizations[0]
    : {};
  const desiredManager = updates.relations && updates.relations.length
    ? updates.relations[0].value
    : '';

  const changes = [];
  addChangeForLog_(changes, 'recoveryEmail', existing.recoveryEmail, updates.recoveryEmail);
  addChangeForLog_(changes, 'recoveryPhone', existing.recoveryPhone, updates.recoveryPhone);
  addChangeForLog_(changes, 'mobilePhone', getPhoneForLog_(existing, 'mobile'), member.phone);
  addChangeForLog_(changes, 'otherEmail', getEmailForLog_(existing, 'other'), member.otherEmail);
  addChangeForLog_(changes, 'managerEmail', getRelationForLog_(existing, 'manager'), desiredManager);
  addChangeForLog_(changes, 'orgUnitPath', existing.orgUnitPath, updates.orgUnitPath);
  addChangeForLog_(changes, 'employeeTitle', existingOrg.title, desiredOrg.title);
  addChangeForLog_(changes, 'department', existingOrg.department, desiredOrg.department);
  addChangeForLog_(changes, 'costCenter', existingOrg.costCenter, desiredOrg.costCenter);
  addChangeForLog_(changes, 'displayName', existing.name && existing.name.displayName, updates.name.displayName);
  addChangeForLog_(changes, 'suspended', existing.suspended, updates.suspended);
  addChangeForLog_(changes, 'CAPWATCH.Grade', getGradeForLog_(existing), member.rank);

  Logger.info('Workspace user update diff', {
    email: primaryEmail,
    capsn: member.capsn,
    changeCount: changes.length,
    changes: changes
  });
}

/**
 * Updates or creates a Google Workspace user for a CAP member
 *
 * Process:
 * 1. Attempts to update existing user
 * 2. If not found, creates new user
 * 3. Adds email alias for new users
 * 4. Suspends users in excluded organizations
 * 5. Fixes existing member accounts with .2, .3, etc
 *
 * @param {Object} member - Member object containing CAP data
 * @returns {void}
 */


function addOrUpdateUser(member) {
  const baseEmail = `${member.firstName}.${member.lastName}`.toLowerCase().replace(/\s+/g, '');
  let primaryEmail =
    workspaceEmailByCapid[member.capsn] ||
    (baseEmail + CONFIG.EMAIL_DOMAIN);

  if (workspaceEmailByCapid[member.capsn]) {
    Logger.info('Preserving existing primary email', {
      capsn: member.capsn,
      email: primaryEmail
    });
  }

  let user;

  // Build Gmail/Directory Send-As display name once
  const sendAsDisplayName = [
    member.lastName + (member.suffix ? ' ' + member.suffix : ''),
    ', ',
    member.firstName,
    member.middleName ? ' ' + member.middleName.charAt(0) : '',
    member.rank ? ' ' + member.rank : ''
  ].join('').trim();

  let updates = {
    employeeId: String(member.capsn),
    externalIds: [{ type: 'organization', value: String(member.capsn) }],
    organizations: [{
      title: member.dutyPositions && member.dutyPositions.length > 0
        ? member.dutyPositions
            .map(d => d.assistant ? d.id + ' (A)' : d.id)
            .join(', ')
        : 'No Duty Assignment',
      // Squadron / Unit Name (Organization.name)
      department: toTitleCase(member.orgName || member.charter || ''),
      // Charter (PCR-HI-077 etc.)
      costCenter: member.charter || '',
      description: member.type || '',
      type: 'work',
      primary: true
    }],
    orgUnitPath: member.orgPath,
    recoveryEmail: member.secondaryEmail || member.parentEmail || '',
    recoveryPhone: member.phone || '',
    phones: member.phone ? [{ type: 'mobile', value: member.phone }] : [],
    emails: member.otherEmail ? [{ type: 'other', address: member.otherEmail }] : [],
    suspended: CONFIG.EXCLUDED_ORG_IDS.includes(String(member.orgid)),
    name: {
      givenName: member.firstName,
      familyName: member.lastName,
      middleName: member.middleName || '',
      fullName: [
        member.firstName,
        member.middleName || '',
        member.lastName,
        member.suffix || ''
      ].filter(Boolean).join(' '),
      displayName: sendAsDisplayName
    },
    relations: [{
      type: 'manager',
      value: member.managerEmail || ''
    }],
    customSchemas: {
      CAPWATCH: {
        Grade: member.rank || ''
      }
    }
  };

  // Try updating existing user
  try {
    logWorkspaceUserUpdateDiff_(primaryEmail, member, updates);
    user = executeWithRetry(() =>
      AdminDirectory.Users.update(updates, primaryEmail)
    );
    Logger.info('User updated', {
      email: primaryEmail,
      capsn: member.capsn
    });

  } catch (e) {
    if (String(e.message).includes("Resource Not Found")) {
      // Possible archived user — attempt to fetch
      let archivedCheck = null;
      try {
        archivedCheck = AdminDirectory.Users.get(primaryEmail, { projection: "full" });
      } catch (ignored) {}

      if (archivedCheck && archivedCheck.archived) {
        try {
          AdminDirectory.Users.update({ archived: false, suspended: false }, primaryEmail);
          user = AdminDirectory.Users.update(updates, primaryEmail);
          Logger.info("Archived user restored and updated", {
            email: primaryEmail,
            capsn: member.capsn
          });
          return;
        } catch (err2) {
          Logger.error("Failed to unarchive/update archived user", {
            email: primaryEmail,
            capsn: member.capsn,
            errorMessage: err2.message
          });
        }
      }

      // treat as non-existent user
      user = null;
    } else {
      Logger.error('Unable to update user', {
        email: primaryEmail,
        capsn: member.capsn,
        name: member.firstName + ' ' + member.lastName,
        orgid: member.orgid,
        charter: member.charter,
        orgPath: member.orgPath,
        errorMessage: e.message,
        errorCode: e.details?.code
      });
    }
  }

  // Create new user if update failed
  if (!user) {
    // Generate a random temp password. Must not be derivable from public
    // member data (WING, CAPID, provisioning date) — see generateTempPassword_.
    const generatedPassword = generateTempPassword_();
    user = {
      employeeId: String(member.capsn),
      externalIds: [{ type: 'organization', value: String(member.capsn) }],
      organizations: [{
        title: member.dutyPositions && member.dutyPositions.length > 0
          ? member.dutyPositions
              .map(d => d.assistant ? d.id + ' (A)' : d.id)
              .join(', ')
          : 'No Duty Assignment',
        // Squadron / Unit Name (Organization.name)
        department: member.orgName || member.charter || '',
        // Charter (PCR-HI-077 etc.)
        costCenter: member.charter || '',
        description: member.type || '',
        type: 'work',
        primary: true
      }],
      primaryEmail: primaryEmail,
      name: {
        givenName: member.firstName,
        familyName: member.lastName,
        middleName: member.middleName || '',
        fullName: [
          member.firstName,
          member.middleName || '',
          member.lastName,
          member.suffix || ''
        ].filter(Boolean).join(' '),
        displayName: sendAsDisplayName
      },
      suspended: CONFIG.EXCLUDED_ORG_IDS.includes(String(member.orgid)),
      changePasswordAtNextLogin: true,
      password: generatedPassword,
      orgUnitPath: member.orgPath,
      recoveryEmail: member.secondaryEmail || member.parentEmail || '',
      recoveryPhone: member.phone || '',
      phones: member.phone ? [{ type: 'mobile', value: member.phone }] : [],
      emails: member.otherEmail ? [{ type: 'other', address: member.otherEmail }] : [],
      relations: [{
        type: 'manager',
        value: member.managerEmail || ''
      }],
      customSchemas: {
        CAPWATCH: {
          Grade: member.rank || ''
        }
      }
    };

    try {
      let newUser;
      try {
        newUser = executeWithRetry(() =>
          AdminDirectory.Users.insert(user)
        );
      } catch (insertErr) {
        // If custom schema fails (e.g. Invalid Input), try without it
        if (insertErr.message && insertErr.message.includes("Invalid Input")) {
          Logger.warn("Insert failed with custom schemas, retrying without...", {
            email: primaryEmail,
            error: insertErr.message
          });
          delete user.customSchemas;
          newUser = executeWithRetry(() =>
            AdminDirectory.Users.insert(user)
          );
          Logger.warn("User created WITHOUT custom schemas (CAPWATCH/Grade skipped)", {
            email: primaryEmail
          });
        } else {
          throw insertErr;
        }
      }

      Logger.info('New user created', {
        email: primaryEmail,
        capsn: member.capsn,
        name: member.firstName + ' ' + member.lastName
      });
    } catch (e) {
      Logger.error('Failed to create new user', {
        email: primaryEmail,
        capsn: member.capsn,
        name: member.firstName + ' ' + member.lastName,
        orgid: member.orgid,
        charter: member.charter,
        orgPath: member.orgPath,
        errorMessage: e.message,
        errorCode: e.details?.code
      });
      return;
    }

    // Account exists past this point — post-provisioning steps are independent
    // of each other and of account creation, so a failure here must not be
    // mislabeled as (or mask) a failed user creation.

    // Queue for delayed Gmail setup
    try {
      queueForDelayedGmailSetup(primaryEmail, sendAsDisplayName, member);
    } catch (e) {
      Logger.error('User created, but failed to queue delayed Gmail setup', {
        email: primaryEmail,
        capsn: member.capsn,
        errorMessage: e.message,
        errorCode: e.details?.code
      });
    }

    // Send welcome email
    try {
      sendWelcomeEmail(member, primaryEmail, generatedPassword);
    } catch (e) {
      Logger.error('User created, but failed to send welcome email', {
        email: primaryEmail,
        capsn: member.capsn,
        errorMessage: e.message,
        errorCode: e.details?.code
      });
    }
  }
}

/**
 * Gets all active members from CAPWATCH data who are also an eligible member
 * type for a Workspace seat (CONFIG.MEMBER_TYPES.ACTIVE). Members with an
 * ACTIVE CAPWATCH status but an ineligible type (e.g. PATRON) are excluded so
 * that suspendExpiredMembers()/reactivateRenewedMembers() do not treat them
 * as eligible for an account and re-suspend/reactivate them accordingly.
 *
 * @returns {Object} Active, eligible members indexed by CAPID with join date values
 */
function getActiveMembers() {
  let activeMembers = {};
  let memberData = parseFile('Member');
  const eligibleTypes = CONFIG.MEMBER_TYPES.ACTIVE;

  for (let i = 0; i < memberData.length; i++) {
    if (memberData[i][24] === 'ACTIVE' && eligibleTypes.indexOf(memberData[i][21]) > -1) {
      activeMembers[memberData[i][0]] = memberData[i][16];
    }
  }

  Logger.info('Active members retrieved', {
    count: Object.keys(activeMembers).length
  });
  return activeMembers;
}

/**
 * Suspends a Google Workspace user account
 *
 * @param {string} email - User's email address
 * @returns {boolean} True if suspension successful, false otherwise
 */
function suspendMember(email) {
  try {
    executeWithRetry(() =>
      AdminDirectory.Users.update({suspended: true}, email)
    );
    Logger.info('Member suspended', { email: email });
    return true;
  } catch (e) {
    Logger.error('Error suspending member', {
      email: email,
      errorMessage: e.message,
      errorCode: e.details?.code
    });
    return false;
  }
}

/**
 * Retrieves all active (non-suspended) users from Google Workspace
 * Filters to non-admin users
 *
 * @returns {Array<Object>} Array of user objects with email, capid, and lastUpdated
 */
function getActiveUsers() {
  let activeUsers = [];
  let nextPageToken = '';

  do {
    let page = AdminDirectory.Users.list({
      customer: "my_customer",
      maxResults: 500,
      query: 'isSuspended=false isAdmin=false',
      fields: 'users(primaryEmail,externalIds),nextPageToken',
      pageToken: nextPageToken
    });

    nextPageToken = page.nextPageToken;

    if (page.users) {
      for (let i = 0; i < page.users.length; i++) {
        const ids = page.users[i].externalIds || [];
        const capidExt = ids.find(id => id.type === 'organization');

        if (capidExt) {
          activeUsers.push({
            email: page.users[i].primaryEmail,
            capid: capidExt.value,
            lastUpdated: null // no schema anymore
          });
        }
      }
    }
  } while (nextPageToken);

  Logger.info('Active users retrieved from Workspace', {
    count: activeUsers.length
  });
  return activeUsers;
}

/**
 * Returns a Set of CAPIDs whose Level I achievement is marked ACTIVE in
 * MbrAchievements. Used to gate senior account provisioning.
 *
 * @returns {Set<string>}
 */
function loadLevel1CompletedCapids() {
  const completed = new Set();
  const rows = parseFile('MbrAchievements');
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][1]) === '96' && rows[i][2] === 'ACTIVE') {
      completed.add(String(rows[i][0]));
    }
  }
  Logger.info('Level I completions loaded', { count: completed.size });
  return completed;
}

/**
 * Main function to update all member accounts in Google Workspace
 *
 * Process:
 * 1. Retrieves current CAPWATCH member data
 * 2. Compares with previously saved data
 * 3. Updates only changed members
 * 4. Saves current data for future comparison
 * 5. Logs progress every 100 members
 *
 * @returns {void}
 */
function updateAllMembers() {
  clearCache(); // Clear cache for fresh data
  const start = new Date();

  Logger.info('Starting member update process');

  let members = getMembers();
  let currentMembers = getCurrentMemberData();

  const activeUsers = getActiveUsers();
  workspaceEmailByCapid = {};
  activeUsers.forEach(u => {
    workspaceEmailByCapid[String(u.capid)] = u.email;
  });
  assignManagerEmails(members);

  const totalMembers = Object.keys(members).length;

  let processed = 0;

  // Build reduced set of members whose data changed
  let toProcess = {};
  let skipped = 0;
  let missingWorkspaceUsers = 0;

  for (const capsn in members) {
    const isMissingWorkspaceUser = !workspaceEmailByCapid[String(capsn)];
    if (memberUpdated(members[capsn], currentMembers[capsn]) || isMissingWorkspaceUser) {
      toProcess[capsn] = members[capsn];
      if (isMissingWorkspaceUser) missingWorkspaceUsers++;
    } else {
      skipped++;
    }
  }

  // Gate new senior accounts on Level I completion
  if (CONFIG.REQUIRE_LEVEL_I_FOR_SENIORS) {
    const level1Capids = loadLevel1CompletedCapids();
    let level1Skipped = 0;
    for (const capsn in toProcess) {
      const member = toProcess[capsn];
      const isSeniorType = ['SENIOR', 'FIFTY YEAR', 'INDEFINITE', 'CADET SPONSOR'].indexOf(member.type) > -1;
      const isNewAccount = !workspaceEmailByCapid[String(capsn)];
      if (isSeniorType && isNewAccount && !level1Capids.has(String(capsn))) {
        Logger.info('Senior skipped — Level I not complete', {
          capsn: capsn,
          name: member.firstName + ' ' + member.lastName,
          charter: member.charter
        });
        delete toProcess[capsn];
        level1Skipped++;
      }
    }
    if (level1Skipped > 0) {
      Logger.info('New senior accounts withheld pending Level I', { count: level1Skipped });
    }
  }

  Logger.info("Batch update starting", {
    changedMembers: Object.keys(toProcess).length,
    missingWorkspaceUsers: missingWorkspaceUsers,
    skipped: skipped,
    total: totalMembers
  });

  // Log details about what changed
  let changeBreakdown = {
    rankChanged: 0,
    charterChanged: 0,
    dutyChanged: 0,
    statusChanged: 0,
    emailChanged: 0,
    secondaryEmailChanged: 0,
    otherEmailChanged: 0,
    parentEmailChanged: 0,
    phoneChanged: 0,
    managerEmailChanged: 0
  };

  for (const capsn in toProcess) {
    const prev = currentMembers[capsn];
    const curr = members[capsn];
    if (prev) {
      if (curr.rank !== prev.rank) changeBreakdown.rankChanged++;
      if (curr.charter !== prev.charter) changeBreakdown.charterChanged++;
      if (JSON.stringify(curr.dutyPositions) !== JSON.stringify(prev.dutyPositions)) changeBreakdown.dutyChanged++;
      if (curr.status !== prev.status) changeBreakdown.statusChanged++;
      if (curr.email !== prev.email) changeBreakdown.emailChanged++;
      if (curr.secondaryEmail !== prev.secondaryEmail) changeBreakdown.secondaryEmailChanged++;
      if (curr.otherEmail !== prev.otherEmail) changeBreakdown.otherEmailChanged++;
      if (curr.parentEmail !== prev.parentEmail) changeBreakdown.parentEmailChanged++;
      if (curr.phone !== prev.phone) changeBreakdown.phoneChanged++;
      if (curr.managerEmail !== prev.managerEmail) changeBreakdown.managerEmailChanged++;
    }
  }

  Logger.info("Change breakdown", changeBreakdown);

  // 🔥 MAIN CHANGE — batch processing instead of inline updates
  batchUpdateMembers(toProcess);

  // 👉 Alias updates — from sheet tab Aliases
  //addAliasesFromSheet();

  Logger.info("Saving current member data snapshot");
  saveCurrentMemberData(members);

  // Reactivate any members who renewed
  Logger.info('Checking for renewed members to reactivate');
  const reactivationStart = new Date();
  let totalReactivated = 0;

  try {
    // Get inactive users before calling reactivateRenewedMembers
    const activeMembers = getActiveMembers();
    const inactiveUsers = getInactiveUsers();
    let reactivated = 0;
    let unarchived = 0;

    for (let i = 0; i < inactiveUsers.length; i++) {
      const user = inactiveUsers[i];

      if (user.capid && (user.capid in activeMembers)) {
        const wasArchived = user.archived;
        const success = reactivateMember(user.email, wasArchived);

        if (success) {
          if (wasArchived) {
            unarchived++;
          } else {
            reactivated++;
          }
        }
      }
    }

    totalReactivated = reactivated + unarchived;

    Logger.info('Renewed member reactivation completed', {
      duration: new Date() - reactivationStart + 'ms',
      reactivated: reactivated,
      unarchived: unarchived,
      total: totalReactivated
    });
  } catch (err) {
    Logger.error('Reactivation check failed', {
      errorMessage: err.message
    });
  }

  Logger.info('Member update completed', {
    duration: new Date() - start + 'ms',
    totalProcessed: processed,
    updated: Object.keys(toProcess).length,
    skipped: skipped,
    reactivated: totalReactivated
  });
}

/**
 * Suspends Google Workspace accounts for members who are no longer active in CAPWATCH
 *
 * Process:
 * 1. Gets active members from CAPWATCH
 * 2. Gets active users from Google Workspace
 * 3. Identifies users not in CAPWATCH
 * 4. Suspends after grace period expires
 *
 * @returns {void}
 */
function suspendExpiredMembers() {
  const start = new Date();
  Logger.info('Starting expired member suspension process');

  let activeMembers = getActiveMembers();
  let users = getActiveUsers();
  // Manual members are not present in CAPWATCH Member.txt, so they will not appear in getActiveMembers().
  // Protect them from suspension by building a CAPID allow-list from the Manual Members sheet.
  const manualCapids = {};
  try {
    const ss = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Manual Members') || ss.getSheetByName('ManualMembers');
    if (sheet) {
      const rows = sheet.getDataRange().getValues();
      if (rows && rows.length > 1) {
        const header = rows[0].map(h => String(h || '').trim());
        const capidIdx = header.indexOf('CAPID');
        if (capidIdx > -1) {
          for (let r = 1; r < rows.length; r++) {
            const capid = String(rows[r][capidIdx] || '').trim();
            if (capid) manualCapids[capid] = 1;
          }
        }
      }
    }
  } catch (e) {
    Logger.warn('Unable to read Manual Members sheet for suspension bypass', {
      errorMessage: e && e.message ? e.message : String(e)
    });
  }
  let suspended = 0;
  let pending = 0;
  const suspensionTime = new Date().getTime() - (CONFIG.SUSPENSION_GRACE_DAYS * 86400000);

  for(let i = 0; i < users.length; i++) {
    if (users[i].capid && !(users[i].capid in activeMembers)) {
      // Do not suspend manual members (PCR/NHQ/etc. added via Manual Members tab)
      if (manualCapids[users[i].capid]) {
        continue;
      }
      if (!users[i].lastUpdated || suspensionTime > new Date(users[i].lastUpdated).getTime()) {
        let success = suspendMember(users[i].email);
        if (success) {
          suspended++;
        }
      } else {
        Logger.info('Member expired - pending suspension', {
          email: users[i].email,
          capid: users[i].capid,
          lastUpdated: users[i].lastUpdated,
          graceDaysRemaining: Math.ceil((new Date(users[i].lastUpdated).getTime() + (CONFIG.SUSPENSION_GRACE_DAYS * 86400000) - new Date().getTime()) / 86400000)
        });
        pending++;
      }
    }
  }

  Logger.info('Expired member suspension completed', {
    duration: new Date() - start + 'ms',
    suspended: suspended,
    pending: pending,
    graceDays: CONFIG.SUSPENSION_GRACE_DAYS
  });
}

/**
 * Reactivates Google Workspace accounts for members who renewed after being suspended or archived
 *
 * Process:
 * 1. Gets active members from CAPWATCH
 * 2. Gets suspended/archived users from Google Workspace
 * 3. Identifies users who are now active in CAPWATCH
 * 4. Unsuspends and/or unarchives them
 *
 * This handles both:
 * - Members who renewed within 1 year (suspended only)
 * - Members who renewed after 1+ year (archived)
 *
 * @returns {void}
 */
function reactivateRenewedMembers() {
  const start = new Date();
  Logger.info('Starting renewed member reactivation process');

  const activeMembers = getActiveMembers();
  const inactiveUsers = getInactiveUsers();
  let reactivated = 0;
  let unarchived = 0;
  let failed = 0;

  for (let i = 0; i < inactiveUsers.length; i++) {
    const user = inactiveUsers[i];

    // Check if user is now active in CAPWATCH
    if (user.capid && (user.capid in activeMembers)) {
      const wasArchived = user.archived;
      const success = reactivateMember(user.email, wasArchived);

      if (success) {
        if (wasArchived) {
          unarchived++;
          Logger.info('Archived member reactivated', {
            email: user.email,
            capid: user.capid,
            wasArchived: true
          });
        } else {
          reactivated++;
          Logger.info('Suspended member reactivated', {
            email: user.email,
            capid: user.capid
          });
        }
      } else {
        failed++;
      }
    }
  }

  Logger.info('Renewed member reactivation completed', {
    duration: new Date() - start + 'ms',
    reactivated: reactivated,
    unarchived: unarchived,
    failed: failed,
    total: reactivated + unarchived
  });
}

/**
 * Reactivates a Google Workspace user account
 * Handles both suspended and archived users
 *
 * @param {string} email - User's email address
 * @param {boolean} wasArchived - Whether the user was archived (vs just suspended)
 * @returns {boolean} True if reactivation successful, false otherwise
 */
function reactivateMember(email, wasArchived = false) {
  try {
    const updateObject = {
      suspended: false,
      archived: false
    };

    executeWithRetry(() =>
      AdminDirectory.Users.update(updateObject, email)
    );

    const status = wasArchived ? 'Member unarchived and unsuspended' : 'Member unsuspended';
    Logger.info(status, { email: email });
    return true;
  } catch (e) {
    Logger.error('Error reactivating member', {
      email: email,
      wasArchived: wasArchived,
      errorMessage: e.message,
      errorCode: e.details?.code
    });
    return false;
  }
}

/**
 * Retrieves all inactive (suspended or archived) users from Google Workspace
 * Filters to non-admin users with CAPID
 *
 * @returns {Array<Object>} Array of user objects with email, capid, archived status
 */
function getInactiveUsers() {
  let inactiveUsers = [];
  let nextPageToken = '';

  do {
    let page = AdminDirectory.Users.list({
      customer: "my_customer",
      maxResults: 500,
      query: 'isSuspended=true isAdmin=false',
      fields: 'users(primaryEmail,suspended,archived,externalIds),nextPageToken',
      pageToken: nextPageToken
    });

    nextPageToken = page.nextPageToken;

    if (page.users) {
      for (let i = 0; i < page.users.length; i++) {
        const ids = page.users[i].externalIds || [];
        const capidExt = ids.find(id => id.type === 'organization');

        if (capidExt) {
          inactiveUsers.push({
            email: page.users[i].primaryEmail,
            capid: capidExt.value,
            archived: page.users[i].archived || false
          });
        }
      }
    }
  } while (nextPageToken);

  Logger.info('Inactive users retrieved from Workspace', {
    count: inactiveUsers.length
  });
  return inactiveUsers;
}

/**
 * Debug helper: lists the inactive users that reactivation logic can see.
 *
 * This is useful for confirming which suspended users have the required
 * CAPID mapping and are therefore eligible for automatic reactivation.
 *
 * @returns {Array<Object>} Visible inactive users
 */
function debugListInactiveUsersVisibleToReactivation() {
  const users = getInactiveUsers()
    .slice()
    .sort((a, b) => {
      const emailA = String(a.email || '').toLowerCase();
      const emailB = String(b.email || '').toLowerCase();
      return emailA.localeCompare(emailB);
    });

  Logger.info('Inactive users visible to reactivation', {
    count: users.length,
    users: users
  });

  return users;
}

/**
 * Debug helper: finds active Workspace users who appear to be CAP accounts
 * but are invisible to suspendExpiredMembers() because they do not have the
 * externalIds.organization mapping used by getActiveUsers().
 *
 * This is the main class of "users like Abigail":
 * - active in Workspace
 * - have an Employee ID / CAPWATCH footprint
 * - not active in current CAPWATCH Member.txt
 * - missing externalIds.organization, so suspension logic never sees them
 *
 * @returns {Array<Object>} Candidate users that suspension will skip
 */
function debugFindExpiredUsersMissingExternalIds() {
  const activeMembers = getActiveMembers();
  const candidates = [];
  let nextPageToken = '';

  do {
    const page = AdminDirectory.Users.list({
      customer: "my_customer",
      maxResults: 500,
      query: 'isSuspended=false isAdmin=false',
      projection: 'full',
      pageToken: nextPageToken
    });

    nextPageToken = page.nextPageToken;

    if (!page.users) continue;

    for (let i = 0; i < page.users.length; i++) {
      const user = page.users[i];
      const employeeId = String(user.employeeId || '').trim();
      const ids = user.externalIds || [];
      const capidExt = ids.find(id => id.type === 'organization');
      const externalCapid = capidExt ? String(capidExt.value || '').trim() : '';
      const schemaCapwatch = user.customSchemas && user.customSchemas.CAPWATCH
        ? user.customSchemas.CAPWATCH
        : null;

      // Ignore users that already have the mapping suspension relies on.
      if (externalCapid) continue;

      // We only care about accounts that still look like managed CAP accounts.
      const fallbackCapid = employeeId ||
        (schemaCapwatch && String(schemaCapwatch.CAPID || '').trim()) ||
        '';
      if (!fallbackCapid) continue;

      // These are the ones that are expired but invisible to suspension.
      if (fallbackCapid in activeMembers) continue;

      candidates.push({
        email: user.primaryEmail,
        employeeId: employeeId || '',
        fallbackCapid: fallbackCapid,
        displayName: user.name && user.name.fullName ? user.name.fullName : '',
        hasCapwatchSchema: !!schemaCapwatch,
        visibleToSuspension: false
      });
    }
  } while (nextPageToken);

  candidates.sort((a, b) => {
    const emailA = String(a.email || '').toLowerCase();
    const emailB = String(b.email || '').toLowerCase();
    return emailA.localeCompare(emailB);
  });

  Logger.info('Expired active users missing externalIds.organization', {
    count: candidates.length,
    users: candidates
  });

  return candidates;
}

/**
 * Debug helper: lists active Workspace users that do not have
 * externalIds.type === 'organization'.
 *
 * This is useful for finding legacy or malformed accounts that the
 * current suspension/reactivation logic cannot map back to a CAPID.
 *
 * @returns {Array<Object>} Active users missing organization externalIds
 */
function debugListUsersMissingOrganizationExternalId() {
  const candidates = [];
  let nextPageToken = '';

  do {
    const page = AdminDirectory.Users.list({
      customer: "my_customer",
      maxResults: 500,
      query: 'isSuspended=false isAdmin=false',
      projection: 'full',
      pageToken: nextPageToken
    });

    nextPageToken = page.nextPageToken;

    if (!page.users) continue;

    for (let i = 0; i < page.users.length; i++) {
      const user = page.users[i];
      const ids = user.externalIds || [];
      const capidExt = ids.find(id => id.type === 'organization');
      if (capidExt) continue;

      candidates.push({
        email: user.primaryEmail,
        employeeId: String(user.employeeId || '').trim(),
        displayName: user.name && user.name.fullName ? user.name.fullName : '',
        orgUnitPath: String(user.orgUnitPath || ''),
        hasExternalIds: ids.length > 0,
        externalIds: ids,
        hasCapwatchSchema: !!(user.customSchemas && user.customSchemas.CAPWATCH)
      });
    }
  } while (nextPageToken);

  candidates.sort((a, b) => {
    const emailA = String(a.email || '').toLowerCase();
    const emailB = String(b.email || '').toLowerCase();
    return emailA.localeCompare(emailB);
  });

  Logger.info('Active users missing externalIds.organization', {
    count: candidates.length,
    users: candidates
  });

  return candidates;
}

/**
 * Debug helper: reconciles active CAPWATCH members with active Workspace CAPID mappings.
 *
 * This answers two practical questions:
 * 1) Which CAPWATCH-active CAPIDs are not currently visible in active Workspace users?
 * 2) Which active Workspace CAPIDs are not active in CAPWATCH?
 *
 * @returns {Object} Reconciliation payload with both mismatch lists
 */
function debugReconcileActiveCapidMappings() {
  const activeMembers = getActiveMembers();   // CAPID -> joinDate
  const activeUsers = getActiveUsers();       // [{ email, capid, ... }]

  const workspaceByCapid = {};
  for (let i = 0; i < activeUsers.length; i++) {
    const capid = String(activeUsers[i].capid || '').trim();
    if (!capid) continue;
    workspaceByCapid[capid] = activeUsers[i].email;
  }

  const capwatchMissingFromWorkspace = [];
  const workspaceMissingFromCapwatch = [];

  // CAPWATCH active but not represented in active Workspace mapping
  for (const capid in activeMembers) {
    if (!workspaceByCapid[capid]) {
      capwatchMissingFromWorkspace.push({
        capid: capid,
        joined: activeMembers[capid] || ''
      });
    }
  }

  // Active Workspace CAPIDs not active in CAPWATCH
  for (const capid in workspaceByCapid) {
    if (!(capid in activeMembers)) {
      workspaceMissingFromCapwatch.push({
        capid: capid,
        email: workspaceByCapid[capid]
      });
    }
  }

  capwatchMissingFromWorkspace.sort((a, b) =>
    String(a.capid).localeCompare(String(b.capid))
  );
  workspaceMissingFromCapwatch.sort((a, b) =>
    String(a.capid).localeCompare(String(b.capid))
  );

  const result = {
    capwatchActiveCount: Object.keys(activeMembers).length,
    workspaceActiveMappedCount: Object.keys(workspaceByCapid).length,
    capwatchMissingFromWorkspaceCount: capwatchMissingFromWorkspace.length,
    workspaceMissingFromCapwatchCount: workspaceMissingFromCapwatch.length,
    capwatchMissingFromWorkspace: capwatchMissingFromWorkspace,
    workspaceMissingFromCapwatch: workspaceMissingFromCapwatch
  };

  Logger.info('Active CAPID reconciliation', result);
  return result;
}

/**
 * Debug helper: lists Workspace custom schema API names and field names.
 *
 * Use this to verify the exact schemaName and fieldName required by
 * customSchemas payloads, which can differ from Admin Console display labels.
 *
 * @returns {Array<Object>} Custom schema summary
 */
function debugListCustomSchemas() {
  const response = AdminDirectory.Schemas.list('my_customer');
  const schemas = response && response.schemas ? response.schemas : [];

  const result = schemas.map(schema => ({
    schemaName: schema.schemaName || '',
    displayName: schema.displayName || '',
    fields: (schema.fields || []).map(field => ({
      fieldName: field.fieldName || '',
      displayName: field.displayName || '',
      fieldType: field.fieldType || '',
      multiValued: !!field.multiValued,
      readAccessType: field.readAccessType || ''
    }))
  }));

  Logger.info('Workspace custom schemas', {
    count: result.length,
    schemas: result
  });
  return result;
}

/**
 * Adds an email alias to a user account with retry logic for conflicts
 *
 * Tries firstname.lastname first, then firstname.lastname.1, firstname.lastname.2, etc.
 * up to 5 attempts if alias already exists
 *
 * @param {Object} user - User object with name properties
 * @param {Object} user.name - Name object
 * @param {string} user.name.givenName - First name
 * @param {string} user.name.familyName - Last name
 * @param {string} user.primaryEmail - User's primary email
 * @returns {Object|null} Alias object if successful, null if failed
 */

function addAlias(user) {
  const maxRetry = 5;
  let aliasEmail;
  let alias;

  // Try setting default alias first
  try {
    aliasEmail = user.name.givenName.replace(/\s/g, '') + '.' +
                 user.name.familyName.replace(/\s/g, '') + CONFIG.EMAIL_DOMAIN;
    alias = AdminDirectory.Users.Aliases.insert({alias: aliasEmail}, user.primaryEmail);
    if (alias) {
      Logger.info('Alias added', {
        user: user.primaryEmail,
        alias: aliasEmail
      });
      return alias;
    }
  } catch(err) {
    if (err.details?.code !== 409) {
      Logger.error('Failed to add alias', {
        user: user.primaryEmail,
        attemptedAlias: aliasEmail,
        errorMessage: err.message,
        errorCode: err.details?.code
      });
      return null;
    }
    // 409 = Conflict, try with number suffix
  }

  // Make 5 attempts with incrementing numbers
  for (let index = 1; index <= maxRetry; index++) {
    try {
      aliasEmail = user.name.givenName.replace(/\s/g, '') + '.' +
                   user.name.familyName.replace(/\s/g, '') + '.' + index + CONFIG.EMAIL_DOMAIN;
      alias = AdminDirectory.Users.Aliases.insert({alias: aliasEmail}, user.primaryEmail);
      if (alias) {
        Logger.info('Alias added with suffix', {
          user: user.primaryEmail,
          alias: aliasEmail,
          attempt: index
        });
        return alias;
      }
    } catch (err) {
      if (err.details?.code !== 409) {
        Logger.error('Failed to add alias with suffix', {
          user: user.primaryEmail,
          attemptedAlias: aliasEmail,
          attempt: index,
          errorMessage: err.message,
          errorCode: err.details?.code
        });
        return null;
      }
    }
  }

  Logger.error('All alias attempts failed', {
    user: user.primaryEmail,
    attempts: maxRetry + 1
  });
  return null;
}

/**
 * Generates an OAuth2 token for a service account to impersonate a user.
 * This is required for APIs like Gmail settings that have strict delegation rules.
 * @param {string} userToImpersonate The email address of the user to impersonate.
 * @param {string} scope The OAuth2 scope(s) required for the API call.
 * @returns {string} The access token.
 */
function getImpersonatedToken_(userToImpersonate, scope) {
  const props = PropertiesService.getScriptProperties();

  const SERVICE_ACCOUNT_EMAIL = props.getProperty('SA_IMPERSONATION_EMAIL');
  let PRIVATE_KEY = props.getProperty('SA_PRIVATE_KEY');

  if (!SERVICE_ACCOUNT_EMAIL || !PRIVATE_KEY) {
    throw new Error('Missing service account credentials in Script Properties (SA_IMPERSONATION_EMAIL / SA_PRIVATE_KEY).');
  }

  // If the key was stored with literal "\n", convert to real newlines
  PRIVATE_KEY = PRIVATE_KEY.replace(/\\n/g, '\n');

  const now = Math.floor(Date.now() / 1000);

  const claimSet = {
    iss: SERVICE_ACCOUNT_EMAIL,
    sub: userToImpersonate,
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now,
    scope: scope
  };

  const header = { alg: 'RS256', typ: 'JWT' };
  const toSign =
    `${Utilities.base64EncodeWebSafe(JSON.stringify(header))}.` +
    `${Utilities.base64EncodeWebSafe(JSON.stringify(claimSet))}`;

  const signature = Utilities.computeRsaSha256Signature(toSign, PRIVATE_KEY);
  const jwt = `${toSign}.${Utilities.base64EncodeWebSafe(signature)}`;

  const response = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt
    },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error(`Token exchange failed (${code}): ${body}`);
  }

  const token = JSON.parse(body);
  return token.access_token;
}

/**
 * Updates Gmail "Send As" display name for a user via impersonation.
 *
 * @param {string} primaryEmail - User's primary email
 * @param {string} displayName - Gmail Send-As display name
 */
function updateGmailSendAsDisplayName(primaryEmail, displayName) {
  const scope =
    'https://www.googleapis.com/auth/gmail.settings.basic ' +
    'https://www.googleapis.com/auth/gmail.settings.sharing';
  let accessToken;

  // Get impersonation token
  try {
    accessToken = getImpersonatedToken_(primaryEmail, scope);
  } catch (e) {
    Logger.error('Impersonation token failed', {
      user: primaryEmail,
      errorMessage: e.message
    });
    return;
  }

  //
  // ---------------------------------------------------------
  //  STEP 1 — OPTIONAL ALIAS CREATION FROM "Aliases" SHEET
  // ---------------------------------------------------------
  //
  try {
    const sheet = SpreadsheetApp
      .openById(CONFIG.AUTOMATION_SPREADSHEET_ID)
      .getSheetByName('Aliases');

    if (sheet) {
      const rows = sheet.getDataRange().getValues();
      rows.shift(); // remove header row

      const row = rows.find(r =>
        String(r[0]).trim().toLowerCase() === primaryEmail.toLowerCase()
      );

      if (row) {
        const aliasEmail = String(row[1] || '').trim();
        const aliasDisplay = String(row[2] || displayName).trim();
        const aliasSignature = row[3] || '';

        if (aliasEmail) {
          const aliasBody = {
            sendAsEmail: aliasEmail,
            displayName: aliasDisplay,
            signature: aliasSignature,
            treatAsAlias: true
          };

          const aliasUrl =
            'https://gmail.googleapis.com/gmail/v1/users/me/settings/sendAs';

          const aliasResp = UrlFetchApp.fetch(aliasUrl, {
            method: 'post',
            contentType: 'application/json',
            headers: { Authorization: 'Bearer ' + accessToken },
            payload: JSON.stringify(aliasBody),
            muteHttpExceptions: true
          });

          const code = aliasResp.getResponseCode();
          if (code < 200 || code >= 300) {
            // Only log non-409 failures
            if (code !== 409) {
              Logger.warn('Alias add failed', {
                primary: primaryEmail,
                alias: aliasEmail,
                code: code,
                response: aliasResp.getContentText()
              });
            }
          }
        }
      }
    }
  } catch (aliasErr) {
    Logger.error('Alias-add attempt failed', {
      primary: primaryEmail,
      error: aliasErr.message
    });
  }

  //
  // ---------------------------------------------------------
  //  STEP 2 — PATCH PRIMARY SEND-AS IDENTITY
  // ---------------------------------------------------------
  //
  try {
    const apiUrl =
      `https://gmail.googleapis.com/gmail/v1/users/${encodeURIComponent(primaryEmail)}/settings/sendAs/${encodeURIComponent(primaryEmail)}`;

    const sendAsBody = {
      displayName: displayName
    };

    const response = UrlFetchApp.fetch(apiUrl, {
      method: 'patch',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + accessToken },
      payload: JSON.stringify(sendAsBody),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();

    // Quiet mode — only log failures
    if (code < 200 || code >= 300) {
      Logger.warn('Primary Send-As display name update failed', {
        email: primaryEmail,
        code: code,
        response: response.getContentText()
      });
    }

  } catch (e) {
    Logger.error('Primary Send-As patch threw exception', {
      email: primaryEmail,
      error: e.message
    });
  }
}

/**
 * Pushes signatures for all members to all aliases.
 * This is the production version replacing pushSignatureToAllAliases454201().
 */
/**
 * Pushes signatures for all Workspace users to all aliases.
 * Reads actual Workspace users, matches via CAPID, generates signature,
 * and updates Gmail signatures for each alias.
 */
function pushAllSignatures() {
  Logger.info('Starting pushAllSignatures');

  // 1. Fetch CAPWATCH members
  const members = getMembers();
  const memberIndex = {};

  // Build quick lookup: CAPID → member object
  for (const capid in members) {
    memberIndex[String(capid)] = members[capid];
  }

  // 2. Fetch all Workspace users
  let pageToken = null;
  const workspaceUsers = [];

  do {
    const page = AdminDirectory.Users.list({
      customer: "my_customer",
      maxResults: 200,
      projection: "full",
      pageToken: pageToken
    });

    if (page.users) {
      workspaceUsers.push(...page.users);
    }

    pageToken = page.nextPageToken;

  } while (pageToken);

  // 3. Loop through Workspace users
  workspaceUsers.forEach(user => {
      const capid = user.externalIds?.[0]?.value;
      const member = memberIndex[capid];

    if (!member) {
      Logger.warn('Workspace user has no matching CAPWATCH record', {
        email: user.primaryEmail
      });
      return;
    }


    // Generate signature
    const signature = generateEmailSignature(member);

    // Push signature to all aliases
    updateSignatureForAllAliases(user.primaryEmail, signature);

    Utilities.sleep(200);
  });

  Logger.info('pushAllSignatures completed for all users');
}

// ==========================================================================
// SHARED SIGNATURE HELPERS
// ==========================================================================

function toTitleCase(str) {
  return (str || '')
    .toLowerCase()
    .replace(/\b\w+/g, t => t[0].toUpperCase() + t.substring(1));
}

function formatPhone(phone) {
  if (!phone) return '';
  const digits = phone.replace(/\D/g, '').slice(-10);
  return digits ? digits.replace(/(\d{3})(\d{3})(\d{4})/, '$1.$2.$3') : '';
}

function getPublicRank(rank) {
  const MAP = {
    "SSgt": "Staff Sgt.",
    "TSgt": "Tech. Sgt.",
    "MSgt": "Master Sgt.",
    "SMSgt": "Senior Master Sgt.",
    "CMSgt": "Chief Master Sgt.",
    "FO": "Flight Officer",
    "TFO": "Tech. Flight Officer",
    "SFO": "Senior Flight Officer",
    "2d Lt": "2nd Lt.",
    "1st Lt": "1st Lt.",
    "Capt": "Capt.",
    "Maj": "Maj.",
    "Lt Col": "Lt. Col.",
    "Col": "Col.",
    "Brig Gen": "Brig. Gen.",
    "Maj Gen": "Maj. Gen."
  };
  return MAP[rank] || rank || '';
}

function getWingCode() {
  return CONFIG.WING.toLowerCase() + "wg";
}

function getDutyBlock(member) {
  if (!member.dutyPositions || member.dutyPositions.length === 0) {
    return 'Member';
  }
  return member.dutyPositions
    .filter(dp => !dp.assistant)
    .map(dp => `${toTitleCase(member.orgName)} ${dp.id}`)
    .join('<br>');
}

/**
 * Generates an HTML email signature for a member using the standard CAP template.
 * This produces a Gmail-ready HTML block that can be pasted into the user's signature settings
 * or pushed programmatically.
 *
 * @param {Object} member - A fully constructed member object from getMembers()
 * @returns {string} HTML signature block
 */
function generateEmailSignature(member) {
  const rank = getPublicRank(member.rank);
  const first = member.firstName || '';
  const last = member.lastName || '';
  const duty = getDutyBlock(member);
  const wingCode = getWingCode(member);
  const phoneDigits = member.phone ? member.phone.replace(/\D/g, '').slice(-10) : '';
  const phoneFormatted = formatPhone(member.phone);

  return `
<!DOCTYPE html>
<html>
<body>
<br />

<h1 style="font-size: 12px; line-height: 12px;
           font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
           color: #001871; font-weight: bold; margin: 0 0 5px;">
  ${rank} ${first} ${last}
</h1>

<h2 style="font-size: 12px; line-height: 12px;
           font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
           color: #000000; font-weight: normal; margin: 0 0 5px;">
  ${duty}
</h2>

<h2 style="font-size: 12px; line-height: 12px;
           font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
           color: #000000; margin: 0 0 20px;">
  Civil Air Patrol, U.S. Air Force Auxiliary
</h2>

<p style="font-size: 12px; line-height: 12px;
          font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
          color: #000000; font-weight: normal; margin: 0 0 5px;">
  (M) <a href="tel:+1${phoneDigits}" style="color: #000000; text-decoration: none;">${phoneFormatted}</a>
</p>

<p style="font-size: 12px; line-height: 12px;
          font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
          color: #001871; font-weight: normal; margin: 0 0 5px;">
  <a href="https://www.GoCivilAirPatrol.com"
     style="color: #000000; text-decoration: underline;">
     GoCivilAirPatrol.com
  </a>
</p>

<p style="font-size: 12px; line-height: 12px;
        font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
        color: #001871; font-weight: normal; margin: 0 0 5px;">
 <a href="https://${wingCode}.cap.gov"
   style="color: #000000; text-decoration: underline;">
   ${wingCode}.cap.gov
 </a>
</p>

<a href="https://www.GoCivilAirPatrol.com">
  <img
    style="margin: 15px 0 20px -5px;"
    src="https://cdn-assets-cloud.frontify.com/s3/frontify-cloud-files-us/eyJwYXRoIjoiZnJvbnRpZnlcL2ZpbGVcLzFNOTF6UzJGRXByV0pNUjQ1cnkxLnBuZyJ9:frontify:kAS-l74of5IyRDi0IxD9lG5yNH6i-Zg30PGjGUjU8y0?width=200">
</a>

<p style="font-size: 12px; line-height: 14px;
          font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
          color:#000000; font-weight: normal; font-style: italic;
          white-space: normal; margin: 0 0 5px;">
  Volunteers serving America&apos;s communities, saving lives, and shaping futures.
</p>

</body>
</html>
  `;
}

/**
 * Generates a random temporary password for a new Workspace account.
 *
 * Entropy comes from two type-4 (random) UUIDs — ~244 bits before mapping.
 * The result is NOT derivable from public member data (CAPID, WING, date),
 * and every account gets a distinct password. changePasswordAtNextLogin is
 * still set on the account so this value is only valid until first login.
 *
 * Guarantees at least one lowercase letter, one uppercase letter, one digit,
 * and one special character to satisfy any password-complexity policy.
 *
 * @returns {string} A random 14-character password.
 */
function generateTempPassword_() {
  const hex = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  const letters = 'abcdefghijkmnpqrstuvwxyz'; // no ambiguous l/o
  let body = '';
  for (let i = 0; i < 20; i += 2) {
    const byte = parseInt(hex.substr(i, 2), 16);
    const ch = letters.charAt(byte % letters.length);
    // Upper-case roughly every other character for mixed case.
    body += (i % 4 === 0) ? ch.toUpperCase() : ch;
  }

  // Explicitly guarantee one char from each required class so the result
  // always satisfies a complexity policy regardless of the letters drawn above.
  const specials = '!@#$%^&*';
  const lower = letters.charAt(parseInt(hex.substr(20, 2), 16) % letters.length);
  const upper = letters.charAt(parseInt(hex.substr(22, 2), 16) % letters.length).toUpperCase();
  const digit = String(parseInt(hex.substr(24, 2), 16) % 10);
  const special = specials.charAt(parseInt(hex.substr(26, 2), 16) % specials.length);
  return body + lower + upper + digit + special;
}

// -----------------------------------------
// NEW — Welcome Email Sender
// -----------------------------------------
function sendWelcomeEmail(member, email, tempPassword) {
  const html = HtmlService
    .createTemplateFromFile('recruiting-and-retention/WelcomeEmail')
    .getRawContent();

  const mergedHtml = html
    .replace(/{{WING}}/g, CONFIG.WING)
    .replace(/{{firstName}}/g, member.firstName)
    .replace(/{{lastName}}/g, member.lastName)
    .replace(/{{email}}/g, email)
    .replace(/{{password}}/g, tempPassword)
    .replace(/{{ITSUPPORT_EMAIL}}/g, ITSUPPORT_EMAIL)
    .replace(/{{DOMAIN}}/g, CONFIG.DOMAIN)
    .replace(/{{rank}}/g, member.rank || '')
    .replace(/{{primaryEmail}}/g, member.email || '')
    .replace(/{{secondaryEmail}}/g, member.secondaryEmail || '');

  MailApp.sendEmail({
    to: [member.email, member.secondaryEmail].filter(Boolean).join(','),
    cc: ITSUPPORT_EMAIL,
    subject: `New Workspace Account – ${member.rank ? member.rank + ' ' : ''}${member.firstName} ${member.lastName}`,
    htmlBody: mergedHtml
  });

  Logger.info("Welcome email sent to IT", {
    user: email,
    support: ITSUPPORT_EMAIL
  });
}

function queueForDelayedGmailSetup(email, displayName, member) {
  const script = ScriptApp.newTrigger('runDelayedGmailSetup')
    .timeBased()
    .after(5 * 60 * 1000)  // 5 minutes
    .create();

  PropertiesService.getScriptProperties().setProperty(
    'gmailsetup_' + script.getUniqueId(),
    JSON.stringify({ email: email, displayName: displayName, capsn: member.capsn })
  );
}

function runDelayedGmailSetup(e) {
  const triggers = ScriptApp.getProjectTriggers();
  const props = PropertiesService.getScriptProperties();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runDelayedGmailSetup') {

      const emailKey = 'gmailsetup_' + trigger.getUniqueId();
      const record = props.getProperty(emailKey);

      if (record) {
        const data = JSON.parse(record);

        try {
          updateGmailSendAsDisplayName(data.email, data.displayName);

          const member = { capsn: data.capsn };
          const signature = generateEmailSignature(member);
          updateSignatureForAllAliases(data.email, signature);

          Logger.info("Delayed Gmail setup completed", data);

        } catch (err) {
          Logger.error("Delayed Gmail setup failed", {
            email: data.email,
            error: err.message
          });
        }

        props.deleteProperty(emailKey);
      }

      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Processes members in batches to manage API rate limits
 *
 * @param {Object} members - Members object to process
 * @param {number} batchSize - Number of members per batch (default: 50)
 * @returns {void}
 */
function batchUpdateMembers(members, batchSize = CONFIG.BATCH_SIZE) {
  const memberArray = Object.values(members);
  const totalBatches = Math.ceil(memberArray.length / batchSize);

  Logger.info('Starting batch member update', {
    totalMembers: memberArray.length,
    batchSize: batchSize,
    totalBatches: totalBatches
  });

  let successCount = 0;
  let errorCount = 0;
  const batchStart = new Date();

  for (let i = 0; i < memberArray.length; i += batchSize) {
    const batch = memberArray.slice(i, i + batchSize);
    const batchNumber = Math.floor(i / batchSize) + 1;
    const batchIterationStart = new Date();

    Logger.info('Processing batch', {
      batch: batchNumber,
      totalBatches: totalBatches,
      batchSize: batch.length,
      progress: `${i + batch.length}/${memberArray.length}`,
      percentComplete: Math.round(((i + batch.length) / memberArray.length) * 100) + '%'
    });

    // Process batch
    batch.forEach((member, idx) => {
      try {
        addOrUpdateUser(member);
        successCount++;
      } catch (e) {
        errorCount++;
        Logger.error('Batch member update failed', {
          capsn: member.capsn,
          email: member.firstName + '.' + member.lastName + CONFIG.EMAIL_DOMAIN,
          batchNumber: batchNumber,
          batchIndex: idx,
          errorMessage: e.message
        });
      }
    });

    const batchIterationDuration = new Date() - batchIterationStart;
    Logger.info('Batch completed', {
      batch: batchNumber,
      duration: batchIterationDuration + 'ms',
      successInBatch: batch.length - errorCount,
      totalSuccess: successCount,
      totalErrors: errorCount
    });

    // Add delay between batches to avoid rate limits
    if (i + batchSize < memberArray.length) {
      Utilities.sleep(1000); // 1 second delay
    }
  }

  const totalDuration = new Date() - batchStart;
  Logger.info('Batch update completed', {
    totalMembers: memberArray.length,
    batches: totalBatches,
    successCount: successCount,
    errorCount: errorCount,
    duration: totalDuration + 'ms',
    avgTimePerMember: Math.round(totalDuration / memberArray.length) + 'ms'
  });
}

/**
 * Finds squadrons that are missing organizational unit paths
 * Useful for identifying configuration issues
 *
 * @returns {Array<Object>} Array of squadrons missing orgPath
 */
function findMissingOrgPaths() {
  const squadrons = getSquadrons();
  const missing = [];

  for (const orgid in squadrons) {
    if (!squadrons[orgid].orgPath || squadrons[orgid].orgPath === '') {
      missing.push({
        orgid: orgid,
        name: squadrons[orgid].name,
        charter: squadrons[orgid].charter,
        scope: squadrons[orgid].scope
      });
    }
  }

  Logger.info('Missing orgPath check completed', {
    totalSquadrons: Object.keys(squadrons).length,
    missingOrgPaths: missing.length
  });

  if (missing.length > 0) {
    Logger.warn('Squadrons missing orgPaths', {
      count: missing.length,
      squadrons: missing
    });
  }

  return missing;
}

/**
 * Restores Google Workspace users that were manually deleted but still appear in CAPWATCH.
 *
 * This function:
 * 1. Loads all CAPWATCH members
 * 2. Computes expected primary email (firstname.lastname@wingcap.org)
 * 3. Checks if the Workspace user exists
 * 4. Recreates the user if missing
 */
function restoreDeletedUsers() {
  Logger.info("Starting restore-deleted-users process");

  const members = getMembers(); // All active CAPWATCH members
  const wingDomain = CONFIG.EMAIL_DOMAIN; // e.g., "@hiwgcap.org"

  let restored = 0;
  let skipped = 0;
  let errors = 0;

  for (const capid in members) {
    const m = members[capid];

    // Build expected primary email
    const baseEmail = `${m.firstName}.${m.lastName}`
      .toLowerCase()
      .replace(/\s+/g, '');

    const primaryEmail = baseEmail + wingDomain;

    // 1. Does Workspace already have this user?
    let exists = true;
    try {
      const u = AdminDirectory.Users.get(primaryEmail, { projection: "full" });
      if (u.archived) {
        try {
          AdminDirectory.Users.update({ archived: false, suspended: false }, primaryEmail);
          Logger.info("Unarchived user during restore", { email: primaryEmail });
        } catch (errU) {
          Logger.error("Failed to unarchive during restore", {
            email: primaryEmail,
            error: errU.message
          });
          errors++;
          continue;
        }
      }
    } catch (e) {
      if (String(e.message).includes("Resource Not Found")) {
        exists = false;
      } else {
        Logger.error("Error checking user existence", {
          email: primaryEmail,
          error: e.message
        });
        errors++;
        continue;
      }
    }

    if (exists) {
      skipped++;
      continue;
    }

    // 2. User missing → restore
    try {
      addOrUpdateUser(m);
      restored++;
      Logger.info("Restored missing user", {
        capid: capid,
        email: primaryEmail
      });
    } catch (e) {
      Logger.error("Failed to restore missing user", {
        capid: capid,
        email: primaryEmail,
        error: e.message
      });
      errors++;
    }
  }

  Logger.info("Restore-deleted-users completed", {
    restored: restored,
    skippedExisting: skipped,
    errors: errors,
    totalMembers: Object.keys(members).length
  });
}

// ============================================================================
// TEST FUNCTIONS
// ============================================================================

/**
 * Test function for addOrUpdateUser with a specific member
 * @returns {void}
 */
function testaddOrUpdateUser() {
  Logger.info('Starting test - addOrUpdateUser');
  let members = getMembers();
  if (members[443777]) {
    addOrUpdateUser(members[443777]);
    Logger.info('Test completed');
  } else {
    Logger.error('Test member not found', { capsn: 443777 });
  }
}

/**
 * Test function to retrieve and display a specific member
 * @returns {void}
 */
function testGetMember() {
  Logger.info('Starting test - getMember');
  let members = getMembers();
  let member = members[105576];
  if (member) {
    Logger.info('Test member data', { member: member });
  } else {
    Logger.error('Test member not found', { capsn: 105576 });
  }
}

/**
 * Test function to retrieve and display squadron data
 * @returns {void}
 */
function testGetSquadrons() {
  Logger.info('Starting test - getSquadrons');
  let squadrons = getSquadrons();
  if (squadrons[2503]) {
    Logger.info('Test squadron data', { squadron: squadrons[2503] });
  } else {
    Logger.error('Test squadron not found', { orgid: 2503 });
  }
}

/**
 * Test function to save empty member data
 * @returns {void}
 */
function testSaveCurrentMembersData() {
  Logger.info('Starting test - saveCurrentMemberData');
  saveCurrentMemberData({});
  Logger.info('Test completed');
}

function updateSignatureForAllAliases(primaryEmail, signatureHTML) {
  const scope =
    'https://www.googleapis.com/auth/gmail.settings.sharing ' +
    'https://www.googleapis.com/auth/gmail.settings.basic';
  let accessToken;

  try {
    accessToken = getImpersonatedToken_(primaryEmail, scope);
  } catch (e) {
    Logger.error('Failed to get impersonation token', { user: primaryEmail, error: e.message });
    return;
  }

  try {
    const sendAsListUrl =
      `https://gmail.googleapis.com/gmail/v1/users/${encodeURIComponent(primaryEmail)}/settings/sendAs`;
    const listResponse = UrlFetchApp.fetch(sendAsListUrl, {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + accessToken },
      muteHttpExceptions: true
    });

    const listCode = listResponse.getResponseCode();
    if (listCode < 200 || listCode >= 300) {
      Logger.error('Failed to list send-as identities', {
        user: primaryEmail,
        code: listCode,
        response: listResponse.getContentText()
      });
      return;
    }

    const identities = JSON.parse(listResponse.getContentText()).sendAs || [];

    identities.forEach(identity => {
      const aliasEmail = identity.sendAsEmail;

      // 👉 Only update signatures for @pcrcap.org
      //if (!aliasEmail.toLowerCase().endsWith("@pcrcap.org")) {
      //  Logger.info("Skipping non-org alias", { alias: aliasEmail });
      //  return;
      //}

      const apiUrl =
        `https://gmail.googleapis.com/gmail/v1/users/${encodeURIComponent(primaryEmail)}/settings/sendAs/${encodeURIComponent(aliasEmail)}`;
      const payload = JSON.stringify({ signature: signatureHTML });

      const resp = UrlFetchApp.fetch(apiUrl, {
        method: 'patch',
        contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + accessToken },
        payload: payload,
        muteHttpExceptions: true
      });

      const code = resp.getResponseCode();
      if (code < 200 || code >= 300) {
        Logger.error('Failed to update signature for alias', {
          primary: primaryEmail,
          alias: aliasEmail,
          code: code,
          response: resp.getContentText()
        });
      }
    });

  } catch (e) {
    Logger.error('Error updating alias signatures', { user: primaryEmail, error: e.message });
  }
}

/**
 * Adds Gmail aliases and properly updates both alias and primary Send-As identities.
 * Uses service account impersonation with correct Gmail API rules.
 */
function addAliasesFromSheet() {
  Logger.info('Starting alias creation from sheet using direct impersonation.');

  const sheet = SpreadsheetApp
    .openById(CONFIG.AUTOMATION_SPREADSHEET_ID)
    .getSheetByName('Aliases');

  if (!sheet) {
    Logger.error('Aliases sheet not found');
    return;
  }

  const data = sheet.getDataRange().getValues();
  data.shift(); // remove header row

  let totalProcessed = 0;
  let totalAdded = 0;
  let totalFailed = 0;
  let totalSkipped = 0;

  const scope =
    'https://www.googleapis.com/auth/gmail.settings.basic ' +
    'https://www.googleapis.com/auth/gmail.settings.sharing';

  for (let i = 0; i < data.length; i++) {
    const primaryEmail = (data[i][0] || '').trim();
    const aliasEmail = (data[i][1] || '').trim();
    const displayName = (data[i][2] || '').trim();
    const signature = data[i][3] || '';

    if (!primaryEmail || !aliasEmail) continue;
    totalProcessed++;

    // --- Check if user is an admin ---
    try {
      const user = AdminDirectory.Users.get(primaryEmail, { fields: "isAdmin" });
      if (user.isAdmin) {
        Logger.info("Skipping admin user (aliases must be set manually)", {
          user: primaryEmail
        });
        totalSkipped++;
        continue;
      }
    } catch (e) {
      Logger.error("Admin check failed", {
        user: primaryEmail,
        error: e.message
      });
      totalFailed++;
      continue;
    }

    // --- Get impersonated token ---
    let accessToken = "";
    try {
      accessToken = getImpersonatedToken_(primaryEmail, scope);
    } catch (e) {
      Logger.error("Could not impersonate user", {
        user: primaryEmail,
        error: e.message
      });
      totalFailed++;
      continue;
    }

    //
    // ---------------------------------------------------------
    //  STEP 1 — CREATE / UPDATE THE ALIAS (POST)
    // ---------------------------------------------------------
    //
    try {
      const sendAsAliasBody = {
        sendAsEmail: aliasEmail,
        displayName: displayName,
        signature: signature,
        treatAsAlias: true
      };

      const aliasUrl =
        "https://gmail.googleapis.com/gmail/v1/users/me/settings/sendAs";

      const aliasResp = UrlFetchApp.fetch(aliasUrl, {
        method: "post",
        contentType: "application/json",
        headers: { Authorization: "Bearer " + accessToken },
        payload: JSON.stringify(sendAsAliasBody),
        muteHttpExceptions: true
      });

      const code = aliasResp.getResponseCode();

      if (code >= 200 && code < 300) {
        Logger.info("Alias added successfully", {
          primary: primaryEmail,
          alias: aliasEmail
        });
        totalAdded++;
      } else if (code === 409) {
        Logger.info("Alias already exists, skipping", {
          primary: primaryEmail,
          alias: aliasEmail
        });
        totalSkipped++;
      } else {
        Logger.error("Alias creation failed", {
          primary: primaryEmail,
          alias: aliasEmail,
          code: code,
          response: aliasResp.getContentText()
        });
        totalFailed++;
      }

    } catch (e) {
      Logger.error("Unhandled error during alias POST", {
        user: primaryEmail,
        alias: aliasEmail,
        error: e.message
      });
      totalFailed++;
      continue;
    }

    //
    // ---------------------------------------------------------
    //  STEP 2 — FIX PRIMARY SEND-AS NAME (PATCH)
    // ---------------------------------------------------------
    //
    try {
      /**
       * Gmail rules for PRIMARY send-as:
       *   URL MUST be:     /users/{userId}/settings/sendAs/{sendAsEmail}
       *   DO NOT send:     sendAsEmail, treatAsAlias
       *   Allowed fields:  displayName, signature
       */

      const primaryPatchUrl =
        `https://gmail.googleapis.com/gmail/v1/users/${encodeURIComponent(primaryEmail)}/settings/sendAs/${encodeURIComponent(primaryEmail)}`;

      const primaryPatchBody = {
        displayName: displayName,
        signature: signature
      };

      const primaryResp = UrlFetchApp.fetch(primaryPatchUrl, {
        method: "patch",
        contentType: "application/json",
        headers: { Authorization: "Bearer " + accessToken },
        payload: JSON.stringify(primaryPatchBody),
        muteHttpExceptions: true
      });

      const pcode = primaryResp.getResponseCode();

      if (pcode >= 200 && pcode < 300) {
        Logger.info("Primary Send-As display name updated", {
          primary: primaryEmail,
          displayName: displayName
        });
      } else {
        Logger.warn("Primary Send-As update failed", {
          primary: primaryEmail,
          code: pcode,
          response: primaryResp.getContentText()
        });
        totalFailed++;
      }

    } catch (e) {
      Logger.error("Primary Send-As update threw exception", {
        user: primaryEmail,
        error: e.message
      });
      totalFailed++;
    }
  }

  Logger.info("Alias creation completed", {
    processed: totalProcessed,
    added: totalAdded,
    failed: totalFailed,
    skipped: totalSkipped
  });
}

function updateAllSendAsNames() {
  Logger.info('Starting updateAllSendAsNames');

  // 1. Fetch CAPWATCH members → build lookup
  const members = getMembers();
  const memberIndex = {};
  for (const capid in members) {
    memberIndex[String(capid)] = members[capid];
  }

  // 2. Fetch all Workspace users
  let pageToken = null;
  const workspaceUsers = [];

  do {
    const page = AdminDirectory.Users.list({
      customer: "my_customer",
      maxResults: 200,
      projection: "full",
      pageToken: pageToken
    });

    if (page.users) workspaceUsers.push(...page.users);
    pageToken = page.nextPageToken;

  } while (pageToken);

  // 3. Loop users and update Send-As display name
  workspaceUsers.forEach(user => {
    const capid = user.externalIds?.[0]?.value;
    const member = memberIndex[capid];

    if (!member) {
      Logger.warn('No CAPWATCH record for user', { email: user.primaryEmail });
      return;
    }

    // Build display name exactly like addOrUpdateUser()
    const displayName = [
      member.lastName + (member.suffix ? ' ' + member.suffix : ''),
      ', ',
      member.firstName,
      member.middleName ? ' ' + member.middleName.charAt(0) : '',
      member.rank ? ' ' + member.rank : ''
    ].join('').trim();

    // Sync Directory displayName everywhere
    try {
      AdminDirectory.Users.update(
        { name: { displayName: displayName } },
        user.primaryEmail
      );
      Logger.info('Directory displayName synced', {
        email: user.primaryEmail,
        displayName: displayName
      });
    } catch (e) {
      Logger.error('Failed to sync Directory displayName', {
        email: user.primaryEmail,
        error: e.message
      });
    }
    updateGmailSendAsDisplayName(user.primaryEmail, displayName);
    Utilities.sleep(200);
  });

  Logger.info('Completed updateAllSendAsNames for all Workspace users');
}

/**
 * FORCE UPDATE ALL MEMBERS
 * Forces an update of all members regardless of whether they've changed.
 * Bypasses the memberUpdated() check and processes every member.
 * Useful for ensuring all users have the latest field updates applied.
 */
function forceUpdateAllMembers() {
  clearCache();
  const start = new Date();

  Logger.info('Starting FORCE update of all members (bypassing change detection)');

  let members = getMembers();

  const activeUsers = getActiveUsers();
  workspaceEmailByCapid = {};
  activeUsers.forEach(u => {
    workspaceEmailByCapid[String(u.capid)] = u.email;
  });
  assignManagerEmails(members);

  const totalMembers = Object.keys(members).length;

  Logger.info("Force update starting", {
    totalMembers: totalMembers
  });

  // Process ALL members without change detection
  batchUpdateMembers(members);

  saveCurrentMemberData(members);

  Logger.info('Force member update completed', {
    duration: new Date() - start + 'ms',
    totalProcessed: totalMembers
  });
}

/**
 * TEST FUNCTION — Force update of Directory duty titles
 * for all users with no duty assignments.
 *
 * This runs independently of updateAllMembers() and does NOT
 * alter any previous data comparison logic.
 */
function forceDutyTitleRebuild() {
  Logger.info("Starting forceDutyTitleRebuild()");

  const members = getMembers();
  const memberIndex = {};
  for (const capid in members) {
    memberIndex[String(capid)] = members[capid];
  }

  // Fetch all Workspace users
  let pageToken = null;
  const workspaceUsers = [];

  do {
    const page = AdminDirectory.Users.list({
      customer: "my_customer",
      maxResults: 200,
      projection: "full",
      pageToken: pageToken
    });

    if (page.users) workspaceUsers.push(...page.users);
    pageToken = page.nextPageToken;

  } while (pageToken);

  let updated = 0;
  let skipped = 0;

  workspaceUsers.forEach(user => {
    const capid = user.externalIds?.[0]?.value;
    const member = memberIndex[capid];

    if (!member) {
      skipped++;
      return;
    }

    // Only users WITH NO duty positions
    if (member.dutyPositions.length > 0) {
      skipped++;
      return;
    }

    // Build the updated org title
    const orgTitle = "No Duty Assignment";

    try {
      AdminDirectory.Users.update({
        organizations: [{
          title: orgTitle,
          department: toTitleCase(member.orgName || member.charter || ''),
          costCenter: member.charter || '',
          description: member.type || '',
          type: 'work',
          primary: true
        }]
      }, user.primaryEmail);

      Logger.info("Forced duty-title update", {
        email: user.primaryEmail,
        title: orgTitle
      });

      updated++;

    } catch (e) {
      Logger.error("Failed duty-title force update", {
        email: user.primaryEmail,
        error: e.message
      });
    }

    Utilities.sleep(150); // rate limit protection
  });

  Logger.info("forceDutyTitleRebuild() completed", {
    updated: updated,
    skipped: skipped,
    totalUsers: workspaceUsers.length
  });
}

function testImpersonationToken() {
  try {
    const who = Session.getEffectiveUser().getEmail();
    const t = getImpersonatedToken_(
      who,
      'https://www.googleapis.com/auth/gmail.settings.basic'
    );
    console.log('Impersonation OK for ' + who + ' — token length: ' + (t ? t.length : 0));
  } catch (e) {
    console.log('Impersonation FAILED: ' + (e && e.message ? e.message : e));
  }
}
