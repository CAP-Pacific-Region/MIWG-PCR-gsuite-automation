/*******************************************************
 * Group Membership Synchronization Module
 *
 * Version: 1.3.7
 * Filename: UpdateGroups.gs
 * Saved: 2026-04-06 16:17 PDT
 *
 * Manages Google Groups memberships based on CAPWATCH data and configuration:
 * - Reads group configuration from automation spreadsheet
 * - Builds group memberships based on member attributes (type, rank, duty positions, committee assignments, etc.)
 * - Creates both wing-level and group-level groups automatically
 * - Calculates membership deltas (add/remove changes)
 * - Applies changes to Google Workspace groups
 * - Auto-creates groups that don't exist
 * - Supports manual member additions via spreadsheet
 * - Tracks and logs errors to spreadsheet for review
 * - Fixed external contacts and parent/guardian emails
 * - FIxed Group name and description
 * - Added new column to Groups tab EXT where you can create a Group as allow external members
 *******************************************************/

var workspaceUsers = {};
var workspaceEmailMap = {};

// Desired metadata (name/description) computed from the Groups sheet for logging and group creation
var desiredGroupMeta = {};
// Map base group name -> Attribute from the Groups sheet (built in getEmailGroupDeltas)
var groupAttributeByName = {};
// Map base group name -> allow external members boolean (from Groups sheet Add EXT column)
var groupAllowExternalByName = {};

const DRY_RUN = false; // change to false for real updates

function updateEmailGroups() {
  clearCache();
  const start = new Date();
  let deltas = getEmailGroupDeltas();
  let errorEmails = {};
  let totalAdded = 0;
  let totalRemoved = 0;
  let totalErrors = 0;
  let processedCategories = 0;
  const totalCategories = Object.keys(deltas).length;

  // For dry-run summary
  let dryRunSummary = [];

  for(const category in deltas) {
    processedCategories++;
    for (const group in deltas[category]) {
      let added = 0;
      let removed = 0;
      const groupEmail = group + CONFIG.EMAIL_DOMAIN;
      const baseGroupName = group.includes('.') ? group.split('.').slice(1).join('.') : group;

      // Ensure existing group name/description match desired metadata from Groups sheet
      const metaForGroup = desiredGroupMeta[groupEmail.toLowerCase()] || {};
      applyGroupMeta_(groupEmail, metaForGroup);
      applyManagedGroupSettings_(groupEmail, {
        allowExternalMembers: !!groupAllowExternalByName[baseGroupName],
        whoCanViewMembership: 'ALL_IN_DOMAIN_CAN_VIEW',
        whoCanPostMessage: 'ANYONE_CAN_POST'
      });

      let dryRunMembers = [];

      for (const email in deltas[category][group]) {
        switch(deltas[category][group][email]) {
          case -1:
            // Remove member
            try {
              const finalEmail = workspaceEmailMap[email.replace(/\D/g, '')] || email;
              if (DRY_RUN) {
                Logger.info('💡 [Dry-Run] Would remove member', {
                  member: email,
                  group: groupEmail
                });
                dryRunMembers.push({ email: finalEmail, action: 'REMOVE' });
              } else {
                executeWithRetry(() =>
                  AdminDirectory.Members.remove(groupEmail, finalEmail)
                );
                removed++;
                  Logger.info('Removed member from group', {
                    member: email,
                    group: groupEmail
                  });
              }
            } catch (e) {
              Logger.error('Failed to remove member from group', {
                member: email,
                group: groupEmail,
                category: category,
                errorMessage: e.message,
                errorCode: e.details?.code,
                errorReason: e.details?.errors?.[0]?.reason
              });

              // Track removal errors too
              if (!errorEmails[email]) {
                errorEmails[email] = {
                  email: email,
                  attempts: [],
                  firstSeen: new Date().toISOString()
                };
              }
              errorEmails[email].attempts.push({
                group: group,
                groupEmail: groupEmail,
                category: category,
                action: 'REMOVE',
                errorCode: e.details?.code || 'Unknown',
                errorMessage: e.message || 'Unknown error',
                timestamp: new Date().toISOString()
              });

              totalErrors++;
            }
            break;
          case 1:
            // Add member
            try {
              const finalEmail = workspaceEmailMap[email.replace(/\D/g, '')] || email;

              // Skip external emails except for groups whose Attribute is 'contact'
              if (!finalEmail.endsWith(CONFIG.EMAIL_DOMAIN) && groupAttributeByName[category] !== 'contact') continue;

              if (DRY_RUN) {
                Logger.info('💡 [Dry-Run] Would add member', {
                  member: finalEmail,
                  group: groupEmail
                });
                dryRunMembers.push({ email: finalEmail, action: 'ADD' });
                // Continue to next member, do not actually add
                continue;
              }

              executeWithRetry(() =>
                AdminDirectory.Members.insert({
                  email: finalEmail,
                  role: 'MEMBER'
                }, groupEmail)
              );
              added++;
                Logger.info('Added member to group', {
                  member: email,
                  group: groupEmail
                });

                // Throttle between API insert calls
                Utilities.sleep(CONFIG.API_DELAY_MS);

                // Periodic quota cooldown
                if (added > 0 && added % 25 === 0) {
                  Logger.info('Pausing briefly to allow API quota refill', { added });
                  Utilities.sleep(15000); // 15 sec every 25 adds
                }
            } catch (e) {
              // Check if member is already in group (409 = Conflict/Duplicate)
              if (e.details?.code === 409) {
                Logger.warn('Member already in group', {
                  member: email,
                  group: groupEmail,
                  category: category
                });
              }
              else if (e.details?.code === 404) {
                Logger.warn('Cannot add external member - not found', {
                  member: email,
                  group: groupEmail,
                  category: category,
                  note: 'Email may not exist or group settings prevent external members'
                });

                // Track detailed error info
                if (!errorEmails[email]) {
                  errorEmails[email] = {
                    email: email,
                    attempts: [],
                    firstSeen: new Date().toISOString()
                  };
                }
                errorEmails[email].attempts.push({
                  group: group,
                  groupEmail: groupEmail,
                  category: category,
                  errorCode: 404,
                  errorMessage: 'Resource Not Found',
                  timestamp: new Date().toISOString()
                });
              }
              // All other errors
              else {
                Logger.error('Failed to add member to group', {
                  member: email,
                  group: groupEmail,
                  category: category,
                  errorMessage: e.message,
                  errorCode: e.details?.code,
                  errorReason: e.details?.errors?.[0]?.reason,
                  fullError: JSON.stringify(e.details)
                });

                // Track detailed error info
                if (!errorEmails[email]) {
                  errorEmails[email] = {
                    email: email,
                    attempts: [],
                    firstSeen: new Date().toISOString()
                  };
                }
                errorEmails[email].attempts.push({
                  group: group,
                  groupEmail: groupEmail,
                  category: category,
                  errorCode: e.details?.code || 'Unknown',
                  errorMessage: e.message || 'Unknown error',
                  errorReason: e.details?.errors?.[0]?.reason || 'Unknown',
                  timestamp: new Date().toISOString()
                });
              }

              totalErrors++;
            }
            break;
          case 0:
            // Member already in group
            break;
        }
      }

      totalAdded += added;
      totalRemoved += removed;

      if (DRY_RUN && dryRunMembers.length > 0) {
        dryRunSummary.push({
          group: groupEmail,
          category: category,
          members: dryRunMembers
        });

        try {
          let folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
          let dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
          let safeGroup = group.replace(/[^a-zA-Z0-9.-]/g, '_');
          let fileName = `DryRun-${safeGroup}-${dateStr}.csv`;

          // Use same columns as members_template.csv
          let csvHeader = 'Group Email [Required],Member Email,Member Type,Member Role\n';
          let csvContent = csvHeader;

          dryRunMembers.forEach(m => {
            let memberType = m.action === 'ADD' ? 'User' : 'Removed';
            let memberRole = 'MEMBER';
            csvContent += `${groupEmail},${m.email},${memberType},${memberRole}\n`;
          });

          let file = folder.createFile(fileName, csvContent, MimeType.CSV);
          Logger.info('💡 [Dry-Run] Group CSV saved', {
            fileName: fileName,
            url: file.getUrl(),
            memberCount: dryRunMembers.length
          });
        } catch (e) {
          Logger.error('💡 [Dry-Run] Failed to save CSV for group', {
            group: groupEmail,
            error: e.message
          });
        }
      }

      const meta = desiredGroupMeta[groupEmail.toLowerCase()] || {};
      Logger.info('Updated group', {
        groupId: group,
        group: groupEmail,
        name: meta.name || '',
        description: meta.description || '',
        added: added,
        removed: removed
      });
    }
    if (processedCategories % 5 === 0 || processedCategories === totalCategories) {
      Logger.info('Progress update', {
        processed: processedCategories,
        total: totalCategories,
        percentComplete: Math.round((processedCategories / totalCategories) * 100)
      });
    }
  }

  saveErrorEmails(errorEmails);

  // Dry-run: Save summary file and log
  if (DRY_RUN) {
    try {
      let folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
      let dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
      let fileName = 'DryRun-Groups-' + dateStr + '.json';
      let content = JSON.stringify(dryRunSummary, null, 2);
      let file = folder.createFile(fileName, content, MimeType.PLAIN_TEXT);
      Logger.info('💡 [Dry-Run] Summary saved', {
        url: file.getUrl(),
        groupCount: dryRunSummary.length
      });
    } catch (e) {
      Logger.error('💡 [Dry-Run] Failed to save summary', {
        error: e.message
      });
    }
  }

  Logger.info('Email group update completed', {
    duration: new Date() - start + 'ms',
    totalAdded: totalAdded,
    totalRemoved: totalRemoved,
    totalErrors: totalErrors,
    errorEmailsCount: Object.keys(errorEmails).length
  });
}

/**
 * Loads CAPWATCH committee membership and attaches it to the in memory `members` map.
 *
 * Adds:
 *   members[capid].committeeAssignments = [{ orgid: ..., name: ... }, ...]
 *
 * Only WING and GROUP scoped committees are attached (no squadron/unit committees).
 *
 * Source:
 *   parseFile('MbrCommittee') rows shaped like: [CAPID, Committee, Chair, ORGID, DateAssigned]
 *
 * @param {Object<string, Object>} members - Members object indexed by CAPID
 * @param {Object<string, Object>} squadrons - Squadrons object indexed by orgid
 * @returns {void}
 */
function attachCommitteesToMembers(members, squadrons) {
  let rows = [];
  try {
    rows = parseFile('MbrCommittee');
  } catch (e) {
    Logger.warn('MbrCommittee not available; committeeAssignments will be empty', { error: e.message });
    return;
  }

  // Initialize committeeAssignments arrays for every member we know about
  for (const capid in members) {
    if (!Array.isArray(members[capid].committeeAssignments)) {
      members[capid].committeeAssignments = [];
    }
  }

  // CAPWATCH MbrCommittee columns:
  // [0]=CAPID, [1]=Committee, [2]=Chair, [3]=ORGID, [4]=DateAssigned
  let attached = 0;
  let skippedNoMember = 0;
  let skippedMissing = 0;
  let skippedNonWingGroup = 0;

  for (let i = 0; i < rows.length; i++) {
    const capid = (rows[i][0] || '').toString().trim();
    const committeeName = (rows[i][1] || '').toString().trim();
    const orgid = (rows[i][3] || '').toString().trim();

    if (!capid || !committeeName || !orgid) {
      skippedMissing++;
      continue;
    }

    if (!members[capid]) {
      skippedNoMember++;
      continue;
    }

    const org = squadrons && squadrons[orgid] ? squadrons[orgid] : null;
    if (!org || (org.scope !== 'WING' && org.scope !== 'GROUP')) {
      // Explicit requirement: ONLY wing + group committees, no squadrons/units
      skippedNonWingGroup++;
      continue;
    }

    const assigns = members[capid].committeeAssignments;
    const exists = assigns.some(a => a && a.orgid === orgid && a.name === committeeName);
    if (!exists) {
      assigns.push({ orgid: orgid, name: committeeName });
      attached++;
    }
  }

  Logger.info('Attached committeeAssignments to members (WING+GROUP only)', {
    membersCount: Object.keys(members).length,
    rows: rows.length,
    attached: attached,
    skippedNoMember: skippedNoMember,
    skippedMissing: skippedMissing,
    skippedNonWingGroup: skippedNonWingGroup
  });
}

/**
 * Converts a name to title style capitalization.
 * - Converts ALL CAPS words to Title Case (e.g., "CALIFORNIA" -> "California")
 * - Preserves known acronyms (e.g., CAP, FAA, USAF) and WG-style acronyms (e.g., CAWG, HIWG)
 * - Preserves tokens with digits
 * @param {string} s
 * @returns {string}
 */
function toSentenceCase_(s) {
  const str = (s || '').toString().trim();
  if (!str) return '';

  const preserve = new Set([
    'CAP', 'USAF', 'FAA', 'DOT', 'TSA', 'ICAO', 'EASA', 'HQ'
  ]);

  function isWingAcronym_(tok) {
    // CAWG, HIWG, ORWG, NVWG, PCR, etc.
    return /^[A-Z]{2,4}WG$/.test(tok) || tok === 'PCR';
  }

  function titleToken_(tok) {
    if (!tok) return tok;
    if (/\d/.test(tok)) return tok;

    // Strip trailing punctuation for evaluation, re-attach later
    const m = tok.match(/^(.+?)([.,;:)]?)$/);
    const core = m ? m[1] : tok;
    const punct = m ? m[2] : '';

    const upper = core.toUpperCase();

    // Preserve known acronyms / wing acronyms
    if (preserve.has(upper) || isWingAcronym_(upper)) return upper + punct;

    // If it's ALL CAPS, convert to Title Case
    if (core === upper) {
      return (upper.charAt(0) + upper.slice(1).toLowerCase()) + punct;
    }

    // Otherwise just capitalize first letter and lower the rest
    return (core.charAt(0).toUpperCase() + core.slice(1).toLowerCase()) + punct;
  }

  return str
    .split(/\s+/)
    .map(titleToken_)
    .join(' ');
}

function isCAWGTenant_() {
  return String((CONFIG && CONFIG.WING) || '').trim().toUpperCase() === 'CA';
}

function stripLeadingHonorific_(value) {
  return String(value || '')
    .replace(/^(lt\.?\s*col|col|maj|capt|1st\s*lt|2nd\s*lt|lt)\s+/i, '')
    .trim();
}

function abbreviateManagedGroupCommonShortName_(value) {
  const normalized = String(value || '').trim();
  if (!normalized) return '';

  if (/^los angeles county$/i.test(normalized)) return 'LA County';
  if (/^san francisco bay$/i.test(normalized)) return 'SF Bay';

  return normalized;
}

function abbreviateManagedGroupOrgDisplayName_(org) {
  if (!org || !org.name) return '';

  const fullName = toSentenceCase_(String(org.name || '').trim());
  if (!fullName) return '';

  if (!isCAWGTenant_()) return fullName;

  const scope = String(org.scope || '').trim().toUpperCase();
  const unit = String(org.unit || '').trim().replace(/^0+/, '');

  if (scope === 'WING') {
    return 'CAWG';
  }

  if (scope === 'GROUP') {
    const match = fullName.match(/^(.*)\bGroup\s+(\d+)\b$/i);
    if (match) {
      const shortName = abbreviateManagedGroupCommonShortName_(String(match[1] || '').trim());
      const number = String(match[2] || '').trim();
      return shortName ? `Grp ${number} ${shortName}` : `Grp ${number}`;
    }
    return unit ? `Grp ${unit} ${fullName}` : fullName;
  }

  const leadingNumberUnitMatch = fullName.match(/^(\d+(?:st|nd|rd|th))\s+(.*?)(?:\s+(?:Challenger\s+)?)?(?:Cadet|Composite)?\s*Sq(?:dn|uadron)?\b$/i);
  if (leadingNumberUnitMatch) {
    const number = String(leadingNumberUnitMatch[1] || '').trim().toUpperCase();
    const shortName = abbreviateManagedGroupCommonShortName_(
      stripLeadingHonorific_(String(leadingNumberUnitMatch[2] || '').trim())
    );
    return shortName ? `Sqdn ${number} ${shortName}` : `Sqdn ${number}`;
  }

  const unitMatch = fullName.match(/^(.*?)(?:\s+(?:Challenger\s+)?)?(?:Cadet|Composite)?\s*Sq(?:dn|uadron)?\s+(\d+)\b$/i)
    || fullName.match(/^(.*?)\s+Squadron\s+(\d+)\b$/i);
  if (unitMatch) {
    const shortName = abbreviateManagedGroupCommonShortName_(
      stripLeadingHonorific_(String(unitMatch[1] || '').trim())
    );
    const number = String(unitMatch[2] || '').trim();
    return shortName ? `Sqdn ${number} ${shortName}` : `Sqdn ${number}`;
  }

  return unit ? `Sqdn ${unit} ${fullName}` : fullName;
}

/**
 * Ensures an existing Google Group has the desired name/description.
 * Uses PATCH so only provided fields are updated.
 * Dry-run aware.
 * @param {string} groupEmail Full group email address
 * @param {{name?: string, description?: string}} meta Desired metadata
 * @returns {void}
 */
function applyGroupMeta_(groupEmail, meta) {
  const email = (groupEmail || '').toString().toLowerCase();
  if (!email || !meta) return;

  const desiredName = (meta.name || '').toString().trim();
  const desiredDesc = (meta.description || '').toString().trim();
  if (!desiredName && !desiredDesc) return;

  try {
    const existing = AdminDirectory.Groups.get(email);
    const currentName = (existing.name || '').toString().trim();
    const currentDesc = (existing.description || '').toString().trim();

    const patch = {};
    if (desiredName && desiredName !== currentName) patch.name = desiredName;
    if (desiredDesc && desiredDesc !== currentDesc) patch.description = desiredDesc;

    if (Object.keys(patch).length === 0) return;

    if (DRY_RUN) {
      Logger.info('💡 [Dry-Run] Would update group metadata', {
        group: email,
        fromName: currentName,
        toName: desiredName,
        fromDescription: currentDesc,
        toDescription: desiredDesc
      });
      return;
    }

    executeWithRetry(() => AdminDirectory.Groups.patch(patch, email));
    Logger.info('Updated group metadata', { group: email, ...patch });
  } catch (e) {
    Logger.warn('Failed to update group metadata', {
      group: email,
      errorMessage: e.message,
      errorCode: e.details?.code
    });
  }
}

/**
 * Calculates email group membership deltas by comparing desired state with current state
 * Returns object with delta values: 1 = add, 0 = no change, -1 = remove
 * @returns {Object} Groups object with delta values for each member
 */
function getEmailGroupDeltas() {
  const start = new Date();
  let groups = {};
  let groupsConfig = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID).getSheetByName('Groups').getDataRange().getValues();

// Build Group Name -> Attribute lookup for use during updateEmailGroups filtering
groupAttributeByName = {};
groupAllowExternalByName = {};

const groupsHeader = (groupsConfig[0] || []).map(h => (h || '').toString().trim().toLowerCase());
const addExtIdx = groupsHeader.indexOf('add ext');

function isTruthyAddExt_(v) {
  const t = (v || '').toString().trim().toLowerCase();
  return t === 'y' || t === 'yes' || t === 'x' || t === 'true';
}

for (let r = 1; r < groupsConfig.length; r++) {
  const gName = (groupsConfig[r][1] || '').toString().trim(); // Group Name
  const attr = (groupsConfig[r][2] || '').toString().trim();  // Attribute
  if (gName) groupAttributeByName[gName] = attr;
  if (gName) groupAllowExternalByName[gName] = addExtIdx > -1 ? isTruthyAddExt_(groupsConfig[r][addExtIdx]) : false;
}

  // Map base group name -> spreadsheet description (if provided)
  // Expected Groups sheet columns:
  // [0]=Category, [1]=Group Name, [2]=Attribute, [3]=Values, [4]=Description
  const descriptionByGroupName = {};
  for (let r = 1; r < groupsConfig.length; r++) {
    const gName = (groupsConfig[r][1] || '').toString().trim();
    const desc = (groupsConfig[r][4] || '').toString().trim();
    if (gName) {
      descriptionByGroupName[gName] = desc;
    }
  }
  let squadrons = getSquadrons();
  let members = getMembers();
  // --- Build CAPWATCH → Workspace email map ---
  attachCommitteesToMembers(members, squadrons);
  workspaceEmailMap = {};
  let token = '';
  try {
    do {
      const page = AdminDirectory.Users.list({
        domain: CONFIG.DOMAIN,
        maxResults: 500,
        projection: 'full',
        fields: 'users(primaryEmail,externalIds),nextPageToken',
        pageToken: token
      });
      if (page.users) {
        page.users.forEach(u => {
          const capidField = (u.externalIds || []).find(x => x.type === 'organization');
          if (capidField && capidField.value) {
            workspaceEmailMap[capidField.value.toString()] = u.primaryEmail.toLowerCase();
            if (members[capidField.value]) members[capidField.value].email = u.primaryEmail.toLowerCase();
          }
        });
      }
      token = page.nextPageToken;
    } while (token);
    Logger.info('Workspace CAPID→Email map built', { count: Object.keys(workspaceEmailMap).length });
  } catch (err) {
    Logger.error('Failed to build Workspace CAPID→Email map', { message: err.message });
  }

  // --- Build Workspace user lookup map (for internal members only) ---
  workspaceUsers = {};
  let pageToken = '';
  try {
    do {
      const res = AdminDirectory.Users.list({
        domain: CONFIG.DOMAIN,
        maxResults: 500,
        projection: 'basic',
        fields: 'users(primaryEmail),nextPageToken',
        pageToken: pageToken
      });
      if (res.users) {
        res.users.forEach(u => {
          workspaceUsers[u.primaryEmail.toLowerCase()] = true;
        });
      }
      pageToken = res.nextPageToken;
    } while (pageToken);
    Logger.info('Loaded Workspace user list', {
      count: Object.keys(workspaceUsers).length
    });
  } catch (err) {
    Logger.error('Failed to build Workspace user map', { error: err.message });
  }

  // Build desired group membership state
  for(let i = 1; i < groupsConfig.length; i++) {
    const groupName = groupsConfig[i][1];
    const generatedGroups = getGroupMembers(
      groupName,
      groupsConfig[i][2],
      groupsConfig[i][3],
      members,
      squadrons,
    );

    // Multiple sheet rows may intentionally target the same base group name.
    // Merge their generated memberships instead of letting the last row win.
    if (!groups[groupName]) groups[groupName] = {};
    for (const groupId in generatedGroups) {
      if (!groups[groupName][groupId]) groups[groupName][groupId] = {};
      for (const email in generatedGroups[groupId]) {
        groups[groupName][groupId][email] = generatedGroups[groupId][email];
      }
    }
  }

  // Preserve manual members from "User Additions" by treating them as desired
  // memberships before the current-vs-desired delta pass.
  const manualByGroup = getManualGroupMembersFromUserAdditions_();
  const mergeStats = mergeManualGroupMembersIntoDesired_(groups, manualByGroup);
  Logger.info('Manual User Additions merged into desired memberships', mergeStats);

  // Calculate deltas by comparing with current state
  for (const category in groups) {
    for (const group in groups[category]) {
      // Use Groups sheet Description column for group creation whenever provided.
      // group is like "hiwg.<name>", "hi073.<name>", etc. Base name is everything after the first dot.
      const baseGroupName = group.includes('.') ? group.split('.').slice(1).join('.') : group;

      let spreadsheetDescription = (descriptionByGroupName[baseGroupName] || '').toString().trim();

      // Duty-position groups: if Description is blank, fall back to Values column as description.
      if (!spreadsheetDescription && baseGroupName.indexOf('dty.') > -1) {
        const cfgRow = groupsConfig.find(row => (row[1] || '').toString().trim() === baseGroupName);
        const valuesCol = cfgRow ? (cfgRow[3] || '').toString().trim() : '';
        spreadsheetDescription = valuesCol || 'Unknown Duty Position';
      }

      // Compute friendly group metadata:
      // - description: from spreadsheet (or fallback)
      // - name: "<Org Name> - <description>"
      const groupEmail = (group + CONFIG.EMAIL_DOMAIN).toLowerCase();
      const metaDescriptionRaw = (spreadsheetDescription || baseGroupName).toString().trim();
      const metaDescription = metaDescriptionRaw;
      let metaNameSuffix = metaDescriptionRaw;
      const groupAttribute = String(groupAttributeByName[baseGroupName] || '').trim().toLowerCase();

      // Achievement descriptions often follow "ABBR - Full Name".
      // Use the short code for the group name while keeping the full
      // description unchanged for directory details.
      if (groupAttribute === 'achievements' && metaDescriptionRaw.indexOf(' - ') > -1) {
        metaNameSuffix = String(metaDescriptionRaw.split(' - ')[0] || '').trim() || metaDescriptionRaw;
      }

      // Determine the org display name based on the groupId prefix (wing-level "ca.*" or unit-level "ca008.*")
      const groupPrefix = (group.split('.')[0] || '').toString().trim().toLowerCase();
      let orgRecord = null;

      try {
        if (groupPrefix === CONFIG.WING.toLowerCase()) {
          // Wing-level group: find the WING org record
          orgRecord = Object.values(squadrons).find(o =>
            o && String(o.scope || '').toUpperCase() === 'WING' &&
            String(o.wing || '').toLowerCase() === CONFIG.WING.toLowerCase()
          );
        } else {
          // Unit/group prefix like "ca008" or "ca445": match by wing+unit
          const wing = groupPrefix.substring(0, 2);
          const unit = groupPrefix.substring(2);
          orgRecord = Object.values(squadrons).find(o =>
            o &&
            String(o.wing || '').toLowerCase() === wing &&
            String(o.unit || '') === unit
          );
        }
      } catch (e) {
        orgRecord = null;
      }

      const orgDisplayName = orgRecord && orgRecord.name ? toSentenceCase_(String(orgRecord.name || '')) : '';
      const shortOrgDisplayName = abbreviateManagedGroupOrgDisplayName_(orgRecord);
      const metaName = (shortOrgDisplayName ? `${shortOrgDisplayName} - ${metaNameSuffix}` : metaNameSuffix);
      const metaFullDescription = (orgDisplayName ? `${orgDisplayName} - ${metaDescription}` : metaDescription);

      desiredGroupMeta[groupEmail] = {
        name: metaName,
        description: metaFullDescription
      };

      const allowExternal = !!groupAllowExternalByName[baseGroupName];
      const currentMembers = getCurrentGroup(group, squadrons, desiredGroupMeta[groupEmail], allowExternal);
      for (let i = 0; i < currentMembers.length; i++) {
        const currentEmail = currentMembers[i].email;
        const currentRole = (currentMembers[i].role || 'MEMBER').toString().toUpperCase();

        if (groups[category][group][currentEmail]) {
          // Member already in group - no change needed
          groups[category][group][currentEmail] = 0;
        } else if (currentRole === 'MEMBER') {
          // Only auto-remove plain members. Leave MANAGER/OWNER entries alone
          // unless they are explicitly managed elsewhere (for example User Additions).
          groups[category][group][currentEmail] = -1;
        }
      }
    }
  }

  saveEmailGroups(groups);
  Logger.info('Group deltas generated', {
    duration: new Date() - start + 'ms',
    categories: Object.keys(groups).length
  });
  return groups;
}

/**
 * Loads manual member/group mappings from "User Additions".
 * Expected columns (same as updateAdditionalGroupMembers):
 * - [1] Email
 * - [3] Groups (comma-separated group IDs, optionally full group emails)
 *
 * @returns {Object<string, Object<string, number>>} groupId -> { email: 1 }
 */
function getManualGroupMembersFromUserAdditions_() {
  const out = {};
  try {
    const sheet = SpreadsheetApp
      .openById(CONFIG.AUTOMATION_SPREADSHEET_ID)
      .getSheetByName('User Additions');
    if (!sheet) {
      Logger.warn('User Additions tab not found; skipping manual preserve merge');
      return out;
    }

    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const email = (rows[i][1] || '').toString().trim().toLowerCase();
      const groupsCell = (rows[i][3] || '').toString().trim();
      if (!email || !groupsCell) continue;

      const groupTokens = groupsCell.split(',')
        .map(g => (g || '').toString().trim().toLowerCase())
        .filter(Boolean);

      for (let j = 0; j < groupTokens.length; j++) {
        let groupId = groupTokens[j];
        if (groupId.endsWith(CONFIG.EMAIL_DOMAIN.toLowerCase())) {
          groupId = groupId.slice(0, -CONFIG.EMAIL_DOMAIN.length);
        }
        if (!groupId) continue;

        if (!out[groupId]) out[groupId] = {};
        out[groupId][email] = 1;
      }
    }

    Logger.info('Manual members loaded from User Additions', {
      groups: Object.keys(out).length
    });
  } catch (e) {
    Logger.warn('Failed to load User Additions for preserve merge', {
      errorMessage: e.message
    });
  }
  return out;
}

/**
 * Merges manual group members into desired memberships across matching groups.
 *
 * @param {Object} groups Desired memberships by category/group/email
 * @param {Object<string, Object<string, number>>} manualByGroup groupId -> { email: 1 }
 * @returns {{groupsMatched:number, groupsUnmatched:number, membersMerged:number}}
 */
function mergeManualGroupMembersIntoDesired_(groups, manualByGroup) {
  const index = {};
  for (const category in groups) {
    for (const groupId in groups[category]) {
      if (!index[groupId]) index[groupId] = [];
      index[groupId].push(category);
    }
  }

  let groupsMatched = 0;
  let groupsUnmatched = 0;
  let membersMerged = 0;

  for (const groupId in manualByGroup) {
    const categories = index[groupId] || [];
    if (!categories.length) {
      groupsUnmatched++;
      continue;
    }
    groupsMatched++;

    for (let i = 0; i < categories.length; i++) {
      const category = categories[i];
      const desired = groups[category][groupId];
      for (const email in manualByGroup[groupId]) {
        if (!desired[email]) {
          desired[email] = 1;
          membersMerged++;
        }
      }
    }
  }

  return { groupsMatched, groupsUnmatched, membersMerged };
}

/**
 * Builds group membership lists based on member attributes
 * Creates wing, group, and (for member-type only) unit-level groups
 * @param {string} groupName - Base name of the group
 * @param {string} attribute - Member attribute to filter by (type, rank, dutyPositionIds, etc.)
 * @param {string} attributeValues - Comma-separated list of values to match
 * @param {Object} members - Members object indexed by CAPID
 * @param {Object} squadrons - Squadrons object indexed by orgid
 * @returns {Object} Groups object with member emails
 */
function getGroupMembers(groupName, attribute, attributeValues, members, squadrons) {
  let groups = {};
  let wingGroupId = CONFIG.WING.toLowerCase() + '.' + groupName;
  let values = (attributeValues == null ? '' : String(attributeValues)).split(',');
  values = values.map(v => (v || '').toString().trim()).filter(v => v.length > 0);
  // True if the member attribute (string or array) matches ANY of the configured values.
  function matchesAny_(memberValue, allowedValues) {
    if (!memberValue) return false;
    if (typeof memberValue === 'string') {
      return allowedValues.indexOf(memberValue) > -1;
    }
    if (Array.isArray(memberValue)) {
      return allowedValues.some(v => memberValue.indexOf(v) > -1);
    }
    return false;
  }

  // Manual/IAOD rows may store multiple duty IDs as a single comma-separated string
  // (optionally with trailing markers like "(A)" / "(P)").
  function normalizeDutyId_(s) {
    return (s || '')
      .toString()
      .replace(/\s*\((a|p)\)\s*$/i, '')
      .trim()
      .toLowerCase();
  }

  function expandDutyIds_(memberValue) {
    const out = [];
    if (!memberValue) return out;

    const pushToken_ = (v) => {
      const parts = (v || '').toString().split(',');
      for (let i = 0; i < parts.length; i++) {
        const norm = normalizeDutyId_(parts[i]);
        if (norm) out.push(norm);
      }
    };

    if (typeof memberValue === 'string') {
      pushToken_(memberValue);
      return out;
    }

    if (Array.isArray(memberValue)) {
      for (let i = 0; i < memberValue.length; i++) {
        pushToken_(memberValue[i]);
      }
      return out;
    }

    return out;
  }

  function matchesAnyDutyId_(memberValue, allowedValues) {
    const memberIds = expandDutyIds_(memberValue);
    if (!memberIds.length) return false;

    const allowed = {};
    for (let i = 0; i < allowedValues.length; i++) {
      const norm = normalizeDutyId_(allowedValues[i]);
      if (norm) allowed[norm] = 1;
    }

    for (let i = 0; i < memberIds.length; i++) {
      if (allowed[memberIds[i]]) return true;
    }
    return false;
  }

  let charterOrgLookup_ = null;

  function getOrgByCharter_(charter) {
    const normalized = (charter || '').toString().trim().toUpperCase();
    if (!normalized) return null;

    if (!charterOrgLookup_) {
      charterOrgLookup_ = {};
      for (const orgid in squadrons) {
        const org = squadrons[orgid];
        const orgCharter = (org && org.charter ? String(org.charter) : '').trim().toUpperCase();
        if (orgCharter) {
          charterOrgLookup_[orgCharter] = org;
        }
      }
    }

    return charterOrgLookup_[normalized] || null;
  }

  function getDutyAssignmentOrg_(dutyPosition) {
    const value = dutyPosition && dutyPosition.value ? String(dutyPosition.value) : '';
    const match = value.match(/\(([^()]+)\)\s*$/);
    return match ? getOrgByCharter_(match[1]) : null;
  }

  let groupId;
  groups[wingGroupId] = {};

  switch (attribute) {
    case 'type':
    case 'dutyPositionIds':
    case 'rank':
      // Pre-seed parent-group variants for configured group-level lists so they
      // continue to exist even when a parent group currently has zero members.
      // Keep unit-level member-type groups dynamic; only real parent GROUP orgs
      // are pre-created here.
      if (attribute !== 'type') {
        for (const orgid in squadrons) {
          const org = squadrons[orgid];
          if (
            org &&
            String(org.scope || '').toUpperCase() === 'GROUP' &&
            String(org.wing || '').toLowerCase() === CONFIG.WING.toLowerCase() &&
            String(org.unit || '') !== '000' &&
            String(org.unit || '') !== '001'
          ) {
            const seededGroupId = String(org.wing || '').toLowerCase() + String(org.unit || '') + '.' + groupName;
            if (!groups[seededGroupId]) groups[seededGroupId] = {};
          }
        }
      }

      for (const member in members) {
        const isMatch = (attribute === 'dutyPositionIds')
          ? matchesAnyDutyId_(members[member][attribute], values)
          : matchesAny_(members[member][attribute], values);

        if (
          isMatch &&
          members[member].email
        ) {
          // Wing-level group
          groups[wingGroupId][members[member].email] = 1;

          // Group-level group (only if parent org is a real GROUP)
          const parent = squadrons[members[member].group];
          if (parent && parent.scope === 'GROUP') {
            groupId =
              squadrons[members[member].orgid].wing.toLowerCase() +
              parent.unit +
              '.' +
              groupName;
            if (!groups[groupId]) groups[groupId] = {};
            groups[groupId][members[member].email] = 1;
          }

          // Unit-level groups (only for member-type categories)
          if (attribute === 'type') {
            const org = squadrons[members[member].orgid];
            if (org && org.unit && org.scope === 'UNIT' && org.unit !== '001') {
              const unitGroupId = org.wing.toLowerCase() + org.unit + '.' + groupName;
              if (!groups[unitGroupId]) groups[unitGroupId] = {};
              groups[unitGroupId][members[member].email] = 1;
            }
          }
        }
      }
      break;

    case 'dutyPositionIdsWingHQ':
      // Wing HQ ONLY (unit 001) for the configured duty position titles.
      // Spreadsheet usage:
      //   Category: duty-position
      //   Group Name: <your DL base name>
      //   Attribute: dutyPositionIdsWingHQ
      //   Values: Director of Safety,Safety Officer
      // Result:
      //   Creates ONLY: <wing>.<groupName> with Wing HQ members only
      //   (e.g., hi.all-safety-hq)

      groupId = wingGroupId;
      if (!groups[groupId]) groups[groupId] = {};

      for (const member in members) {
        const org = squadrons[members[member].orgid];
        if (
          org &&
          org.scope === 'WING' &&
          org.unit === '001' &&
          matchesAnyDutyId_(members[member].dutyPositionIds, values) &&
          members[member].email
        ) {
          groups[groupId][members[member].email] = 1;
        }
      }

      // Prevent creating an empty group
      if (Object.keys(groups[groupId]).length === 0) {
        delete groups[groupId];
      }
      break;

    case 'dutyPositionIdsAndLevel':
      // Prevent creation of Wing HQ-level (hi001.* or 000.*) duty lists
      groupId = groupName;

      // Only build duty groups for Group- and Squadron-level orgs (not Wing HQ or placeholders)
      if (!groups[groupId]) groups[groupId] = {};

      for (const member in members) {
        const org = squadrons[members[member].orgid];
        // Only process if org is not Wing HQ or placeholder units
        if (
          org &&
          org.scope !== 'WING' &&
          org.unit !== '000' &&
          org.unit !== '001' &&
          members[member][attribute] &&
          (
            (typeof members[member][attribute] === 'string' &&
              values.indexOf(members[member][attribute]) > -1) ||
            (Array.isArray(members[member][attribute]) &&
              members[member][attribute].indexOf(values[0]) > -1)
          ) &&
          members[member].email
        ) {
          groups[groupId][members[member].email] = 1;
        }
      }
      // If no members were added, remove the empty group (prevents hi001.* creation)
      if (Object.keys(groups[groupId]).length === 0) {
        delete groups[groupId];
      }
      break;

    case 'dutyPositionLevel':
      groupId = wingGroupId;
      if (groupId && !groups[groupId]) {
        groups[groupId] = {};
      }
      for(const member in members) {
        for (let i = 0; i < members[member].dutyPositions.length; i++) {
          if (members[member].dutyPositions[i].level === values[0] && members[member].email) {
            groups[groupId][members[member].email] = 1;
            break;
          }
        }
      }
      break;

    case 'dutyPositionLevelStaff':
      const staffLevel = (values[0] || '').toString().trim().toUpperCase();

      if (staffLevel === 'WING') {
        for (const member in members) {
          if (!members[member].email || !Array.isArray(members[member].dutyPositions)) continue;

          for (let i = 0; i < members[member].dutyPositions.length; i++) {
            const dutyPosition = members[member].dutyPositions[i];
            const dutyOrg = getDutyAssignmentOrg_(dutyPosition);

            if (
              String(dutyPosition.level || '').toUpperCase() === 'WING' &&
              dutyOrg &&
              String(dutyOrg.scope || '').toUpperCase() === 'WING' &&
              String(dutyOrg.unit || '') === '001' &&
              String(dutyOrg.wing || '').toUpperCase() === CONFIG.WING.toUpperCase()
            ) {
              groups[wingGroupId][members[member].email] = 1;
              break;
            }
          }
        }
      } else if (staffLevel === 'GROUP') {
        delete groups[wingGroupId];

        for (const orgid in squadrons) {
          const org = squadrons[orgid];
          if (
            org &&
            String(org.scope || '').toUpperCase() === 'GROUP' &&
            String(org.wing || '').toUpperCase() === CONFIG.WING.toUpperCase() &&
            String(org.unit || '') !== '000' &&
            String(org.unit || '') !== '001'
          ) {
            groupId = String(org.wing || '').toLowerCase() + String(org.unit || '') + '.' + groupName;
            if (!groups[groupId]) groups[groupId] = {};
          }
        }

        for (const member in members) {
          if (!members[member].email || !Array.isArray(members[member].dutyPositions)) continue;

          for (let i = 0; i < members[member].dutyPositions.length; i++) {
            const dutyPosition = members[member].dutyPositions[i];
            const dutyOrg = getDutyAssignmentOrg_(dutyPosition);

            if (
              String(dutyPosition.level || '').toUpperCase() === 'GROUP' &&
              dutyOrg &&
              String(dutyOrg.scope || '').toUpperCase() === 'GROUP' &&
              String(dutyOrg.wing || '').toUpperCase() === CONFIG.WING.toUpperCase() &&
              String(dutyOrg.unit || '') !== '000' &&
              String(dutyOrg.unit || '') !== '001'
            ) {
              groupId = String(dutyOrg.wing || '').toLowerCase() + String(dutyOrg.unit || '') + '.' + groupName;
              if (!groups[groupId]) groups[groupId] = {};
              groups[groupId][members[member].email] = 1;
              break;
            }
          }
        }
      } else {
        delete groups[wingGroupId];
      }
      break;

    case 'achievements':
      let achievements = parseFile('MbrAchievements');
      for(let i = 0; i < achievements.length; i++) {
        if (members[achievements[i][0]] &&
            members[achievements[i][0]].email &&
            values.indexOf(achievements[i][1]) > -1 &&
            ['ACTIVE', 'TRAINING'].indexOf(achievements[i][2]) > -1) {
          groups[wingGroupId][members[achievements[i][0]].email] = 1;
          // Group-level achievement DLs: ONLY when the member's parent org is a real GROUP.
          // Prevents duplicate Wing HQ groups like "hi001.*".
          const parent = members[achievements[i][0]].group ? squadrons[members[achievements[i][0]].group] : null;
          if (parent && parent.scope === 'GROUP' && parent.unit && parent.unit !== '001' && parent.unit !== '000') {
            groupId =
              squadrons[members[achievements[i][0]].orgid].wing.toLowerCase() +
              parent.unit +
              '.' +
              groupName;
            if (!groups[groupId]) {
              groups[groupId] = {};
            }
            groups[groupId][members[achievements[i][0]].email] = 1;
          }
        }
      }
      break;

    case 'contact':
      // Always include ALL cadets (Workspace primary emails) at wing level only
      for (const member in members) {
        const m = members[member];
        if (!m) continue;

        // Exclude Wing HQ (unit 001)
        const org = squadrons[m.orgid];
        if (org && org.scope === 'WING' && org.unit === '001') continue;

        if (m.email && (m.type || '').toString().trim() === 'CADET') {
          groups[wingGroupId][m.email] = 1;
        }
      }
      let contacts = parseFile('MbrContact');
      for (let i = 0; i < contacts.length; i++) {
        if (members[contacts[i][0]] &&
            values.indexOf(contacts[i][1]) > -1 &&
            contacts[i][6] == 'False') {
          // Exclude Wing HQ (unit 001)
          const org = squadrons[members[contacts[i][0]].orgid];
          if (org && org.scope === 'WING' && org.unit === '001') continue;
          let contact = sanitizeEmail(contacts[i][3]);
          if (contact) {
            groups[wingGroupId][contact] = 1;
            groupId = members[contacts[i][0]].group ?
              (squadrons[members[contacts[i][0]].orgid].wing.toLowerCase() +
               squadrons[members[contacts[i][0]].group].unit + '.' + groupName) : '';
            if (groupId) {
              if (!groups[groupId]) {
                groups[groupId] = {};
              }
              groups[groupId][contact] = 1;
            }
          } else {
            Logger.warn('Invalid contact email - skipping', {
              capsn: contacts[i][0],
              rawEmail: contacts[i][3],
              contactType: contacts[i][1]
            });
          }
        }
      }
      break;

    case 'committeeIds':
      for (const member in members) {
        const email = members[member].email;
        if (!email) continue;

        const assigns = members[member].committeeAssignments;
        if (!Array.isArray(assigns) || assigns.length === 0) continue;

        for (let i = 0; i < assigns.length; i++) {
          const a = assigns[i];
          if (!a || !a.name || !a.orgid) continue;
          if (values.indexOf(a.name) === -1) continue;

          const committeeOrg = squadrons[a.orgid];
          if (!committeeOrg) continue;

          // Wing-scoped committee DL
          if (committeeOrg.scope === 'WING') {
            groups[wingGroupId][email] = 1;
            continue;
          }

          // Group-scoped committee DL (no squadron/unit committees by requirement)
          if (committeeOrg.scope === 'GROUP') {
            groupId =
              committeeOrg.wing.toLowerCase() +
              committeeOrg.unit +
              '.' +
              groupName;
            if (!groups[groupId]) groups[groupId] = {};
            groups[groupId][email] = 1;
          }
        }
      }
      break;

    case 'manualOnly':
      // Create exactly the group IDs listed in Values without deriving members
      // from local CAPWATCH data. User Additions can then supply nested external
      // groups or other managed members later in the pipeline.
      delete groups[wingGroupId];

      if (!values.length) {
        groups[wingGroupId] = {};
        break;
      }

      for (let i = 0; i < values.length; i++) {
        let explicitGroupId = String(values[i] || '').trim().toLowerCase();
        if (!explicitGroupId) continue;
        if (explicitGroupId.endsWith(CONFIG.EMAIL_DOMAIN.toLowerCase())) {
          explicitGroupId = explicitGroupId.slice(0, -CONFIG.EMAIL_DOMAIN.length);
        }
        if (!explicitGroupId) continue;
        if (!groups[explicitGroupId]) groups[explicitGroupId] = {};
      }
      break;

    default:
      Logger.warn('Unknown attribute type', {
        attribute: attribute,
        groupName: groupName
      });
  }
  return groups;
}

/**
 * Saves email groups data to file for tracking and debugging
 * @param {Object} emailGroups - Groups object with member emails
 * @returns {void}
 */
function saveEmailGroups(emailGroups) {
  let folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  let files = folder.getFilesByName('EmailGroups.txt');

  if (files.hasNext()) {
    let file = files.next();
    let content = JSON.stringify(emailGroups);
    file.setContent(content);
    Logger.info('Email groups saved', {
      fileName: 'EmailGroups.txt',
      categories: Object.keys(emailGroups).length
    });
  } else {
    Logger.warn('EmailGroups.txt file not found', {
      folderId: CONFIG.CAPWATCH_DATA_FOLDER_ID
    });
  }
}

/**
 * Saves problematic email addresses to spreadsheet for manual review
 * Includes detailed error information, CAPID mapping, and multiple attempts per email
 * @param {Object} errorEmails - Object mapping email addresses to error details
 * @returns {void}
 */
function saveErrorEmails(errorEmails) {
  if (Object.keys(errorEmails).length === 0) {
    Logger.info('No error emails to save');
    return;
  }

  try {
    // Map emails to CAPIDs
    const contacts = parseFile('MbrContact');
    const emailMap = contacts.reduce(function(map, obj) {
      const cleanEmail = (obj[3] || '').trim().toLowerCase();
      if (cleanEmail) {
        map[cleanEmail] = obj[0];
      }
      return map;
    }, {});

    const sheet = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID)
      .getSheetByName('Error Emails');

    // Clear existing data (keep header row)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }

    // Set up headers if not present
    const headers = sheet.getRange(1, 1, 1, 9).getValues()[0];
    if (headers[0] !== 'Email' || headers.length < 9) {
      sheet.getRange(1, 1, 1, 9).setValues([[
        'Email',
        'CAPID',
        'Error Count',
        'Groups Affected',
        'Error Codes',
        'Last Error Message',
        'Categories',
        'First Seen',
        'Last Seen'
      ]]);

      // Format header row
      sheet.getRange(1, 1, 1, 9)
        .setFontWeight('bold')
        .setBackground('#4285f4')
        .setFontColor('#ffffff');
    }

    // Build rows with detailed information
    const values = [];

    for (const email in errorEmails) {
      const errorInfo = errorEmails[email];
      const attempts = errorInfo.attempts || [];

      if (attempts.length === 0) continue;

      // Extract unique values from attempts
      const groups = [...new Set(attempts.map(a => a.group))].join(', ');
      const errorCodes = [...new Set(attempts.map(a => a.errorCode))].join(', ');
      const categories = [...new Set(attempts.map(a => a.category))].join(', ');

      // Get last error message
      const lastAttempt = attempts[attempts.length - 1];
      const lastErrorMessage = lastAttempt.errorMessage || 'Unknown';

      // Get timestamps
      const firstSeen = errorInfo.firstSeen || attempts[0].timestamp || 'Unknown';
      const lastSeen = lastAttempt.timestamp || 'Unknown';

      // Format dates for spreadsheet
      const firstSeenDate = firstSeen !== 'Unknown' ? new Date(firstSeen) : 'Unknown';
      const lastSeenDate = lastSeen !== 'Unknown' ? new Date(lastSeen) : 'Unknown';

      // Look up CAPID
      const capid = emailMap[email.toLowerCase()] || 'Unknown';

      values.push([
        email,
        capid,
        attempts.length,
        groups,
        errorCodes,
        lastErrorMessage,
        categories,
        firstSeenDate,
        lastSeenDate
      ]);
    }

    // Sort by error count (descending) then by email
    values.sort((a, b) => {
      if (b[2] !== a[2]) return b[2] - a[2]; // Sort by error count
      return a[0].localeCompare(b[0]); // Then by email
    });

    // Write to spreadsheet
    if (values.length > 0) {
      sheet.getRange(2, 1, values.length, 9).setValues(values);

      // Format the data
      const dataRange = sheet.getRange(2, 1, values.length, 9);
      dataRange.setVerticalAlignment('top');

      // Format date columns
      if (values.length > 0) {
        sheet.getRange(2, 8, values.length, 2).setNumberFormat('yyyy-mm-dd hh:mm:ss');
      }

      // Add conditional formatting for error count
      const errorCountRange = sheet.getRange(2, 3, values.length, 1);
      const rules = sheet.getConditionalFormatRules();

      // High errors (5+) = Red
      const redRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(5)
        .setBackground('#f4cccc')
        .setRanges([errorCountRange])
        .build();

      // Medium errors (2-4) = Yellow
      const yellowRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(2, 4)
        .setBackground('#fff2cc')
        .setRanges([errorCountRange])
        .build();

      rules.push(redRule);
      rules.push(yellowRule);
      sheet.setConditionalFormatRules(rules);

      // Auto-resize columns
      for (let i = 1; i <= 9; i++) {
        sheet.autoResizeColumn(i);
      }

      Logger.info('Error emails saved to spreadsheet', {
        count: values.length,
        totalAttempts: values.reduce((sum, row) => sum + row[2], 0),
        sheetName: 'Error Emails'
      });
    }

  } catch (e) {
    Logger.error('Failed to save error emails', {
      errorMessage: e.message,
      errorCount: Object.keys(errorEmails).length
    });
  }
}

/**
 * Retrieves current members of a Google Group
 * Creates the group if it doesn't exist
 * @param {string} groupId - Group identifier (without domain)
 * @param {Object} squadrons - Squadrons object indexed by orgid
 * @param {{name?: string, description?: string}} [meta] - Desired metadata for the group
 * @returns {{email:string, role:string}[]} Array of current group members with roles
 */
function getCurrentGroup(groupId, squadrons, meta = {}, allowExternalMembers = false) {
  const email = groupId + CONFIG.EMAIL_DOMAIN;
  let members = [];
  let nextPageToken = '';

  try {
    do {
      let page = AdminDirectory.Members.list(email, {
        maxResults: GROUP_MEMBER_PAGE_SIZE,
        pageToken: nextPageToken
      });
      if (page.members) {
        members = members.concat(page.members.map(function(member) {
          return {
            email: (member.email || '').toLowerCase(),
            role: (member.role || 'MEMBER').toString().toUpperCase()
          };
        }));
      }
      nextPageToken = page.nextPageToken;
    } while(nextPageToken);

  } catch(e) {
    if (e.details?.code === ERROR_CODES.NOT_FOUND) {
      // Group not found - create it (dry-run aware)
      try {
        // Prefer the already-computed desired metadata so new groups are created
        // with the same final name/description existing groups are patched to.
        let finalName = (meta && meta.name ? meta.name : '').toString().trim();
        let finalDescription = (meta && meta.description ? meta.description : '').toString().trim();

        if (!finalDescription) {
          const org = Object.values(squadrons).find(o => groupId.includes(o.unit));
          const baseName = groupId.split('.').slice(1).join('.');
          const orgName = org ? org.name.toLowerCase().replace(/\b\w/g, c => c.toUpperCase()) : '';
          const formattedGroupName = baseName.replace(/-/g, '.');
          finalDescription = org ? `${orgName} – ${formattedGroupName}` : formattedGroupName;
        }

        if (!finalName) {
          finalName = finalDescription || groupId;
        }

        if (DRY_RUN) {
          Logger.info('💡 [Dry-Run] Would create group', {
            groupId: groupId,
            name: finalName,
            description: finalDescription,
            email: groupId + CONFIG.EMAIL_DOMAIN,
            allowExternalMembers: allowExternalMembers
          });
          return [];
        } else {
          let newGroup = AdminDirectory.Groups.insert({
            email: groupId + CONFIG.EMAIL_DOMAIN,
            name: finalName,
            description: finalDescription
          });

          // Apply "Allow external members" when requested by Groups sheet Add EXT column.
          if (allowExternalMembers) {
            applyAllowExternalMembersSetting_(newGroup.email || (groupId + CONFIG.EMAIL_DOMAIN), true);
          }

          Logger.info('Group created', {
            groupEmail: newGroup.email,
            name: newGroup.name,
            description: newGroup.description,
            allowExternalMembers: allowExternalMembers
          });
        }
      } catch(createError) {
        Logger.error('Failed to create group', {
          groupId: groupId,
          errorMessage: createError.message,
          errorCode: createError.details?.code
        });
      }
    } else {
      Logger.error('Error retrieving group members', {
        groupId: groupId,
        errorMessage: e.message,
        errorCode: e.details?.code
      });
    }
  }

  return members;
}

/**
 * Best-effort apply of the Google Group setting "allowExternalMembers".
 * Requires the Advanced Google Service "Admin SDK Groups Settings API"
 * (AdminGroupsSettings) to be enabled.
 *
 * @param {string} groupEmail
 * @param {boolean} allowExternalMembers
 * @returns {void}
 */
function applyAllowExternalMembersSetting_(groupEmail, allowExternalMembers) {
  try {
    if (typeof AdminGroupsSettings === 'undefined' || !AdminGroupsSettings.Groups || !AdminGroupsSettings.Groups.patch) {
      Logger.warn('AdminGroupsSettings API not available; cannot set allowExternalMembers', {
        group: groupEmail,
        allowExternalMembers: allowExternalMembers
      });
      return;
    }

    executeWithRetry(() =>
      AdminGroupsSettings.Groups.patch({
        allowExternalMembers: allowExternalMembers ? 'true' : 'false'
      }, groupEmail)
    );

    Logger.info('Applied allowExternalMembers setting', {
      group: groupEmail,
      allowExternalMembers: allowExternalMembers
    });
  } catch (e) {
    Logger.warn('Failed to apply allowExternalMembers setting', {
      group: groupEmail,
      allowExternalMembers: allowExternalMembers,
      errorMessage: e.message
    });
  }
}

/**
 * Best-effort apply of managed Google Group settings for UpdateGroups-managed groups.
 * Currently used to keep membership visibility open to the whole domain while
 * preserving the per-group allowExternalMembers behavior from the Groups sheet.
 *
 * @param {string} groupEmail
 * @param {{allowExternalMembers?: boolean, whoCanViewMembership?: string}} settings
 * @returns {void}
 */
function applyManagedGroupSettings_(groupEmail, settings) {
  try {
    const desired = {};

    if (typeof settings.allowExternalMembers === 'boolean') {
      desired.allowExternalMembers = settings.allowExternalMembers ? 'true' : 'false';
    }
    if (settings.whoCanViewMembership) {
      desired.whoCanViewMembership = settings.whoCanViewMembership;
    }
    if (settings.whoCanPostMessage) {
      desired.whoCanPostMessage = settings.whoCanPostMessage;
    }

    if (Object.keys(desired).length === 0) return;

    if (DRY_RUN) {
      Logger.info('💡 [Dry-Run] Would apply managed group settings', {
        group: groupEmail,
        settings: desired
      });
      return;
    }

    if (typeof AdminGroupsSettings === 'undefined' || !AdminGroupsSettings.Groups || !AdminGroupsSettings.Groups.patch) {
      Logger.warn('AdminGroupsSettings API not available; cannot apply managed group settings', {
        group: groupEmail
      });
      return;
    }

    const existing = AdminGroupsSettings.Groups.get(groupEmail);
    const patch = {};

    for (const key in desired) {
      const currentValue = (existing && existing[key] != null) ? String(existing[key]) : '';
      const desiredValue = String(desired[key]);
      if (currentValue !== desiredValue) {
        patch[key] = desiredValue;
      }
    }

    if (Object.keys(patch).length === 0) {
      Logger.info('Managed group settings already correct', {
        group: groupEmail,
        settings: desired
      });
      return;
    }

    executeWithRetry(() => AdminGroupsSettings.Groups.patch(patch, groupEmail));
    Logger.info('Applied managed group settings', {
      group: groupEmail,
      settings: patch
    });
  } catch (e) {
    Logger.warn('Failed to apply managed group settings', {
      group: groupEmail,
      errorMessage: e.message
    });
  }
}

/**
 * Adds additional members to groups based on manual spreadsheet entries
 * Supports MEMBER, MANAGER, and OWNER roles
 * Does not automatically remove members
 * @returns {void}
 */
function updateAdditionalGroupMembers() {
  const start = new Date();
  let additionalMembers = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID)
    .getSheetByName('User Additions')
    .getDataRange()
    .getValues();
  let errorEmails = {};
  const roles = ['MEMBER', 'MANAGER', 'OWNER'];
  let added = 0;
  let skipped = 0;
  let errors = 0;

  for(let i = 1; i < additionalMembers.length; i++) {
    let groups = additionalMembers[i][3].split(',');
    for(let j = 0; j < groups.length; j++) {
      let groupEmail = groups[j].trim() + CONFIG.EMAIL_DOMAIN;
      let email = additionalMembers[i][1];
      let role = additionalMembers[i][2].toLocaleUpperCase();

      if (roles.indexOf(role) < 0) {
        Logger.warn('Invalid role in spreadsheet - skipping', {
          email: email,
          invalidRole: role,
          validRoles: roles.join(', '),
          row: i + 1
        });
        skipped++;
        continue;
      }

      // Add member to group
      try {
        executeWithRetry(() =>
          AdminDirectory.Members.insert({
            email: email,
            role: role
          }, groupEmail)
        );
        Logger.info('Additional member added to group', {
          email: email,
          group: groupEmail,
          role: role
        });
        added++;

      } catch (e) {
        if (e.details?.code === ERROR_CODES.CONFLICT) {
          Logger.info('Member already in group', {
            email: email,
            group: groupEmail,
            role: role
          });
          skipped++;
        } else {
          Logger.error('Failed to add additional member', {
            email: email,
            group: groupEmail,
            role: role,
            row: i + 1,
            errorMessage: e.message,
            errorCode: e.details?.code
          });
          errors++;

          if ([ERROR_CODES.BAD_REQUEST, ERROR_CODES.NOT_FOUND].indexOf(e.details?.code) > -1) {
            // Track detailed error info
            if (!errorEmails[email]) {
              errorEmails[email] = {
                email: email,
                attempts: [],
                firstSeen: new Date().toISOString()
              };
            }
            errorEmails[email].attempts.push({
              group: groups[j].trim(),
              groupEmail: groupEmail,
              category: 'additional-members',
              errorCode: e.details?.code || 'Unknown',
              errorMessage: e.message || 'Unknown error',
              timestamp: new Date().toISOString()
            });
          }
        }
      }
    }
  }

  Logger.info('Additional group members processed', {
    duration: new Date() - start + 'ms',
    added: added,
    skipped: skipped,
    errors: errors,
    errorEmailsCount: Object.keys(errorEmails).length
  });
}

/**
 * Test function for saveErrorEmails
 * @returns {void}
 */
function testSaveErrorEmails() {
  let errorEmails = {
    'bob.rodenhouse@gmail.com': 'test-group-1',
    'mi190.sdavis@live.com': 'test-group-2',
    'michael-shoemaker@sbcglobal.net': 'test-group-3'
  };
  saveErrorEmails(errorEmails);
}

function testEnhancedErrorTracking() {
   // Create test error structure
   const testErrors = {
     'test1@gmail.com': {
       email: 'test1@gmail.com',
       firstSeen: new Date().toISOString(),
       attempts: [
         {
           group: 'test-group-1',
           groupEmail: `test-group-1${CONFIG.EMAIL_DOMAIN}`,
           category: 'test-category',
           errorCode: 404,
           errorMessage: 'Test error message 1',
           timestamp: new Date().toISOString()
         },
         {
           group: 'test-group-2',
           groupEmail: `test-group-2${CONFIG.EMAIL_DOMAIN}`,
           category: 'test-category-2',
           errorCode: 400,
           errorMessage: 'Test error message 2',
           timestamp: new Date().toISOString()
         }
       ]
     },
     'test2@example.com': {
       email: 'test2@example.com',
       firstSeen: new Date().toISOString(),
       attempts: [
         {
           group: 'test-group-3',
           groupEmail: `test-group-3${CONFIG.EMAIL_DOMAIN}`,
           category: 'test-category-3',
           errorCode: 404,
           errorMessage: 'Test error message 3',
           timestamp: new Date().toISOString()
         }
       ]
     }
   };

   saveErrorEmails(testErrors);
   Logger.info('Test completed - check Error Emails sheet');
 }

function debugGroupsTabSource() {
  const spreadsheetId = String(CONFIG.AUTOMATION_SPREADSHEET_ID || '').trim();
  const sheetName = 'Groups';

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Missing sheet: ${sheetName}`);
  }

  const rows = sheet.getDataRange().getValues();
  const header = rows[0] || [];

  Logger.info('Runtime automation source', {
    spreadsheetId: spreadsheetId,
    spreadsheetName: ss.getName(),
    sheetName: sheetName,
    totalRows: rows.length
  });

  Logger.info('Groups header', {
    c1: header[0],
    c2: header[1],
    c3: header[2],
    c4: header[3],
    c5: header[4],
    c6: header[5]
  });

  const matches = [];
  const exactDtyAll = [];
  const staffRows = [];

  for (let i = 1; i < rows.length; i++) {
    const rowNum = i + 1;
    const category = String(rows[i][0] || '').trim();
    const groupName = String(rows[i][1] || '').trim();
    const attribute = String(rows[i][2] || '').trim();
    const values = String(rows[i][3] || '').trim();
    const description = String(rows[i][4] || '').trim();
    const addExt = String(rows[i][5] || '').trim();

    const rowObj = {
      row: rowNum,
      category: category,
      groupName: groupName,
      attribute: attribute,
      values: values,
      description: description,
      addExt: addExt
    };

    if (groupName === 'dty.all') {
      exactDtyAll.push(rowObj);
    }

    if (
      groupName.indexOf('dty.') > -1 ||
      attribute === 'dutyPositionLevelStaff'
    ) {
      matches.push(rowObj);
    }

    if (
      groupName === 'dty.wing-stf-only' ||
      groupName === 'dty.grp-stf-only' ||
      attribute === 'dutyPositionLevelStaff'
    ) {
      staffRows.push(rowObj);
    }
  }

  Logger.info('Exact dty.all rows', {
    count: exactDtyAll.length,
    rows: exactDtyAll
  });

  Logger.info('Staff rows of interest', {
    count: staffRows.length,
    rows: staffRows
  });

  Logger.info('All duty-position style rows', {
    count: matches.length,
    rows: matches
  });
}
