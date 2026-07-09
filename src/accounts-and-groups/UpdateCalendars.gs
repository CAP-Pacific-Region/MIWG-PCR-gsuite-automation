/**
 * -------------------------------------------------------------------------
 * Version: 1.2.3
 * Date: 2026-2-11
 * Author: Lt Col Noel Luneau, Pacific Region
 *
 * Description:
 * - Provides two distinct entry points:
 *   • initializeCalendarsInfrastructure() — one-time / ad-hoc calendar and ACL setup.
 *   • syncMemberCalendarsDaily() — daily member-based calendar assignment.
 * - Introduced SYNC_CALENDAR_USERS configuration control to prevent unwanted
 *   nightly re‑insertion of calendars into user calendar lists.
 * - Standardized calendar naming, metadata updates, and ACL application
 *   ensuring Region, Wing, Group, and Squadron calendars follow consistent
 *   patterns.
 * - Added automatic unit‑level writer group ACLs (###.cal-writers) and
 *   integrated standalone WriterBuilder execution when initializing calendars.
 * - Improved logging for clarity, reproducibility, and troubleshooting across
 *   tenants and automation runs.
 * - Added Director of Administration and Director of Personnel to the writers group
 * - Changed the owner group with Director(s) of IT added to all units
 * - Also added Commander to the Writers group for all units
 * -------------------------------------------------------------------------
 */
const LOG_LEVEL = (typeof CONFIG !== 'undefined' && CONFIG.LOG_LEVEL)
  ? CONFIG.LOG_LEVEL
  : 'INFO';
const SYNC_CALENDAR_USERS =
  (typeof CONFIG !== 'undefined' && typeof CONFIG.SYNC_CALENDAR_USERS !== 'undefined')
    ? CONFIG.SYNC_CALENDAR_USERS
    : false; // default: do NOT touch users' calendar lists on automated runs

function loadPreviousMembersByCapid_() {
  try {
    const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
    const files = folder.getFilesByName('CurrentMembers.txt');

    if (!files.hasNext()) {
      Logger.warn('⚠️ CurrentMembers.txt not found; skipping member-based calendar sync');
      return {};
    }

    const file = files.next();
    const text = file.getBlob().getDataAsString();
    if (!text) return {};

    let data;
    try {
      data = JSON.parse(text);
    } catch (e) {
      Logger.warn('⚠️ Failed to parse CurrentMembers.txt as JSON; skipping member-based calendar sync', {
        errorMessage: e.message
      });
      return {};
    }

    const map = {};

    if (Array.isArray(data)) {
      data.forEach(item => {
        if (!item) return;
        const capid = String(item.CAPID || item.capid || '').trim();
        if (capid) {
          map[capid] = item;
        }
      });
    } else if (typeof data === 'object' && data !== null) {
      Object.keys(data).forEach(key => {
        const item = data[key];
        if (!item) return;
        const capid = String(item.CAPID || item.capid || key || '').trim();
        if (capid) {
          map[capid] = item;
        }
      });
    }

    Logger.info('👥 Loaded previous member state for calendar sync', {
      count: Object.keys(map).length
    });

    return map;
  } catch (e) {
    Logger.warn('⚠️ Error loading previous member state; skipping member-based calendar sync', {
      errorMessage: e.message
    });
    return {};
  }
}

function computeMemberState_() {
  // Load last-saved member state from CurrentMembers.txt
  const previousByCapid = loadPreviousMembersByCapid_();
  if (!Object.keys(previousByCapid).length) {
    Logger.info('ℹ️ No previous member state available; treating as zero new/transferred members');
    return { newMembers: [], transferredMembers: [] };
  }

  // Use the SAME member pipeline used by UpdateMembers
  let membersByCapid = {};
  try {
    membersByCapid = getMembers(CONFIG.MEMBER_TYPES.ACTIVE, false);
    // Apply cadet-lite filtering consistently with UpdateMembers
    if (CONFIG.CADET_LITE === true && Array.isArray(CONFIG.CADET_LITE_EXCLUDED_GRADES)) {
      membersByCapid = Object.fromEntries(
        Object.entries(membersByCapid).filter(([capid, m]) => {
          const rank = (m.rank || '').trim();
          return CONFIG.CADET_LITE_EXCLUDED_GRADES.indexOf(rank) === -1;
        })
      );
    }
  } catch (e) {
    Logger.warn('⚠️ Could not load current members for calendar sync via getMembers()', {
      errorMessage: e.message
    });
    return { newMembers: [], transferredMembers: [] };
  }

  const newMembers = [];
  const transferredMembers = [];

  // Compare CURRENT members (from getMembers) to previous snapshot
  Object.keys(membersByCapid).forEach(capidKey => {
    const m = membersByCapid[capidKey];
    if (!m) return;

    const capid = String(m.capsn || m.CAPID || capidKey || '').trim();
    const orgid = String(m.orgid || m.ORGID || '').trim();
    if (!capid || !orgid) return;

    const prev = previousByCapid[capid];

    if (!prev) {
      newMembers.push({ capid, orgid });
    } else {
      const prevOrgid = String(prev.ORGID || prev.orgid || '').trim();
      if (prevOrgid && prevOrgid !== orgid) {
        transferredMembers.push({
          capid,
          oldOrgid: prevOrgid,
          newOrgid: orgid
        });
      }
    }
  });

  Logger.info('👥 Member state computed for calendar sync (getMembers‑aligned)', {
    newCount: newMembers.length,
    transferredCount: transferredMembers.length
  });

  return { newMembers, transferredMembers };
}
function initializeCalendarsInfrastructure() {
  clearCache();
  const start = new Date();
  Logger.info('📅 Starting Calendar Sync', { started: start.toISOString() });

  const orgs = getSquadrons(); // CAPWATCH org data
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const existingCals = listExistingCalendars_();
  const calendars = {};

  for (const orgid in orgs) {
    const org = orgs[orgid];
    const name = org.name;
    const scope = org.scope?.toUpperCase();
    const unit = (org.unit || '').padStart(3, '0');
    const wing = org.wing?.toUpperCase();

    // Skip inactive or invalid orgs
    if (!name || !scope || (unit === '000' && scope !== 'WING' && scope !== 'REGION')) continue;

    // Calendar naming
    const calName = `${toTitleCase(name)} Calendar`;

    // Base prefix (pcr, hiwg, cawg, nvwg, etc.)
    const base =
      scope === 'REGION'
        ? CONFIG.REGION.toLowerCase()
        : wing.toLowerCase();

    // Units only affect calendar name, not ACLs
    const suffix =
      (['GROUP', 'SQUADRON', 'FLIGHT'].includes(scope) && unit !== '000')
        ? `-${unit}`
        : '';

    // Add these groups to the calendar
    let ownerGroup;
    let readerGroup;

    // Region calendars → pcr.cal-owners / pcr.all (using CONFIG.REGION)
    if (scope === 'REGION') {
      const regionBase = CONFIG.REGION.toLowerCase();
      ownerGroup = `${regionBase}.cal-owners@${CONFIG.DOMAIN}`;
      readerGroup = `${regionBase}.all@${CONFIG.DOMAIN}`;
    }
    // Wing calendars → hiwg.cal-owners / hiwg.all
    else if (scope === 'WING') {
      ownerGroup = `${base}.cal-owners@${CONFIG.DOMAIN}`;
      readerGroup = `${base}wg.all@${CONFIG.DOMAIN}`;
    }
    // Group, Squadron, Flight → Wing-level owners + unit-level readers
    else {
      ownerGroup = `${base}.cal-owners@${CONFIG.DOMAIN}`;
      // Correct unit-level group pattern: or034.all or hi057.all
      const unitId = (unit && unit !== '000')
        ? `${base}${unit}`    // ex: or + 034 → or034
        : `${base}wg`;        // wing-level: hiwg.all

      readerGroup = `${unitId}.all@${CONFIG.DOMAIN}`;
    }

    const calSummary = calName.toLowerCase();

    let calId = existingCals[calSummary];

    try {
      if (calId) {
        Logger.info('✅ Calendar already exists', { calName, calId });
          // Ensure description and timezone are up to date
          try {
            const cal = CalendarApp.getCalendarById(calId);
            // const tz = getTimezoneForWing_(org?.wing || CONFIG.WING);
            cal.setDescription(`${toTitleCase(name)} Calendar`);
            // cal.setTimeZone(tz);

            Logger.info('📝 Updated calendar metadata', {
              calName,
              calId,
              newDescription: `${toTitleCase(name)} Calendar`
              // timeZone: tz
            });
          } catch (metaErr) {
            Logger.warn('⚠️ Could not update calendar metadata', {
              calName,
              errorMessage: metaErr.message
            });
          }
      } else {
        Logger.info('📅 Creating new calendar', { calName, ownerGroup });
        const calendar = CalendarApp.createCalendar(calName, {
          summary: calName,
          description: `${name} Calendar`
          // timeZone: CONFIG.TIMEZONE || 'America/Los_Angeles'
        });
        calId = calendar.getId();
        Logger.info('✅ Calendar created', { calId, calName });
      }

      // Apply ACLs (Group ownership + readers via Google Groups)
      // All of these principals are Google Groups, so use scopeType 'group'
      tryApplyAcl_(calId, ownerGroup, 'owner', 'group');
      tryApplyAcl_(calId, readerGroup, 'reader', 'group');
      tryApplyAcl_(calId, `${base}.all@${CONFIG.DOMAIN}`, 'reader', 'group'); // Wing-wide readers

      // NEW: add unit writer group ACL (hi###.cal-writers)
      try {
        const unitWriterGroup = `${base}${unit}.cal-writers@${CONFIG.DOMAIN}`;
        tryApplyAcl_(calId, unitWriterGroup, 'writer', 'group');
        Logger.info('✏️ Added unit writer ACL', { calId, unitWriterGroup });
      } catch (e) {
        Logger.warn('⚠️ Failed unit writer ACL', { calId, errorMessage: e.message });
      }

      // Make calendar publicly readable (optional)
      try {
        Calendar.Acl.insert({
          role: 'reader',
          scope: { type: 'default' } // "default" = everyone (public)
        }, calId);
        Logger.info('🌍 Calendar made public', { calId, calName });
      } catch (e) {
        if (!/duplicate/i.test(e.message)) {
          Logger.warn('⚠️ Could not make calendar public', { calName, errorMessage: e.message });
        }
      }

      calendars[calSummary] = {
        orgid,
        calId,
        name: calName,
        description: `${toTitleCase(name)} Calendar`,
        scope,
        created: new Date().toISOString()
      };

    } catch (e) {
      Logger.error('❌ Failed to process calendar', {
        orgid, name, scope, errorMessage: e.message, stack: e.stack
      });
    }
  }

  // Save results to Drive
  saveCalendars_(calendars, folder);
  Logger.info('✅ Calendar sync complete', {
    total: Object.keys(calendars).length,
    duration: `${new Date() - start}ms`
  });

  // --- Run standalone unit writer group builder after calendars are created ---
  try {
    WriterBuilder.run();
    Logger.info("📘 Unit writer groups built after calendar sync");
  } catch (e) {
    Logger.warn("⚠️ Unit writer group builder failed", { errorMessage: e.message });
  }
}

/**
 * Daily member-based calendar sync.
 * - Uses existing Calendars.txt metadata (no calendar creation or ACL modification).
 * - Adds calendars for NEW members.
 * - Moves calendars for TRANSFERRED members (remove old org, add new org).
 */
function syncMemberCalendarsDaily() {
  clearCache();
  const start = new Date();
  Logger.info('📅 Starting DAILY member calendar sync', { started: start.toISOString() });

  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const calendars = loadCalendars_(folder);

  if (!calendars || !Object.keys(calendars).length) {
    Logger.warn('⚠️ No calendar metadata found (Calendars.txt missing or empty); skipping member calendar sync');
    return;
  }

  try {
    syncCalendarsForNewAndTransferredUsers_(calendars);
  } catch (e) {
    Logger.warn('⚠️ Member-based calendar user sync failed', {
      errorMessage: e.message
    });
  }

  Logger.info('✅ DAILY member calendar sync complete', {
    duration: `${new Date() - start}ms`
  });
}

function syncCalendarsForNewAndTransferredUsers_(calendars) {
  const { newMembers, transferredMembers } = computeMemberState_();

  // Insert counters and action log
  let addCount = 0;
  let removeCount = 0;
  let actionLog = [];

  if (!newMembers.length && !transferredMembers.length) {
    Logger.info('ℹ️ No new or transferred members detected for calendar user sync');
    return;
  }

  const orgs = getSquadrons();

  const orgCalendars = {};
  Object.keys(calendars || {}).forEach(key => {
    const entry = calendars[key];
    if (!entry || !entry.orgid || !entry.calId) return;
    if (!orgCalendars[entry.orgid]) {
      orgCalendars[entry.orgid] = [];
    }
    orgCalendars[entry.orgid].push(entry.calId);
  });

  let capidToEmail = {};
  try {
    if (WriterBuilder && typeof WriterBuilder.buildCapidToEmailMap === 'function') {
      capidToEmail = WriterBuilder.buildCapidToEmailMap();
    }
  } catch (e) {
    Logger.warn('⚠️ Could not build CAPID → email map for calendar sync', {
      errorMessage: e.message
    });
    return;
  }

  function getOrgChain(orgid) {
    const chain = [];
    let current = orgs[orgid];
    const safetyLimit = 10;
    let steps = 0;

    while (current && steps < safetyLimit) {
      chain.push(current);
      if (!current.nextLevel || (current.scope && String(current.scope).toUpperCase() === 'REGION')) {
        break;
      }
      current = orgs[current.nextLevel];
      steps++;
    }

    return chain;
  }

  // NEW MEMBERS block
  newMembers.forEach(m => {
    const email = capidToEmail[m.capid];
    if (!email) return;

    const chain = getOrgChain(m.orgid);
    chain.forEach(org => {
      if (!org || !org.orgid) return;
      const cals = orgCalendars[org.orgid] || [];
      cals.forEach(calId => {
        try {
          if (!userHasCalendar_(calId, email)) {
            if (!SYNC_CALENDAR_USERS) {
              Logger.info('DRY-RUN: Would ADD calendar for NEW member', {
                capid: m.capid,
                email,
                calId,
                org: org.orgid
              });
              actionLog.push({
                type: 'ADD',
                capid: m.capid,
                email,
                calId,
                org: org.orgid
              });
              addCount++;
            } else {
              addCalendarForUser_(calId, email);
              addCount++;
            }
          }
        } catch (e) {
          Logger.warn('⚠️ Failed to add calendar for new member', {
            capid: m.capid,
            email,
            calId,
            errorMessage: e.message
          });
        }
      });
    });
  });

  // TRANSFERRED MEMBERS block
  transferredMembers.forEach(m => {
    const email = capidToEmail[m.capid];
    if (!email) return;

    const oldChain = getOrgChain(m.oldOrgid);
    const newChain = getOrgChain(m.newOrgid);

    // REMOVAL
    oldChain.forEach(org => {
      if (!org || !org.orgid) return;
      const cals = orgCalendars[org.orgid] || [];
      cals.forEach(calId => {
        try {
          if (!SYNC_CALENDAR_USERS) {
            Logger.info('DRY-RUN: Would REMOVE calendar for TRANSFERRED member', {
              capid: m.capid,
              email,
              calId,
              oldOrg: org.orgid
            });
            actionLog.push({
              type: 'REMOVE',
              capid: m.capid,
              email,
              calId,
              org: org.orgid
            });
            removeCount++;
          } else {
            removeCalendarFromUser_(calId, email);
            removeCount++;
          }
        } catch (e) {
          Logger.warn('⚠️ Failed to remove calendar for transferred member', {
            capid: m.capid,
            email,
            calId,
            errorMessage: e.message
          });
        }
      });
    });

    // ADDITION
    newChain.forEach(org => {
      if (!org || !org.orgid) return;
      const cals = orgCalendars[org.orgid] || [];
      cals.forEach(calId => {
        try {
          if (!userHasCalendar_(calId, email)) {
            if (!SYNC_CALENDAR_USERS) {
              Logger.info('DRY-RUN: Would ADD calendar for TRANSFERRED member', {
                capid: m.capid,
                email,
                calId,
                newOrg: org.orgid
              });
              actionLog.push({
                type: 'ADD',
                capid: m.capid,
                email,
                calId,
                org: org.orgid
              });
              addCount++;
            } else {
              addCalendarForUser_(calId, email);
              addCount++;
            }
          }
        } catch (e) {
          Logger.warn('⚠️ Failed to add calendar for transferred member', {
            capid: m.capid,
            email,
            calId,
            errorMessage: e.message
          });
        }
      });
    });
  });

  // DRY-RUN summary block before final log
  if (!SYNC_CALENDAR_USERS) {
    Logger.info('=== DRY-RUN SUMMARY: MEMBER CALENDAR SYNC ===', {});
    Logger.info('New members detected: ' + newMembers.length, {});
    Logger.info('Transferred members detected: ' + transferredMembers.length, {});
    Logger.info('Calendars that WOULD be added: ' + addCount, {});
    Logger.info('Calendars that WOULD be removed: ' + removeCount, {});
    Logger.info('No changes applied because SYNC_CALENDAR_USERS = false', {});
    Logger.info('Detailed actions:', actionLog);
  }

  Logger.info('✅ Member-based calendar user sync complete', {
    newMembers: newMembers.length,
    transferredMembers: transferredMembers.length
  });
}

/**
 * Returns a map of existing calendars (summary → id)
 */
function listExistingCalendars_() {
  const cals = {};
  const all = CalendarApp.getAllOwnedCalendars();
  for (const cal of all) {
    const name = cal.getName().toLowerCase();
    cals[name] = cal.getId();
  }
  Logger.info('Loaded existing calendars', { count: Object.keys(cals).length });
  return cals;
}

/**
 * Applies ACLs safely with retry/backoff
 * scopeType should be one of: 'user', 'group', 'domain', 'default'
 */
function tryApplyAcl_(calId, scopeValue, role, scopeType = 'user') {
  try {
    let scope;

    if (scopeType === 'default') {
      // "default" scope does not take a value
      scope = { type: 'default' };
    } else if (scopeType === 'domain') {
      scope = { type: 'domain', value: CONFIG.DOMAIN };
    } else {
      // 'user' or 'group' both use the email address as value
      scope = { type: scopeType, value: scopeValue };
    }

    Calendar.Acl.insert({ role, scope }, calId);
    Logger.info('✅ ACL applied', { calId, role, scope });
  } catch (e) {
    if (!/duplicate/i.test(e.message)) {
      Logger.warn('⚠️ Failed to apply ACL', {
        calId,
        role,
        scopeValue,
        errorMessage: e.message
      });
    }
  }
}

/**
 * Removes a calendar from a user's Calendar List.
 */
function removeCalendarFromUser_(calId, userEmail, returnStatus = false) {
  try {
    const scope = 'https://www.googleapis.com/auth/calendar';
    const token = getImpersonatedToken_(userEmail, scope);

    const url = 'https://www.googleapis.com/calendar/v3/users/me/calendarList/' + encodeURIComponent(calId);

    const resp = UrlFetchApp.fetch(url, {
      method: 'delete',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    const code = resp.getResponseCode();
    if (code >= 200 && code < 300) {
      Logger.info('📤 Calendar removed from user list', { user: userEmail, calId });
      if (returnStatus) return code;
    } else if (code !== 404) {
      if (LOG_LEVEL !== 'ERROR' && LOG_LEVEL !== 'NONE') {
        Logger.warn('⚠️ Could not remove calendar from user', {
          user: userEmail,
          code: code,
          response: resp.getContentText()
        });
      }
      if (returnStatus) return code;
    } else {
      if (returnStatus) return 404;
    }
  } catch (err) {
    if (LOG_LEVEL !== 'ERROR' && LOG_LEVEL !== 'NONE') {
      Logger.warn('⚠️ Could not remove calendar from user', { user: userEmail, errorMessage: err.message });
    }
    if (returnStatus) return null;
  }
}

/**
 * Adds a calendar to a user's Calendar List.
 */
function addCalendarForUser_(calId, userEmail) {
  try {
    const scope = 'https://www.googleapis.com/auth/calendar';
    const token = getImpersonatedToken_(userEmail, scope);

    const url = 'https://www.googleapis.com/calendar/v3/users/me/calendarList';
    const payload = JSON.stringify({ id: calId });

    const resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: payload,
      muteHttpExceptions: true
    });

    const code = resp.getResponseCode();
    if (code >= 200 && code < 300) {
      Logger.info('📥 Calendar added to user list', { user: userEmail, calId });
    } else if (code === 409) {
      Logger.info('Calendar already in user list', { user: userEmail });
    } else {
      Logger.warn('⚠️ Could not add calendar to user', {
        user: userEmail,
        code: code,
        response: resp.getContentText()
      });
    }
  } catch (err) {
    Logger.warn('⚠️ Could not add calendar to user', { user: userEmail, errorMessage: err.message });
  }
}

/**
 * Checks whether a user already has a calendar.
 */
function userHasCalendar_(calId, userEmail) {
  try {
    const scope = 'https://www.googleapis.com/auth/calendar.readonly';
    const token = getImpersonatedToken_(userEmail, scope);

    const url = 'https://www.googleapis.com/calendar/v3/users/me/calendarList/' +
      encodeURIComponent(calId);

    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    const code = resp.getResponseCode();
    if (code === 200) return true;
    if (code === 404) return false;

    Logger.warn('⚠️ Unexpected userHasCalendar_ response', {
      user: userEmail,
      calId,
      code,
      body: resp.getContentText()
    });

    return false;

  } catch (err) {
    Logger.warn('⚠️ userHasCalendar_ exception', {
      user: userEmail,
      calId,
      errorMessage: err.message
    });
    return false;
  }
}

/**
 * Synchronizes calendar membership for users in a group.
 * Ensures users in the group have the calendar and removes users no longer in the group.
 */
function syncCalendarUsers_(calId, groupEmail) {
  const cache = CacheService.getScriptCache();
  let cachedGroup = cache.get(`groupMembers_${groupEmail}`);
  let groupMembers;

  if (cachedGroup) {
    groupMembers = JSON.parse(cachedGroup);
  } else {
    groupMembers =
      AdminDirectory.Members.list(groupEmail, { customer: CONFIG.CUSTOMER_ID })
        .members || [];
    cache.put(`groupMembers_${groupEmail}`, JSON.stringify(groupMembers), 60);
  }

  const desiredUsers = groupMembers.map(m => m.email);

  let addedCount = 0;

  desiredUsers.forEach(email => {
    try {
      if (!userHasCalendar_(calId, email)) {
        addCalendarForUser_(calId, email);
        addedCount++;
      }
    } catch (e) {
      Logger.warn("⚠️ Failed to add calendar to user", {
        email,
        calId,
        message: e.message
      });
    }
  });

  Logger.info("📊 Calendar sync summary", {
    calendar: calId,
    group: groupEmail,
    added: addedCount
  });
}

/**
 * Saves calendar metadata to Drive (Calendars.txt)
 */
function saveCalendars_(calendars, folder) {
  const fileName = 'Calendars.txt';
  let file;
  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    file = files.next();
    file.setContent(JSON.stringify(calendars, null, 2));
  } else {
    file = folder.createFile(fileName, JSON.stringify(calendars, null, 2), MimeType.PLAIN_TEXT);
  }
  Logger.info('📁 Calendars saved', { fileName, count: Object.keys(calendars).length });
}

/**
 * Loads calendar metadata from Drive (Calendars.txt).
 */
function loadCalendars_(folder) {
  const fileName = 'Calendars.txt';
  const files = folder.getFilesByName(fileName);
  if (!files.hasNext()) {
    Logger.warn('⚠️ Calendars.txt not found when attempting to load calendars');
    return {};
  }

  try {
    const file = files.next();
    const text = file.getBlob().getDataAsString();
    if (!text) {
      Logger.warn('⚠️ Calendars.txt is empty');
      return {};
    }

    const calendars = JSON.parse(text);
    Logger.info('📁 Calendars loaded', {
      fileName,
      count: Object.keys(calendars || {}).length
    });
    return calendars || {};
  } catch (e) {
    Logger.warn('⚠️ Failed to load Calendars.txt', {
      fileName,
      errorMessage: e.message
    });
    return {};
  }
}

/**
 * Collision-proof standalone writer group builder.
 * All symbols live inside the WriterBuilder namespace.
 */
const WriterBuilder = (() => {

  // Unit-level duty positions (applied only to the member's own unit writers group)
  const UNIT_DUTIES = [
    "Commander",
    "Deputy Commander",
    "Deputy Commander for Cadets",
    "Deputy Commander for Seniors",
    "Public Affairs Officer",
    "Information Technologies Officer",
    "Director of IT",
    "Personnel Officer",
    "Director of Administration"
  ];

  // Wing-level duty positions (applied to EVERY unit writers group in that wing)
  const WING_WRITER_DUTIES = [
    "Commander"
  ];

  // Wing-level duty positions (applied to EVERY unit owners group in that wing)
  const WING_OWNER_DUTIES = [
    "Director of IT",
    "Information Technologies Officer"
  ];

  function run() {
    Logger.info("▶ Starting standalone writer group builder");

    const members = loadMembers();
    const duties = loadDuties();
    const capidMap = buildCapidToEmailMap();

    // Build ORGID -> wing map so we can apply wing-level duty holders to every unit in the same wing
    const orgs = getSquadrons(); // CAPWATCH org structure
    const orgidToWing = {};
    Object.keys(orgs || {}).forEach(id => {
      const o = orgs[id];
      if (!o) return;
      const wing = String(o.wing || o.Wing || '').toLowerCase();
      const orgid = String(o.orgid || o.ORGID || id);
      if (wing && orgid) orgidToWing[orgid] = wing;
    });

    // Pre-compute wing -> writer emails for WING_WRITER_DUTIES using the duty Level field (e.g., "WING")
    const wingWriterEmailsByWing = {};
    duties.forEach(d => {
      const level = String(d.Level || '').toUpperCase();
      if (level !== 'WING') return;
      if (!WING_WRITER_DUTIES.includes(d.Duty)) return;

      const wing = orgidToWing[String(d.ORGID)] || '';
      if (!wing) return;

      const email = capidMap[d.CAPID];
      if (!email) return;

      if (!wingWriterEmailsByWing[wing]) wingWriterEmailsByWing[wing] = [];
      wingWriterEmailsByWing[wing].push(email);
    });

    // De-duplicate per wing
    Object.keys(wingWriterEmailsByWing).forEach(wing => {
      wingWriterEmailsByWing[wing] = [...new Set(wingWriterEmailsByWing[wing])];
    });

    // Pre-compute wing -> owner emails for WING_OWNER_DUTIES using the duty Level field (e.g., "WING")
    const wingOwnerEmailsByWing = {};
    duties.forEach(d => {
      const level = String(d.Level || '').toUpperCase();
      if (level !== 'WING') return;
      if (!WING_OWNER_DUTIES.includes(d.Duty)) return;

      const wing = orgidToWing[String(d.ORGID)] || '';
      if (!wing) return;

      const email = capidMap[d.CAPID];
      if (!email) return;

      if (!wingOwnerEmailsByWing[wing]) wingOwnerEmailsByWing[wing] = [];
      wingOwnerEmailsByWing[wing].push(email);
    });

    // De-duplicate per wing
    Object.keys(wingOwnerEmailsByWing).forEach(wing => {
      wingOwnerEmailsByWing[wing] = [...new Set(wingOwnerEmailsByWing[wing])];
    });

    const orgMap = {};

    members.forEach(m => {
      if (!orgMap[m.ORGID]) orgMap[m.ORGID] = [];
      orgMap[m.ORGID].push(m);
    });

    for (const orgid in orgMap) {
      const unitMembers = orgMap[orgid];
      const first = unitMembers[0];

      const wing = (first.Wing || "").toLowerCase();
      const unit = (first.Unit || "").padStart(3, "0");

      if (!wing || !unit || unit === "000") {
        Logger.info(`Skipping ORGID ${orgid} (invalid unit)`);
        continue;
      }

      const groupEmail = `${wing}${unit}.cal-writers@${CONFIG.DOMAIN}`;

      ensureGroup(groupEmail, `Unit ${unit} Calendar Writers`);

      const unitCapids = new Set(unitMembers.map(m => m.CAPID));

      const unitDuties = duties.filter(d =>
        d.ORGID === orgid &&
        UNIT_DUTIES.includes(d.Duty)
      );

      const unitWriterEmails = unitDuties
        .map(d => capidMap[d.CAPID])
        .filter(Boolean);

      const wingWriterEmails = wingWriterEmailsByWing[wing] || [];

      const writerEmails = [...new Set([
        ...unitWriterEmails,
        ...wingWriterEmails
      ])];

      syncMembers(groupEmail, writerEmails);

      // Create / ensure unit owners group and add wing-level IT owners (add-only, does not remove existing owners)
      const ownersGroupEmail = `${wing}${unit}.cal-owners@${CONFIG.DOMAIN}`;
      ensureGroup(ownersGroupEmail, `Unit ${unit} Calendar Owners`);

      const ownerEmails = wingOwnerEmailsByWing[wing] || [];
      addMembers(ownersGroupEmail, ownerEmails);

      Logger.info(
        `📘 ORGID ${orgid} → ${groupEmail} populated with ${writerEmails.length} writers`
      );
    }

    Logger.info("✔ Complete.");
  }

  /* --------------------------------------------------------------------- */
  /* ----------------------------- LOADERS -------------------------------- */
  /* --------------------------------------------------------------------- */

  function loadMembers() {
    const file = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID)
      .getFilesByName("Member.txt")
      .next();

    const rows = Utilities.parseCsv(file.getBlob().getDataAsString());
    const header = rows.shift();
    const idx = {};
    header.forEach((h, i) => (idx[h] = i));

    return rows.map(r => ({
      CAPID: r[idx.CAPID],
      ORGID: r[idx.ORGID],
      Wing: r[idx.Wing],
      Unit: r[idx.Unit]
    }));
  }

  function loadDuties() {
    const file = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID)
      .getFilesByName("DutyPosition.txt")
      .next();

    const rows = Utilities.parseCsv(file.getBlob().getDataAsString());
    const header = rows.shift();
    const idx = {};
    header.forEach((h, i) => (idx[h] = i));

    // Detect duty level column dynamically (CAPWATCH header names can vary)
    const levelKey = header.find(h => {
      const k = String(h || '').toLowerCase().trim();
      return k === 'level' || k === 'lvl' || k === 'scope' || k === 'type' || k === 'mbrlevel';
    });

    return rows.map(r => ({
      CAPID: r[idx.CAPID],
      ORGID: r[idx.ORGID],
      Duty: r[idx.Duty],
      Level: levelKey ? r[idx[levelKey]] : ''
    }));
  }

  function buildCapidToEmailMap() {
    const map = {};
    let token = null;

    do {
      const page = AdminDirectory.Users.list({
        customer: "my_customer",
        maxResults: 500,
        pageToken: token,
        fields: "users(primaryEmail,externalIds),nextPageToken"
      });

      if (page.users) {
        page.users.forEach(u => {
          (u.externalIds || []).forEach(x => {
            if (x.value) map[x.value] = u.primaryEmail; // CAPID → primaryEmail
          });
        });
      }

      token = page.nextPageToken;
    } while (token);

    return map;
  }

  /* --------------------------------------------------------------------- */
  /* ------------------------- GROUP MGMT -------------------------------- */
  /* --------------------------------------------------------------------- */

  function ensureGroup(email, name) {
    let exists = false;

    // First check if the group already exists
    try {
      AdminDirectory.Groups.get(email, { customer: CONFIG.CUSTOMER_ID });
      exists = true;
    } catch (e) {
      // group does not exist
    }

    if (!exists) {
      Logger.info(`📌 Creating group ${email}`);
      const groupName =
        (name && String(name).trim())
          ? String(name).trim()
          : `${email.substring(0, 5).toLowerCase()} Calendar Group`;
      AdminDirectory.Groups.insert(
        { email, name: groupName },
        { customer: CONFIG.CUSTOMER_ID }
      );
    }

    // ----------------------------------------
    // Wait for Google Directory to become consistent
    // Retry AdminDirectory.Groups.get up to 10 times
    // ----------------------------------------
    let attempts = 0;
    const maxAttempts = 10;
    const baseDelay = 500;

    while (attempts < maxAttempts) {
      try {
        AdminDirectory.Groups.get(email, { customer: CONFIG.CUSTOMER_ID });
        return;  // success
      } catch (e) {
        Utilities.sleep(baseDelay * Math.pow(2, attempts)); // exponential backoff
        attempts++;
      }
    }

    Logger.warn(`⚠️ Group ${email} still not visible after retries`);
  }

  function syncMembers(groupEmail, desiredEmails) {
    let current = [];

    try {
      const list = AdminDirectory.Members.list(groupEmail, { customer: CONFIG.CUSTOMER_ID }).members || [];
      current = list.map(m => m.email);
    } catch (e) {
      Logger.info(`⚠️ Could not read members of ${groupEmail}`);
    }

    const addList = desiredEmails.filter(e => !current.includes(e));
    const removeList = current.filter(e => !desiredEmails.includes(e));

    addList.forEach(email => {
      try {
        AdminDirectory.Members.insert({ email }, groupEmail, { customer: CONFIG.CUSTOMER_ID });
        Logger.info(`➕ Added ${email}`);
      } catch (e) {
        Logger.info(`⚠️ Add failed for ${email}: ${e.message}`);
      }
    });

    removeList.forEach(email => {
      try {
        AdminDirectory.Members.remove(groupEmail, email, { customer: CONFIG.CUSTOMER_ID });
        Logger.info(`➖ Removed ${email}`);
      } catch (e) {
        Logger.info(`⚠️ Remove failed for ${email}: ${e.message}`);
      }
    });

    Logger.info(`📊 ${groupEmail} final count: ${desiredEmails.length}`);
  }

  // Add-only membership helper (does not remove existing members)
  function addMembers(groupEmail, emailsToEnsure) {
    if (!emailsToEnsure || !emailsToEnsure.length) {
      Logger.info(`📊 ${groupEmail} add-only ensured: 0`);
      return;
    }

    let current = [];
    try {
      const list = AdminDirectory.Members.list(groupEmail, { customer: CONFIG.CUSTOMER_ID }).members || [];
      current = list.map(m => m.email);
    } catch (e) {
      Logger.info(`⚠️ Could not read members of ${groupEmail}`);
    }

    const addList = emailsToEnsure.filter(e => e && !current.includes(e));

    addList.forEach(email => {
      try {
        AdminDirectory.Members.insert({ email }, groupEmail, { customer: CONFIG.CUSTOMER_ID });
        Logger.info(`➕ Added ${email} to ${groupEmail}`);
      } catch (e) {
        Logger.info(`⚠️ Add failed for ${email} in ${groupEmail}: ${e.message}`);
      }
    });

    Logger.info(`📊 ${groupEmail} add-only ensured: ${addList.length}`);
  }

  return { run, loadMembers, buildCapidToEmailMap };

})();

/**
 * Call this to run the standalone builder:
 * WriterBuilder.run();
 */

function runWriterGroups() {
  WriterBuilder.run();
}
