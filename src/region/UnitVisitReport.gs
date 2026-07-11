/*******************************************************
 * PCR / Region Unit Visit Tracking Report Generator
 * Source: CAPWATCH TXT exports in Drive folder
 * Output: One sheet tab per Wing, formatted to match
 *
 * Version: 1.0.1
 * Date: 2026-07-09
 * Changes: (1.0.1) Reset frozen rows/columns in buildWingTab_ before the title
 *   merge — clear()/clearFormats() don't reset freeze, so a leftover frozen
 *   column broke the merge. (1.0.0) Folded into the shared src/ from the Pacific
 *   Region project; gated behind PROFILE_.RUN_UNIT_VISIT_REPORT; destination
 *   spreadsheet/calendar IDs read from Script Properties. See PCR_CHANGELOG.md.
 *******************************************************/

/*************************************************************
 * UNIT VISIT REPORT — LOCAL CONFIG
 * -----------------------------------------------------------
 * This config is intentionally local.
 * CAPWATCH filenames and folder contents are assumed stable.
 *************************************************************/

const UNIT_VISIT_CONFIG = {
  // Per-tenant values come from Script Properties — no tenant literals in shared code.
  REGION_CODE: (PropertiesService.getScriptProperties().getProperty('TENANT_REGION') || 'PCR').trim(),

  // If blank, uses the active spreadsheet
  SPREADSHEET_ID: (PropertiesService.getScriptProperties().getProperty('TENANT_UNIT_VISIT_SPREADSHEET_ID') || '').trim(),

  // true = wipe formatting and rebuild tabs each run
  REBUILD_TABS: true,

  // Build a PCR member roster tab (PCR-attached and IAOD-to-PCR)
  PCR_MEMBER_ROSTER: {
    ENABLED: true,
    TAB_NAME: 'PCR Members',
    REBUILD_TAB: true
  },

  TLC_DEBUG: {
    ENABLED: false,
    TAB_NAME: 'TLC Debug',
    ORGIDS: [] // e.g. ['527'] to debug one unit ORGID
  },

  // Optional: Calendar events yearly export tab
  CALENDAR_EVENTS_TAB: {
    ENABLED: true,
    TAB_NAME: 'PCR Calendar',

    // If blank, uses the script owner default calendar
    CALENDAR_ID: (PropertiesService.getScriptProperties().getProperty('TENANT_UNIT_VISIT_CALENDAR_ID') || '').trim(),

    // If omitted, uses the current year
    YEAR: ''
  },

  FORMAT: {
    COL_WIDTHS: {
      A: 7.19,
      B: 29.90,
      C: 19.33,
      D: 24.33,
      E: 37.19,
      F: 41.14,
      G: 18.33,
      H: 18.33,
      I: 18.33
    },

    TITLE_ROW_HEIGHT: 36,
    STAMP_ROW_HEIGHT: 18,
    SPACER_ROW_HEIGHT: 13,
    GROUP_ROW_HEIGHT: 36,
    HEADER_ROW_HEIGHT: 28,
    DATA_ROW_HEIGHT: 16,

    COLORS: {
      GROUP_YELLOW: '#FFD691',
      HEADER_BLUE: '#ADD8E6',
      DATA_GRAY: '#D3D3D3'
    },

    FONTS: {
      TITLE: { family: 'Segoe UI', size: 28, bold: true },
      GROUP: { family: 'Segoe UI', size: 14, bold: true },
      HEADER: { family: 'Segoe UI', size: 10, bold: true },
      DATA: { family: 'Arial', size: 10, bold: false },
      DATA_BOLD: { family: 'Arial', size: 10, bold: true }
    }
  }
};

// Shared utility functions
const norm_ = (v) => String(v || '').trim();
const upper_ = (v) => norm_(v).toUpperCase();
const isPrimary_ = (p) => p === 'PRIMARY' || p === 'Y' || p === 'YES' || p === 'TRUE';

function getUnitVisitRegionCode_() {
  return String(UNIT_VISIT_CONFIG.REGION_CODE || 'PCR').trim();
}

function getUnitVisitSpreadsheet_() {
  const id = String(UNIT_VISIT_CONFIG.SPREADSHEET_ID || '').trim();
  if (id) return SpreadsheetApp.openById(id);

  // If this is a container bound script, ActiveSpreadsheet will be available.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;

  // Standalone scripts have no active spreadsheet.
  throw new Error(
    'No active spreadsheet. Set UNIT_VISIT_CONFIG.SPREADSHEET_ID to the destination Google Sheet ID, or run this script as a container bound script from within the target spreadsheet.'
  );
}

function removeDefaultSheet_(ss) {
  const sheets = ss.getSheets();
  const toDelete = sheets.filter(sh => /^Sheet\d*$/i.test(sh.getName()));

  // Never delete the last remaining sheet
  if (!toDelete.length) return;

  const remaining = sheets.length - toDelete.length;
  if (remaining < 1) return;

  toDelete.forEach(sh => ss.deleteSheet(sh));
}

/**
 * CAPWATCH data folder resolver
 * Uses the shared CONFIG.CAPWATCH_DATA_FOLDER_ID from config.gs.
 */
function getCapwatchDataFolder() {
  if (typeof CONFIG === 'undefined' || !CONFIG || !CONFIG.CAPWATCH_DATA_FOLDER_ID) {
    throw new Error('CONFIG.CAPWATCH_DATA_FOLDER_ID is not set (required to locate CAPWATCH TXT files)');
  }
  return DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
}

/**
 * Entry point
 */
function buildRegionUnitVisitReport() {
  if (!PROFILE_.RUN_UNIT_VISIT_REPORT) {
    Logger.info('buildRegionUnitVisitReport skipped (RUN_UNIT_VISIT_REPORT=false for this tenant profile)');
    return;
  }
  const ss = getUnitVisitSpreadsheet_();

  // NEW
  removeDefaultSheet_(ss);

  const capwatch = precomputeOrgTextCaches_(loadCapwatch_());

  // Build Wing list from Organization rows in REGION
  const REGION_CODE = String(UNIT_VISIT_CONFIG.REGION_CODE || 'PCR').trim();

  const wings = Array.from(
    new Set(
      capwatch.org
        .filter(r => (r.Region || '').trim() === REGION_CODE && (r.Wing || '').trim())
        .map(r => (r.Wing || '').trim())
    )
  ).sort();

  // Build each Wing tab
  wings.forEach(wing => {
    buildWingTab_(ss, wing, capwatch);
  });

  // Optional: PCR member roster tab
  if (UNIT_VISIT_CONFIG.PCR_MEMBER_ROSTER && UNIT_VISIT_CONFIG.PCR_MEMBER_ROSTER.ENABLED) {
    buildPcrMemberRosterTab_(ss, capwatch);
  }

  // Optional: Calendar events tab
  if (UNIT_VISIT_CONFIG.CALENDAR_EVENTS_TAB && UNIT_VISIT_CONFIG.CALENDAR_EVENTS_TAB.ENABLED) {
    buildCalendarEventsTab_(ss);
  }

  // Optional: TLC debug tab
  if (UNIT_VISIT_CONFIG.TLC_DEBUG && UNIT_VISIT_CONFIG.TLC_DEBUG.ENABLED) {
    buildTlcDebugTab_(ss, capwatch);
  }

  SpreadsheetApp.flush();
}

/**
 * Builds a tab that lists Calendar events for the entire year (date, event name, location).
 * Uses CalendarApp and writes a simple 3 column table.
 */
function buildCalendarEventsTab_(ss) {
  const cfg = UNIT_VISIT_CONFIG.CALENDAR_EVENTS_TAB || {};
  const tabName = String(cfg.TAB_NAME || 'Calendar Events');

  const now = new Date();
  const year = String(cfg.YEAR || '').trim() ? parseInt(String(cfg.YEAR).trim(), 10) : now.getFullYear();
  if (!year || isNaN(year)) throw new Error('CALENDAR_EVENTS_TAB.YEAR is not a valid year');

  const start = new Date(year, 0, 1);
  const end = new Date(year + 1, 0, 1);

  let sh = ss.getSheetByName(tabName);
  if (!sh) sh = ss.insertSheet(tabName);

  // Keep formatting consistent with other rebuild behaviors
  sh.clear({ contentsOnly: true });
  sh.clearFormats();

  const calendarId = String(cfg.CALENDAR_ID || '').trim();
  const cal = calendarId ? CalendarApp.getCalendarById(calendarId) : CalendarApp.getDefaultCalendar();
  if (!cal) throw new Error('Unable to open calendar. Check CALENDAR_EVENTS_TAB.CALENDAR_ID');

  const tz = Session.getScriptTimeZone();
  const events = cal.getEvents(start, end);

  // Header
  const header = ['Date', 'Event', 'Location'];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  // Keep Date column as plain text and left aligned (prevents Sheets auto-parsing to YYYY-MM-DD)
  sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@');
  sh.getRange(1, 1, sh.getMaxRows(), 1).setHorizontalAlignment('left');

  // Build sortable items
  const items = events.map(e => {
    const title = e.getTitle() || '';
    const location = e.getLocation() || '';
    const startTime = e.getStartTime();
    const endTime = e.getEndTime();
    const isAllDay = e.isAllDayEvent();

    // For all-day events, CalendarApp endTime is typically the next day at 00:00.
    // Convert to an inclusive end date for display.
    let endDisplay = endTime;
    if (isAllDay && endTime instanceof Date) {
      endDisplay = new Date(endTime.getTime() - 24 * 60 * 60 * 1000);
    }

    return { title, location, startTime, endDisplay, isAllDay };
  });

  // Sort by start
  items.sort((a, b) => a.startTime.getTime() - b.startTime.getTime());

  // Build output rows, inserting month headers like "JAN 2026"
  const rows = [];
  let currentMonthKey = '';

  const monthKeyOf = (d) => Utilities.formatDate(d, tz, 'yyyy-MM');
  const monthLabelOf = (d) => Utilities.formatDate(d, tz, 'MMM yyyy').toUpperCase();

  const dayMon = (d) => Utilities.formatDate(d, tz, 'd MMM');
  const mon = (d) => Utilities.formatDate(d, tz, 'MMM');
  const yyyy = (d) => Utilities.formatDate(d, tz, 'yyyy');

  const dateRangeStr = (start, end) => {
    // Same calendar day
    const sKey = Utilities.formatDate(start, tz, 'yyyy-MM-dd');
    const eKey = Utilities.formatDate(end, tz, 'yyyy-MM-dd');
    if (sKey === eKey) return dayMon(start);

    const sMon = mon(start);
    const eMon = mon(end);
    const sYear = yyyy(start);
    const eYear = yyyy(end);

    // Same month and year: 13-16 Feb
    if (sMon === eMon && sYear === eYear) {
      const sDay = Utilities.formatDate(start, tz, 'd');
      const eDay = Utilities.formatDate(end, tz, 'd');
      return `${sDay}-${eDay} ${sMon}`;
    }

    // Different months (or years): 28 Feb-2 Mar
    const sPart = dayMon(start);
    const ePart = dayMon(end);
    return `${sPart}-${ePart}`;
  };

  items.forEach(it => {
    const start = it.startTime;
    const end = it.endDisplay instanceof Date ? it.endDisplay : it.startTime;

    const mk = monthKeyOf(start);
    if (mk !== currentMonthKey) {
      currentMonthKey = mk;
      rows.push([monthLabelOf(start), '', '']);
    }

    const dateStr = dateRangeStr(start, end);
    rows.push([dateStr, it.title, it.location]);
  });
  // Ensure Date column values are strings so formatting stays consistent
  for (let i = 0; i < rows.length; i++) {
    rows[i][0] = String(rows[i][0] || '');
  }

  if (rows.length) {
    sh.getRange(2, 1, rows.length, 3).setValues(rows);

    // Format month header rows (rows where Event + Location are blank)
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const isMonthHeader = r[0] && !r[1] && !r[2];
      if (!isMonthHeader) continue;

      const sheetRow = 2 + i;
      const rng = sh.getRange(sheetRow, 1, 1, 3);
      rng.merge();
      rng.setFontWeight('bold');
      rng.setHorizontalAlignment('left');
      rng.setNumberFormat('@');
    }
  }

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 3);
}

function buildTlcDebugTab_(ss, capwatch) {
  const cfg = UNIT_VISIT_CONFIG.TLC_DEBUG || {};
  const tabName = String(cfg.TAB_NAME || 'TLC Debug');
  const orgids = (cfg.ORGIDS || []).map(x => String(x || '').trim()).filter(Boolean);

  let sh = ss.getSheetByName(tabName);
  if (!sh) sh = ss.insertSheet(tabName);

  sh.clear({ contentsOnly: true });
  sh.clearFormats();

  const members = Array.isArray(capwatch.members) ? capwatch.members : [];
  const training = Array.isArray(capwatch.training) ? capwatch.training : [];
  const orgRows = Array.isArray(capwatch.org) ? capwatch.org : [];

  if (!orgids.length) {
    sh.getRange('A1').setValue('TLC Debug is enabled. Set UNIT_VISIT_CONFIG.TLC_DEBUG.ORGIDS to one or more ORGIDs to debug.');
    return;
  }

  if (!members.length || !training.length) {
    sh.getRange('A1').setValue('TLC Debug requires Member.txt and Training.txt in the CAPWATCH folder.');
    sh.getRange('A2').setValue(`Member rows loaded: ${members.length}, Training rows loaded: ${training.length}`);
    return;
  }

  const orgByOrgId = indexBy_(orgRows, r => String(r.ORGID || '').trim());

  // Build latest TLC completion per CAPID (same filter as TLC Phase 1)
  const latestTlcByCapId = new Map();

  const parseMdY = (s) => {
    const t = String(s || '').trim();
    if (!t) return null;
    const m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!m) return null;
    const mm = parseInt(m[1], 10);
    const dd = parseInt(m[2], 10);
    const yy = parseInt(m[3], 10);
    const d = new Date(yy, mm - 1, dd);
    return isNaN(d.getTime()) ? null : d;
  };

  training.forEach(t => {
    const capid = String(t.CAPID || '').trim();
    if (!capid) return;
    const type = String(t.TypeCrs || '').trim();
    if (!type) return;
    if (!/^TLC(\b|\s|$)/i.test(type)) return;

    const d = parseMdY(t.Completed);
    if (!d) return;

    const prev = latestTlcByCapId.get(capid);
    if (!prev || d.getTime() > prev.getTime()) latestTlcByCapId.set(capid, d);
  });

  const cutoff = new Date();
  cutoff.setFullYear(cutoff.getFullYear() - 3);

  const header = ['ORGID', 'Charter', 'Unit Name', 'CAPID', 'Last', 'First', 'Member Type', 'Latest TLC Date', 'Current (3yr)'];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');

  const tz = Session.getScriptTimeZone();

  const rows = [];
  orgids.forEach(orgid => {
    const orgRow = orgByOrgId[String(orgid)] || {};
    const charter = orgRow && orgRow.ORGID ? (buildCharter_(orgRow) || '') : '';
    const unitName = String(orgRow.Name || '').trim();

    members.forEach(m => {
      if (String(m.ORGID || '').trim() !== String(orgid)) return;

      const capid = String(m.CAPID || '').trim();
      if (!capid) return;

      const last = String(m.NameLast || '').trim();
      const first = String(m.NameFirst || '').trim();
      const mType = String(m.Type || '').trim();

      const tlcDate = latestTlcByCapId.get(capid) || null;
      const tlcStr = tlcDate ? Utilities.formatDate(tlcDate, tz, 'M/d/yyyy') : '';
      const isCurrent = tlcDate ? (tlcDate.getTime() >= cutoff.getTime()) : false;

      rows.push([String(orgid), charter, unitName, capid, last, first, mType, tlcStr, isCurrent ? 'Y' : '']);
    });
  });

  if (rows.length) {
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  } else {
    sh.getRange('A2').setValue('No members found for the ORGIDs provided.');
  }

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, header.length);
}

/**
 * Loads needed CAPWATCH tables from TXT files
 */
function loadCapwatch_() {
  const folder = getCapwatchDataFolder();
  const org = readCapwatchTxt_(folder, 'Organization.txt');
  const orgAddr = readCapwatchTxt_(folder, 'OrgAddresses.txt');
  const orgMeet = readCapwatchTxt_(folder, 'OrgMeetings.txt');
  const commanders = readCapwatchTxt_(folder, 'Commanders.txt');

  // Index helpers
  const addrByOrgId = groupBy_(orgAddr, r => norm_(r.ORGID));
  const meetByOrgId = groupBy_(orgMeet, r => norm_(r.ORGID));
  const cmdrByOrgId = indexBy_(commanders, r => norm_(r.ORGID));

  // Optional files for PCR member roster (only load if enabled)
  let members = [], dutyPositions = [], mbrContact = [], mbrContactByCapId = {};
  let training = [];
  let tlcCurrentSeniorCountByOrgId = {};

  // Member + Training are used for TLC Phase 1 compliance and PCR roster.
  // (If these files are missing, TLC columns will be blank.)
  if (UNIT_VISIT_CONFIG.PCR_MEMBER_ROSTER && UNIT_VISIT_CONFIG.PCR_MEMBER_ROSTER.ENABLED) {
    members = readCapwatchTxtIfExists_(folder, 'Member.txt');
    dutyPositions = readCapwatchTxtIfExists_(folder, 'DutyPosition.txt');
    mbrContact = readCapwatchTxtIfExists_(folder, 'MbrContact.txt');
    mbrContactByCapId = groupBy_(mbrContact, r => norm_(r.CAPID || r.CapID || r.capid));
  } else {
    // Still try to load Member.txt for TLC even if PCR roster tab is disabled
    members = readCapwatchTxtIfExists_(folder, 'Member.txt');
  }

  training = readCapwatchTxtIfExists_(folder, 'Training.txt');

  // TLC Phase 1: current = completed within last 3 years
  // Count current TLC-qualified SENIOR members by their primary ORGID
  if (members.length && training.length) {
    const latestTlcByCapId = new Map();

    const parseMdY = (s) => {
      const t = String(s || '').trim();
      if (!t) return null;
      const m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (!m) return null;
      const mm = parseInt(m[1], 10);
      const dd = parseInt(m[2], 10);
      const yy = parseInt(m[3], 10);
      if (!mm || !dd || !yy) return null;
      const d = new Date(yy, mm - 1, dd);
      return isNaN(d.getTime()) ? null : d;
    };

    // Build latest TLC completion per CAPID
    training.forEach(t => {
      const capid = String(t.CAPID || '').trim();
      if (!capid) return;
      const type = String(t.TypeCrs || '').trim();
      if (!type) return;
      if (!/^TLC(\b|\s|$)/i.test(type)) return; // TLC, TLC Basic, TLC Intermediate, etc.

      const d = parseMdY(t.Completed);
      if (!d) return;

      const prev = latestTlcByCapId.get(capid);
      if (!prev || d.getTime() > prev.getTime()) latestTlcByCapId.set(capid, d);
    });

    const cutoff = new Date();
    cutoff.setFullYear(cutoff.getFullYear() - 3);

    // Count current TLC seniors by primary ORGID
    const counts = {};
    members.forEach(m => {
      const capid = String(m.CAPID || '').trim();
      if (!capid) return;

      const mType = String(m.Type || '').trim().toUpperCase();
      if (mType !== 'SENIOR') return;

      const orgid = String(m.ORGID || '').trim();
      if (!orgid) return;

      const tlcDate = latestTlcByCapId.get(capid);
      if (!tlcDate) return;
      if (tlcDate.getTime() < cutoff.getTime()) return;

      counts[orgid] = (counts[orgid] || 0) + 1;
    });

    tlcCurrentSeniorCountByOrgId = counts;
  }

  return { org, addrByOrgId, meetByOrgId, cmdrByOrgId, members, dutyPositions, mbrContact, mbrContactByCapId, training, tlcCurrentSeniorCountByOrgId };
}

function precomputeOrgTextCaches_(capwatch) {
  capwatch = capwatch || {};

  const commanderTextByOrgId = {};
  const addressTextByOrgId = {};
  const meetingTextByOrgId = {};

  const orgRows = Array.isArray(capwatch.org) ? capwatch.org : [];
  orgRows.forEach(r => {
    const orgid = String(r.ORGID || '').trim();
    if (!orgid) return;

    const unitStub = { ORGID: orgid };

    if (typeof commanderTextByOrgId[orgid] === 'undefined') {
      commanderTextByOrgId[orgid] = buildCommanderBlock_(unitStub, capwatch) || '';
    }
    if (typeof addressTextByOrgId[orgid] === 'undefined') {
      addressTextByOrgId[orgid] = buildAddressBlock_(unitStub, capwatch) || '';
    }
    if (typeof meetingTextByOrgId[orgid] === 'undefined') {
      meetingTextByOrgId[orgid] = buildMeetingBlock_(unitStub, capwatch) || '';
    }
  });

  capwatch.commanderTextByOrgId = commanderTextByOrgId;
  capwatch.addressTextByOrgId = addressTextByOrgId;
  capwatch.meetingTextByOrgId = meetingTextByOrgId;

  return capwatch;
}

function readCapwatchTxtIfExists_(folder, filename) {
  const it = folder.getFilesByName(filename);
  if (!it.hasNext()) return [];
  return readCapwatchTxt_(folder, filename);
}

/**
 * Build a roster tab that lists members attached to PCR and members who have IAOD duty assignments to PCR.
 * Requires Member.txt and DutyPosition.txt in the CAPWATCH folder. If missing, creates the tab with a note.
 */
function buildPcrMemberRosterTab_(ss, capwatch) {
  const cfg = UNIT_VISIT_CONFIG.PCR_MEMBER_ROSTER || {};
  const tabName = String(cfg.TAB_NAME || 'PCR Members');
  const rebuild = (typeof cfg.REBUILD_TAB === 'undefined') ? true : !!cfg.REBUILD_TAB;

  let sh = ss.getSheetByName(tabName);
  if (!sh) sh = ss.insertSheet(tabName);

  if (rebuild) {
    sh.clear({ contentsOnly: true });
    sh.clearFormats();
  } else {
    sh.clear({ contentsOnly: true });
  }

  const REGION_CODE = getUnitVisitRegionCode_();
  const PCR_HQ_ORGID = '434'; // PCR-PCR-001

  // If required files are missing, write a helpful message and exit.
  const members = Array.isArray(capwatch.members) ? capwatch.members : [];
  const dutyPositions = Array.isArray(capwatch.dutyPositions) ? capwatch.dutyPositions : [];
  if (!members.length || !dutyPositions.length) {
    sh.getRange('A1').setValue('PCR member roster requires Member.txt and DutyPosition.txt in the CAPWATCH folder.');
    sh.getRange('A2').setValue(`Member rows loaded: ${members.length}, DutyPosition rows loaded: ${dutyPositions.length}`);
    return;
  }

  // All ORGIDs inside the PCR region (used to test member primary org)
  const pcrRegionOrgIds = new Set(
    (capwatch.org || [])
      .filter(r => String(r.Region || '').trim() === REGION_CODE)
      .map(r => String(r.ORGID || '').trim())
      .filter(Boolean)
  );

  // CAPWATCH field names are stable in this project
  const CAPID_FIELD = 'CAPID';
  const TYPE_FIELD = 'Type';
  const PRIORITY_FIELD = 'Priority';
  const VALUE_FIELD = 'Contact';

  // Build contact lookup from MbrContact.txt (Primary email + Mobile)
  const contacts = Array.isArray(capwatch.mbrContact) ? capwatch.mbrContact : [];
  const primaryEmailByCapId = new Map();
  const mobileByCapId = new Map();

  contacts.forEach(r => {
    const capid = norm_(r[CAPID_FIELD]);
    if (!capid) return;

    const t = upper_(r[TYPE_FIELD]);
    const pri = upper_(r[PRIORITY_FIELD]);
    const val = norm_(r[VALUE_FIELD]);
    if (!val) return;

    // Primary email (from CAPWATCH MbrContact): prefer EMAIL + PRIMARY, else first EMAIL
    if (t.includes('EMAIL')) {
      if (isPrimary_(pri)) {
        if (!primaryEmailByCapId.has(capid)) primaryEmailByCapId.set(capid, val);
      } else if (!primaryEmailByCapId.has(capid)) {
        primaryEmailByCapId.set(capid, val);
      }
      return;
    }

    // Mobile: prefer MOBILE/CELL + PRIMARY, else first MOBILE/CELL
    if (t.includes('MOBILE') || t.includes('CELL')) {
      if (isPrimary_(pri)) {
        if (!mobileByCapId.has(capid)) mobileByCapId.set(capid, val);
      } else if (!mobileByCapId.has(capid)) {
        mobileByCapId.set(capid, val);
      }
      return;
    }
  });

  const capidToPrimaryOrgId = new Map();
  members.forEach(m => {
    const capid = String(m.CAPID || '').trim();
    const orgid = String(m.ORGID || '').trim();
    if (capid) capidToPrimaryOrgId.set(capid, orgid);
  });

  // Index Organization.txt by ORGID so we can derive full charter numbers for members
  const orgByOrgId = indexBy_((capwatch.org || []), r => String(r.ORGID || '').trim());

  // PCR Email is currently not provisioned in the PCR tenant.
  // Generate a faux PCR email as: first.last@pcr.cap.gov
  const makeFauxPcrEmail = (first, last) => {
    const f = String(first || '').trim().toLowerCase();
    const l = String(last || '').trim().toLowerCase();
    if (!f || !l) return '';

    const clean = (s) => s
      .normalize('NFKD')
      .replace(/\p{Diacritic}/gu, '')
      .replace(/[^a-z0-9]+/g, '.')
      .replace(/^\.+|\.+$/g, '')
      .replace(/\.+/g, '.');

    const local = `${clean(f)}.${clean(l)}`.replace(/\.+/g, '.').replace(/^\.+|\.+$/g, '');
    if (!local) return '';
    return `${local}@pcr.cap.gov`;
  };

  // (old IAOD builder block removed)

  // Attached PCR: primary org is PCR-PCR-001 (ORGID 434)
  const attachedCapIds = new Set();
  capidToPrimaryOrgId.forEach((orgid, capid) => {
    if (String(orgid || '').trim() === PCR_HQ_ORGID) attachedCapIds.add(capid);
  });

  // Members with any duty position assigned to PCR-PCR-001 (ORGID 434)
  const dutyToPcrHq = new Set();
  dutyPositions.forEach(d => {
    const capid = String(d.CAPID || '').trim();
    const orgid = String(d.ORGID || '').trim();
    if (capid && orgid === PCR_HQ_ORGID) dutyToPcrHq.add(capid);
  });

  // IAOD to PCR HQ: primary org within PCR region, not attached to 434, but has duty to 434
  const iaodCapIds = new Set();
  capidToPrimaryOrgId.forEach((orgid, capid) => {
    const primaryOrgId = String(orgid || '').trim();
    if (!primaryOrgId) return;

    const isInPcr = pcrRegionOrgIds.has(primaryOrgId);
    const isAttached = primaryOrgId === PCR_HQ_ORGID;
    const hasDutyToPcr = dutyToPcrHq.has(capid);

    if (isInPcr && !isAttached && hasDutyToPcr) iaodCapIds.add(capid);
  });

  // Combine: PCR attached OR IAOD
  const rows = [];
  members.forEach(m => {
    const capid = String(m.CAPID || '').trim();
    if (!capid) return;

    const isAttached = attachedCapIds.has(capid);
    const isIaod = iaodCapIds.has(capid);
    if (!isAttached && !isIaod) return;

    const last = String(m.NameLast || '').trim();
    const first = String(m.NameFirst || '').trim();
    const grade = String(m.Grade || m.Rank || '').trim();
    const wing = String(m.Wing || '').trim();

    // Unit column should show full charter number (Region-Wing-Unit) from Organization.txt
    const primaryOrgId = String(capidToPrimaryOrgId.get(capid) || '').trim();
    const orgRow = primaryOrgId ? (orgByOrgId[primaryOrgId] || null) : null;
    const unit = orgRow ? (buildCharter_(orgRow) || '') : '';

    const mbrType = String(m.Type || '').trim();
    const status = String(m.MbrStatus || '').trim();

    // Existing email column becomes PCR tenant email
    const pcrEmail = makeFauxPcrEmail(first, last);
    const primaryEmail = primaryEmailByCapId.get(capid) || '';
    const mobile = formatPhone_(mobileByCapId.get(capid) || '');

    rows.push([
      capid,
      last,
      first,
      grade,
      wing,
      unit,
      pcrEmail,
      primaryEmail,
      mobile,
      isAttached ? 'Y' : '',
      isIaod ? 'Y' : '',
      mbrType,
      status
    ]);
  });

  // Sort: Attached first, then IAOD, then Last/First
  // Column indexes in `rows`:
  // 0 CAPID, 1 Last, 2 First, 3 Grade, 4 Wing, 5 Unit,
  // 6 PCR Email, 7 Primary Email, 8 Mobile, 9 Attached, 10 IAOD, 11 Member Type, 12 Status
  rows.sort((a, b) => {
    const aAttached = a[9] === 'Y' ? 0 : 1;
    const bAttached = b[9] === 'Y' ? 0 : 1;
    if (aAttached !== bAttached) return aAttached - bAttached;

    const aIaod = a[10] === 'Y' ? 0 : 1;
    const bIaod = b[10] === 'Y' ? 0 : 1;
    if (aIaod !== bIaod) return aIaod - bIaod;

    const aLast = String(a[1] || '');
    const bLast = String(b[1] || '');
    const cLast = aLast.localeCompare(bLast);
    if (cLast !== 0) return cLast;

    const aFirst = String(a[2] || '');
    const bFirst = String(b[2] || '');
    return aFirst.localeCompare(bFirst);
  });

  // Write header + rows
  const header = ['CAPID', 'Last', 'First', 'Grade', 'Wing', 'Unit Charter', 'PCR Email', 'Primary Email', 'Mobile', 'Attached PCR', 'IAOD to PCR', 'Member Type', 'Status'];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');

  // Center the Attached / IAOD indicator columns
  sh.getRange(1, 10, 1, 2).setHorizontalAlignment('center'); // J:K headers

  if (rows.length) {
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);

    // Center the Attached / IAOD indicator columns
    sh.getRange(2, 10, rows.length, 2).setHorizontalAlignment('center'); // J:K
  }

  // Basic sheet formatting
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, header.length);
}

/**
 * Creates or rebuilds a Wing tab to match the Excel layout
 */
function buildWingTab_(ss, wing, capwatch) {
  const FORMAT = UNIT_VISIT_CONFIG.FORMAT;
  const REBUILD = !!UNIT_VISIT_CONFIG.REBUILD_TABS;
  const name = wing; // tab name exactly wing code
  // CAPWATCH Wing code for California Wing is typically "CA" (not "CAWG")
  const isCAWG = ['CA', 'CAWG'].includes(String(wing || '').trim());
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (REBUILD) {
    sh.clear({ contentsOnly: true });
    sh.clearFormats();
  } else {
    sh.clear({ contentsOnly: true });
  }

  // clear()/clearFormats() do NOT reset freeze state; a leftover frozen column
  // on a pre-existing tab makes the A1:I1 title merge below fail with
  // "You can't merge frozen and non-frozen columns", so reset it explicitly.
  sh.setFrozenColumns(0);
  sh.setFrozenRows(0);

  // Column widths
  Object.entries(FORMAT.COL_WIDTHS).forEach(([colLetter, width]) => {
    sh.setColumnWidth(letterToCol_(colLetter), Math.round(width * 10)); // approximate mapping
  });

  // Title
  sh.getRange("A1:I1").merge();
  sh.setRowHeight(1, FORMAT.TITLE_ROW_HEIGHT);
  const titleCell = sh.getRange("A1");
  titleCell
    .setValue(`${wing} Unit Visit Tracking Report`)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true)
    .setFontFamily(FORMAT.FONTS.TITLE.family)
    .setFontSize(FORMAT.FONTS.TITLE.size)
    .setFontWeight(FORMAT.FONTS.TITLE.bold ? "bold" : "normal");

  // Timestamp
  sh.getRange("A2:I2").merge();
  sh.setRowHeight(2, FORMAT.STAMP_ROW_HEIGHT);

  const runDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "d MMMM yyyy"
  );

  sh.getRange("A2")
    .setValue(`Current as of ${runDate}`)
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(12)
    .setFontWeight("bold");

  // Spacer
  sh.setRowHeight(3, FORMAT.SPACER_ROW_HEIGHT);

  // Build data for this wing
  const orgRows = capwatch.org
    .filter(r => (r.Region || '').trim() === getUnitVisitRegionCode_() && (r.Wing || '').trim() === wing)
    .map(r => ({
      ORGID: String(r.ORGID || ""),
      Wing: (r.Wing || "").trim(),
      Unit: (r.Unit || "").trim(),
      Type: (r.Type || "").trim(),
      Name: (r.Name || "").trim(),
      Status: (r.Status || "").trim(),
      Charter: buildCharter_(r),
    }));

  // Identify Groups vs Units
  // CAPWATCH Organization.Type often includes "GROUP", "SQUADRON", etc.
  const groups = orgRows
    .filter(r => String(r.Type || '').toUpperCase() === 'GROUP');

  // For CAWG, force numeric group ordering (Group 1..8) even when group names are not numeric.
  // For other wings, name sort is fine.
  const groupsSorted = groups.slice().sort((a, b) => {
    if (isCAWG) {
      const keyFor = (g) => {
        // 1) Explicit number in the name ("Group 1", "GROUP 08", etc.)
        const mName = String(g.Name || '').match(/GROUP\s*(\d{1,2})/i);
        if (mName) {
          const n = parseInt(mName[1], 10);
          if (!isNaN(n) && n > 0) return n;
        }

        // 2) Numeric Unit field (often 100, 200, ... for CA groups; sometimes 1..8)
        const unitRaw = String(g.Unit || '').trim();
        if (unitRaw) {
          const u = parseInt(unitRaw, 10);
          if (!isNaN(u)) {
            if (u >= 100) return Math.floor(u / 100); // 100->1, 800->8
            if (u > 0 && u <= 20) return u;           // already a group number
          }
        }

        // 3) Charter segment (PCR-CA-###) if present
        const mChar = String(g.Charter || '').match(/PCR-CA-(\d{1,3})/i);
        if (mChar) {
          const c = parseInt(mChar[1], 10);
          if (!isNaN(c)) {
            if (c >= 100) return Math.floor(c / 100);
            if (c > 0 && c <= 20) return c;
          }
        }

        // 4) Fallback (push unknowns to the bottom)
        return 999;
      };

      const aKey = keyFor(a);
      const bKey = keyFor(b);
      if (aKey !== bKey) return aKey - bKey;
      return String(a.Name || '').localeCompare(String(b.Name || ''));
    }

    return String(a.Name || '').localeCompare(String(b.Name || ''));
  });

  // Determine whether Organization.txt includes a usable NextLevel (parent ORGID) field
  const hasNextLevel = capwatch.org.some(r => typeof r.NextLevel !== 'undefined' && String(r.NextLevel || '').trim() !== '');

  // Units (non-group organizational entries)
  const units = orgRows.filter(r => String(r.Type || '').toUpperCase() !== 'GROUP' && r.Type);

  let rowCursor = 4;

  if (hasNextLevel && groupsSorted.length) {
    // Build index from raw org to include NextLevel relationship
    const rawByOrgId = indexBy_(capwatch.org, r => String(r.ORGID || ""));

    // Map each unit to its parent GROUP ORGID.
    // CAWG often has intermediate parents, so walk the NextLevel chain to the nearest GROUP.
    const unitsByGroupOrgId = groupBy_(units, u => {
      if (isCAWG) return findParentGroupOrgId_(u.ORGID, rawByOrgId);
      return String((rawByOrgId[u.ORGID] || {}).NextLevel || '').trim();
    });

    groupsSorted.forEach(g => {
      const unitList = (unitsByGroupOrgId[String(g.ORGID || "")] || []).sort((a, b) => a.Name.localeCompare(b.Name));
      rowCursor = writeGroupSection_(sh, rowCursor, g, unitList, capwatch);
    });

    // Catch any units not mapped to a group
    const unmapped = unitsByGroupOrgId[""] ? unitsByGroupOrgId[""] : [];
    if (unmapped.length) {
      rowCursor = writeGroupSection_(
        sh,
        rowCursor,
        { Name: `${wing} Units`, ORGID: "", Charter: "" },
        unmapped.sort((a, b) => a.Name.localeCompare(b.Name)),
        capwatch
      );
    }
  } else {
    // Single section fallback
    rowCursor = writeGroupSection_(
      sh,
      rowCursor,
      { Name: `${wing} Units`, ORGID: "", Charter: "" },
      units.sort((a, b) => a.Name.localeCompare(b.Name)),
      capwatch
    );
  }

  // Keep grid tidy
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
}

/**
 * Writes one Group header + table header + unit rows
 */
function writeGroupSection_(sh, startRow, group, unitList, capwatch) {
  const FORMAT = UNIT_VISIT_CONFIG.FORMAT;
  // Group header row (merge A:I)
  sh.setRowHeight(startRow, FORMAT.GROUP_ROW_HEIGHT);
  sh.getRange(startRow, 1, 1, 9).merge(); // A:I

  const groupText = buildGroupHeaderText_(group, capwatch);
  const groupRange = sh.getRange(startRow, 1); // A
  const groupBarRange = sh.getRange(startRow, 1, 1, 9); // A:I

  // RichText: first line uses GROUP font, subsequent lines use DATA_BOLD font (same as Squadron commander blocks)
  const titleStyle = SpreadsheetApp.newTextStyle()
    .setFontFamily(FORMAT.FONTS.GROUP.family)
    .setFontSize(FORMAT.FONTS.GROUP.size)
    .setBold(true)
    .build();

  const cmdStyle = SpreadsheetApp.newTextStyle()
    .setFontFamily(FORMAT.FONTS.DATA_BOLD.family)
    .setFontSize(FORMAT.FONTS.DATA_BOLD.size)
    .setBold(true)
    .build();

  const firstNl = groupText.indexOf("\n");
  const titleEnd = (firstNl === -1) ? groupText.length : firstNl;

  let rtv = SpreadsheetApp.newRichTextValue().setText(groupText);
  if (groupText.length) {
    rtv = rtv.setTextStyle(0, titleEnd, titleStyle);
    if (titleEnd < groupText.length) {
      rtv = rtv.setTextStyle(Math.min(titleEnd + 1, groupText.length), groupText.length, cmdStyle);
    }
  }

  // Put the RichText in the merged cell (A), but apply the yellow fill across the full merged bar (A:G)
  groupRange.setRichTextValue(rtv.build());
  groupBarRange
    .setBackground(FORMAT.COLORS.GROUP_YELLOW)
    .setVerticalAlignment("middle")
    .setWrap(true);

  // Group header bar spans through Last Visit
  startRow++;

  // Column headers (row)
  sh.setRowHeight(startRow, FORMAT.HEADER_ROW_HEIGHT);

  // Header values
  const headerRow = startRow;
  sh.getRange(headerRow, 1, 1, 2).merge(); // A:B
  sh.getRange(headerRow, 1).setValue('Unit/Group Name');

  // Write C:I headers in one call
  sh.getRange(headerRow, 3, 1, 7).setValues([[
    'Charter',
    'Commander (Name/Appointed/Time in Service/Phone/Primary Email)',
    'Unit Address',
    'Meeting Info',
    'TLC Qualified Seniors',
    'TLC Status',
    'Last Visit'
  ]]);

  // Apply shared header formatting across A:I once
  const headerRange = sh.getRange(headerRow, 1, 1, 9);
  headerRange
    .setBackground(FORMAT.COLORS.HEADER_BLUE)
    .setFontFamily(FORMAT.FONTS.HEADER.family)
    .setFontSize(FORMAT.FONTS.HEADER.size)
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setWrap(true);

  // Header alignments
  sh.getRange(headerRow, 1, 1, 4).setHorizontalAlignment('center'); // A:B + C + D
  sh.getRange(headerRow, 5, 1, 1).setHorizontalAlignment('center'); // E
  sh.getRange(headerRow, 6, 1, 4).setHorizontalAlignment('left');   // F:I

  // Header border once
  headerRange.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

  startRow++;

  // Unit rows
  const dataStartRow = startRow;
  const unitCount = unitList.length;

  if (unitCount > 0) {
    // Set all data row heights at once
    sh.setRowHeights(dataStartRow, unitCount, FORMAT.DATA_ROW_HEIGHT);

    // Build values for C:I in one 2D array, using precomputed caches
    const cmdCache = capwatch.commanderTextByOrgId || {};
    const addrCache = capwatch.addressTextByOrgId || {};
    const meetCache = capwatch.meetingTextByOrgId || {};
    const tlcCounts = capwatch.tlcCurrentSeniorCountByOrgId || {};

    const valuesCtoI = new Array(unitCount);
    const nonCompliantA1s = []; // ranges like "G12:H12"

    for (let i = 0; i < unitCount; i++) {
      const u = unitList[i];
      const orgid = String(u.ORGID || '').trim();

      const tlcCount = orgid ? (tlcCounts[orgid] || 0) : '';
      const tlcStatus = orgid
        ? ((Number(tlcCount || 0) >= 2) ? 'Compliant' : 'Non Compliant')
        : '';

      // Collect non-compliant rows for red fill (all units)
      if (orgid && tlcStatus === 'Non Compliant') {
        const r = dataStartRow + i;              // sheet row
        nonCompliantA1s.push(`G${r}:H${r}`);     // TLC Count + Status cells
      }

      valuesCtoI[i] = [
        u.Charter || '',
        (orgid && cmdCache[orgid]) || '',
        (orgid && addrCache[orgid]) || '',
        (orgid && meetCache[orgid]) || '',
        tlcCount,
        tlcStatus,
        ''
      ];
    }

    // Write C:I values in one call
    sh.getRange(dataStartRow, 3, unitCount, 7).setValues(valuesCtoI);


    // Ensure TLC count is treated as a number and centered
    // TLC Qualified Seniors is column G
    sh.getRange(dataStartRow, 7, unitCount, 1)
      .setNumberFormat('0')
      .setHorizontalAlignment('center');

    // TLC Status is column H
    sh.getRange(dataStartRow, 8, unitCount, 1).setNumberFormat('@');


    // A:B merge across all rows in one call, then write names in a single setValues
    const names = unitList.map(u => [u.Name || '']);
    sh.getRange(dataStartRow, 1, unitCount, 2).mergeAcross(); // merge A:B per row
    sh.getRange(dataStartRow, 1, unitCount, 1).setValues(names);

    // Apply shared data formatting across A:I once
    const dataRange = sh.getRange(dataStartRow, 1, unitCount, 9);
    dataRange
      .setBackground(FORMAT.COLORS.DATA_GRAY)
      .setFontFamily(FORMAT.FONTS.DATA.family)
      .setFontSize(FORMAT.FONTS.DATA.size)
      .setFontWeight('normal')
      .setVerticalAlignment('top')
      .setHorizontalAlignment('left')
      .setWrap(true);

    // Highlight non-compliant TLC rows AFTER base background is applied (otherwise it gets overwritten)
    if (nonCompliantA1s.length) {
      sh.getRangeList(nonCompliantA1s).setBackground('#FF6666');
    }

    // Commander column (D) uses bold weight only; family/size already applied by dataRange
    sh.getRange(dataStartRow, 4, unitCount, 1)
      .setFontWeight('bold');

    // Apply borders once for the whole data block
    dataRange.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

    startRow = dataStartRow + unitCount;
  }

  // Spacer between groups (optional)
  return startRow;
}

/**
 * Build the Charter string similar to PCR-CA-### (best-effort from Wing/Unit fields)
 */
function buildCharter_(orgRow) {
  const wing = (orgRow.Wing || "").trim();
  const unit = (orgRow.Unit || "").trim();
  const region = (orgRow.Region || "").trim();
  if (!wing || !unit || !region) return "";
  return `${region}-${wing}-${unit}`;
}

function getCommanderGrade_(cmd) {
  if (!cmd) return '';
  const pick = (v) => String(v || '').trim();
  return (
    pick(cmd.Grade) ||
    pick(cmd.Rank) ||
    pick(cmd.MbrGrade) ||
    pick(cmd.MemberGrade) ||
    pick(cmd.GradeCode) ||
    ''
  );
}

/**
 * Group header text: "<Group Name> -- (Commander: ... / Phone: ... / Appointed: ...)"
 */
function buildGroupHeaderText_(group, capwatch) {
  const cmd = group.ORGID ? capwatch.cmdrByOrgId[group.ORGID] : null;

  const nameWithCharter = group.Charter ? `${group.Name} (${group.Charter})` : `${group.Name}`;
  if (!cmd) return nameWithCharter;

  const fullName = `${(cmd.NameFirst || "").trim()} ${(cmd.NameLast || "").trim()}`.trim();
  const grade = getCommanderGrade_(cmd);
  const capid = String(cmd.CAPID || cmd.CapID || cmd.capid || '').trim();
  const nameLine = [fullName, grade, capid].filter(Boolean).join(', ');

  const appointedDate = cmd.DateAsg ? new Date(cmd.DateAsg) : null;
  const appointed = appointedDate ? formatDate_(appointedDate) : "";

  let months = "";
  if (appointedDate && !isNaN(appointedDate.getTime())) {
    const now = new Date();
    const m = (now.getFullYear() - appointedDate.getFullYear()) * 12 + (now.getMonth() - appointedDate.getMonth());
    months = String(m);
  }

  const phone = extractPhoneFromCommander_(cmd, capwatch) || "";
  const email = extractPrimaryEmailFromCommander_(cmd, capwatch) || "";

  // Multiline header block
  const lines = [nameWithCharter];
  if (nameLine) lines.push(`Commander: ${nameLine}`);
  if (appointed) lines.push(`Appointed: ${appointed}`);
  if (months) lines.push(`Time in Service: ${months} mths`);
  if (phone) lines.push(phone);
  if (email) lines.push(email);

  return lines.join("\n");
}

function buildCommanderBlock_(unit, capwatch) {
  const cmd = unit.ORGID ? capwatch.cmdrByOrgId[unit.ORGID] : null;
  if (!cmd) return "";

  const name = `${(cmd.NameFirst || "").trim()} ${(cmd.NameLast || "").trim()}`.trim();
  const grade = getCommanderGrade_(cmd);
  const capid = String(cmd.CAPID || cmd.CapID || cmd.capid || '').trim();
  const nameLine = [name, grade, capid].filter(Boolean).join(', ');

  const appointedDate = cmd.DateAsg ? new Date(cmd.DateAsg) : null;

  let months = "";
  if (appointedDate) {
    const now = new Date();
    const m = (now.getFullYear() - appointedDate.getFullYear()) * 12 + (now.getMonth() - appointedDate.getMonth());
    months = String(m);
  }

  const phone = extractPhoneFromCommander_(cmd, capwatch);
  const email = extractPrimaryEmailFromCommander_(cmd, capwatch);

  const lines = [];
  if (nameLine) lines.push(nameLine);
  if (appointedDate) lines.push(formatDate_(appointedDate));
  if (months) lines.push(`${months} mths`);
  if (phone) lines.push(phone);
  if (email) lines.push(email);

  return lines.join("\n");
}

function buildAddressBlock_(unit, capwatch) {
  const rows = unit.ORGID ? (capwatch.addrByOrgId[unit.ORGID] || []) : [];
  if (!rows.length) return "";

  // ONLY use MEETING PRIMARY. If missing, return blank.
  const preferred = rows.find(r => upper_(r.Type) === "MEETING" && isPrimary_(upper_(r.Priority)));
  if (!preferred) return "";

  const parts = [];
  if (preferred.Addr1) parts.push(preferred.Addr1);
  if (preferred.Addr2) parts.push(preferred.Addr2);

  const city = preferred.City || "";
  const state = preferred.State || "";
  const zip = preferred.Zip || "";
  const lastLine = [city, state, zip].filter(Boolean).join(", ").replace(", ,", ",");
  if (lastLine.trim()) parts.push(lastLine);

  return parts.join("\n");
}

function buildMeetingBlock_(unit, capwatch) {
  const rows = unit.ORGID ? (capwatch.meetByOrgId[unit.ORGID] || []) : [];
  if (!rows.length) return "";

  // Your sample shows day/time on first line and frequency on second line
  // CAPWATCH OrgMeetings typically has MeetDay and MeetTime
  // If multiple meeting rows exist, we stack them
  const blocks = rows.map(r => {
    const day = (r.MeetDay || "").trim();
    const time = (r.MeetTime || "").trim();
    const firstLine = [day, time].filter(Boolean).join(" @ ");
    // Some exports use Descr for cadence
    const secondLine = (r.Descr || "").trim();
    return [firstLine, secondLine].filter(Boolean).join("\n");
  });

  return blocks.join("\n");
}

/**
 * Unified contact extraction (phone or email)
 */
function extractContactFromCommander_(cmd, capwatch, mode) {
  const capid = norm_(cmd.CAPID || cmd.CapID || cmd.capid);
  if (!capid) return "";

  const byCapId = capwatch && capwatch.mbrContactByCapId ? capwatch.mbrContactByCapId : null;
  const mine = byCapId ? (byCapId[capid] || []) : [];
  if (!mine.length) return "";

  const typeOf = (row) => upper_(row.Type || row.ContactType || row.MbrContactType || row.Contact || row.Field);
  const priOf = (row) => upper_(row.Priority || row.Pri || row.Primary || row.IsPrimary);

  const pickVal = (row) => {
    const v = row.Value ?? row.Contact ?? row.ContactValue ?? row.Data ?? row.Text ?? row.ContactInfo ?? row.Info;
    if (v !== undefined && v !== null && norm_(v)) return norm_(v);

    if (mode === 'phone') {
      const allVals = Object.values(row).map(norm_).filter(Boolean);
      return allVals.find(x => x.replace(/\D/g, "").length >= 10) || "";
    } else {
      const allVals = Object.values(row).map(norm_).filter(Boolean);
      return allVals.find(x => x.includes("@")) || "";
    }
  };

  if (mode === 'phone') {
    // Prefer: Mobile PRIMARY, then Mobile any, then Phone PRIMARY, then Phone any
    let chosen = null;
    for (const r of mine) {
      const t = typeOf(r);
      const isMobile = t.includes("MOBILE") || t.includes("CELL");
      const isPri = isPrimary_(priOf(r));

      if (isMobile && isPri) { chosen = r; break; }
      if (!chosen && isMobile) chosen = r;
    }
    if (!chosen) {
      for (const r of mine) {
        const t = typeOf(r);
        const isPhone = t.includes("PHONE") || t.includes("HOME") || t.includes("WORK");
        const isPri = isPrimary_(priOf(r));

        if (isPhone && isPri) { chosen = r; break; }
        if (!chosen && isPhone) chosen = r;
      }
    }

    return chosen ? formatPhone_(pickVal(chosen)) : "";
  } else {
    // Email: prefer PRIMARY, else first
    const isEmail = (row) => typeOf(row).includes("EMAIL");
    const emailPrimary = mine.find(r => isEmail(r) && isPrimary_(priOf(r)));
    const emailAny = mine.find(r => isEmail(r));
    const chosen = emailPrimary || emailAny;
    return chosen ? pickVal(chosen) : "";
  }
}

function extractPhoneFromCommander_(cmd, capwatch) {
  return extractContactFromCommander_(cmd, capwatch, 'phone');
}

function extractPrimaryEmailFromCommander_(cmd, capwatch) {
  return extractContactFromCommander_(cmd, capwatch, 'email');
}


function readCapwatchTxt_(folder, filename) {
  const it = folder.getFilesByName(filename);
  if (!it.hasNext()) throw new Error(`Missing CAPWATCH file: ${filename}`);
  const file = it.next();
  const text = file.getBlob().getDataAsString();

  // Parse the entire CSV in one pass (much faster than per-line parseCsv)
  const rows = Utilities.parseCsv(text);
  if (!rows || !rows.length) return [];

  // Header row
  const header = (rows[0] || []).map(h => String(h || '').trim());
  if (!header.length) return [];

  const out = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row.length) continue;

    // Skip completely blank lines
    const joined = row.map(x => String(x || '')).join('').trim();
    if (!joined) continue;

    const obj = {};
    for (let c = 0; c < header.length; c++) {
      const key = header[c];
      if (!key) continue;
      obj[key] = row[c] === undefined ? "" : row[c];
    }
    out.push(obj);
  }

  return out;
}

function groupBy_(arr, keyFn) {
  return arr.reduce((acc, x) => {
    const k = keyFn(x);
    if (!acc[k]) acc[k] = [];
    acc[k].push(x);
    return acc;
  }, {});
}

function indexBy_(arr, keyFn) {
  return arr.reduce((acc, x) => {
    acc[keyFn(x)] = x;
    return acc;
  }, {});
}

/**
 * Walk the Organization hierarchy (NextLevel chain) to find the nearest parent org
 * whose Type is GROUP. Returns the parent GROUP ORGID as a string, or "" if none.
 */
function findParentGroupOrgId_(orgId, rawByOrgId) {
  let current = rawByOrgId[String(orgId || "")] || null;
  let guard = 0;

  while (current && guard < 25) {
    const parentId = String(current.NextLevel || "").trim();
    if (!parentId) return "";

    const parent = rawByOrgId[parentId] || null;
    if (!parent) return "";

    if (String(parent.Type || "").trim().toUpperCase() === "GROUP") return parentId;

    current = parent;
    guard++;
  }

  return "";
}

function formatDate_(d) {
  const dt = (d instanceof Date) ? d : new Date(d);
  return Utilities.formatDate(dt, Session.getScriptTimeZone(), "M/d/yyyy");
}

function formatPhone_(raw) {
  const digits = norm_(raw).replace(/\D/g, '');
  if (!digits) return '';

  // Strip leading country code if present
  const d = (digits.length === 11 && digits.startsWith('1')) ? digits.slice(1) : digits;

  return d.length === 10 ? `${d.slice(0, 3)}-${d.slice(3, 6)}-${d.slice(6)}` : norm_(raw);
}

function letterToCol_(letter) {
  return letter.toUpperCase().charCodeAt(0) - 64;
}
