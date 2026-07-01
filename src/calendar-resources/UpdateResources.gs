/**
 * Resource Management Module
 *
 * Creates and updates Google Workspace calendar resources for CAP aircraft and vehicles.
 *
 * Aircraft (Aircraft.txt):
 *   - Only rows where Tailno starts with 'N' (FAA-registered aircraft)
 *   - Only rows where Status === 'Active'
 *   - Building = ICAO airport identifier (AirPortID)
 *   - Resource name = "<Tailno> - <Model> (<YrManf>)"
 *   - Resource category = Aircraft
 *
 * Vehicles (Vehicles.txt):
 *   - All rows are distinct vehicles (no equipment rows to filter)
 *   - Building = unit's MEETING address from OrgAddresses.txt
 *   - Resource name = "<YrMfgr> <Make> - <VehType> (<VIN>)"
 *   - Resource category = Vehicle
 *
 * Buildings are created automatically for airports and unit addresses.
 * Existing resources and buildings are updated if they already exist.
 *
 * RECOMMENDED SCHEDULE: Run after getCapwatch() completes (daily or as needed)
 *
 * Authors: [Your name here]
 */

// ============================================================================
// AIRCRAFT TYPE CODE → HUMAN-READABLE MODEL NAME
// ============================================================================

/**
 * Maps CAPWATCH ACCode values to human-readable Cessna model names.
 * CAP's fleet is exclusively Cessna aircraft. ACCode roughly follows FAA type
 * designators but may include variant suffixes (e.g. T182T, U206G, C172S).
 *
 * Matching is done by stripping leading letters and trailing letters to find
 * the numeric model core, then falling back to direct and prefix lookups.
 */
const ACCODE_MAP = {
  // Cessna 172 variants
  'C172':   'Cessna 172 Skyhawk',
  'C172R':  'Cessna 172R Skyhawk',
  'C172S':  'Cessna 172S Skyhawk SP',
  '172':    'Cessna 172 Skyhawk',
  '172R':   'Cessna 172R Skyhawk',
  '172S':   'Cessna 172S Skyhawk SP',

  // Cessna 182 variants
  'C182':   'Cessna 182 Skylane',
  'C182T':  'Cessna 182T Skylane',
  'C82R':   'Cessna 182RG Skylane RG',
  '182':    'Cessna 182 Skylane',
  '182T':   'Cessna 182T Skylane',

  // Turbo Cessna 182 variants
  'T182':   'Cessna T182 Turbo Skylane',
  'T182T':  'Cessna T182T Turbo Skylane',

  // Cessna 206 / Stationair variants (T206H is the turbocharged 206H)
  'T206H':  'Cessna T206H Turbo Stationair',
  'C206':   'Cessna 206 Stationair',
  'C206H':  'Cessna 206H Stationair',
  'U206':   'Cessna U206 Stationair',
  'U206G':  'Cessna U206G Stationair',
  'C206U':  'Cessna 206U Stationair',
  'TU206':  'Cessna TU206 Turbo Stationair',
  'TU206G': 'Cessna TU206G Turbo Stationair',
  'P206':   'Cessna P206 Super Skylane',

  // Schweizer gliders (operated by CAP for cadet orientation flights)
  '2-33A':  'Schweizer SGS 2-33A',
  'L23':    'Schweizer SGS 2-33',
};

/**
 * Resolves an ACCode to a human-readable model name.
 * Tries direct lookup first, then uppercase, then strips trailing/leading
 * variant letters to find the best match.
 *
 * @param {string} acCode - Raw ACCode value from Aircraft.txt
 * @returns {string} Human-readable model name, or the original code if unknown
 */
function resolveAircraftModel(acCode) {
  if (!acCode) return 'Unknown Model';

  const code = acCode.trim();

  // Direct lookup
  if (ACCODE_MAP[code]) return ACCODE_MAP[code];

  // Uppercase lookup
  const upper = code.toUpperCase();
  if (ACCODE_MAP[upper]) return ACCODE_MAP[upper];

  // Try progressively stripping trailing characters (e.g. T182T → T182 → 182)
  for (let len = upper.length - 1; len >= 2; len--) {
    const trimmed = upper.substring(0, len);
    if (ACCODE_MAP[trimmed]) return ACCODE_MAP[trimmed];
  }

  // Return original code if no match found
  Logger.warn('Unknown ACCode - using raw value', { acCode: code });
  return `Cessna ${code}`;
}

// ============================================================================
// DUTY POSITION CONSTANTS
// ============================================================================

/** Duty position ID for Maintenance Officer */
const DUTY_MAINTENANCE_OFFICER = 'Maintenance Officer';

/** Duty position ID for Transportation Officer */
const DUTY_TRANSPORTATION_OFFICER = 'Transportation Officer';

// DutyPosition.txt columns (0-indexed)
const DUTY_COL = {
  CAPID:      0,
  DUTY:       1,
  FUNCT_AREA: 2,
  LVL:        3,
  ASST:       4,
  USR_ID:     5,
  DATE_MOD:   6,
  ORGID:      7
};

// ============================================================================
// COLUMN INDEX CONSTANTS
// ============================================================================

// Aircraft.txt columns (0-indexed, after header is stripped by parseFile)
const AIRCRAFT_COL = {
  WING:        0,
  UNIT:        1,
  REGION:      2,
  TAILNO:      3,
  CALLSIGN:    4,
  YR_MANF:     5,
  AC_CODE:     6,
  SERIAL_NO:   7,
  DATE_ACQ:    8,
  DATE_ASG:    9,
  SOURCE:      10,
  ENGINE:      11,
  YR_REFURB:   12,
  COST:        13,
  LAT:         14,
  LONG:        15,
  AIRPORT_ID:  16,
  TTAF_PURCH:  17,
  TTEN_PURCH:  18,
  TTAF_CUR:    19,
  TTEN_CUR:    20,
  FIXED_ASSET: 21,
  INSURED:     22,
  STATUS:      23,
  ORGID:       24,
  STOCK_CLASS: 25
};

// Vehicles.txt columns (0-indexed)
const VEHICLE_COL = {
  WING:        0,
  UNIT:        1,
  REGION:      2,
  CAP_ID:      3,
  DEPT:        4,
  MAKE:        5,
  YR_MFGR:    6,
  VEH_TYPE:    7,
  ROADABLE:    8,
  VIN_ID:      9,
  ACQ_DATE:    10,
  SOURCE:      11,
  SOURCE_TYPE: 12,
  CURR_VALUE:  13,
  INS_PREM:    14,
  ODOMETER:    15,
  FIXED_ASSET: 16,
  ORGID:       17
};

// OrgAddresses.txt columns (0-indexed)
const ADDR_COL = {
  ORGID:    0,
  WING:     1,
  UNIT:     2,
  TYPE:     3,
  PRIORITY: 4,
  ADDR1:    5,
  ADDR2:    6,
  CITY:     7,
  STATE:    8,
  ZIP:      9,
  LAT:      10,
  LONG:     11,
  USR_ID:   12,
  DATE_MOD: 13
};

// ============================================================================
// MAIN ENTRY POINT
// ============================================================================

/**
 * Main function to sync all CAP resources (aircraft and vehicles) to
 * Google Workspace calendar resources.
 *
 * @returns {Object} Summary of operations performed
 */
function updateResources() {
  clearCache();
  const start = new Date();
  Logger.info('Starting resource sync');

  const summary = {
    buildings: { created: 0, updated: 0, errors: 0 },
    aircraft:  { created: 0, updated: 0, skipped: 0, errors: 0 },
    vehicles:  { created: 0, updated: 0, errors: 0 },
    startTime: start.toISOString()
  };

  try {
    // Load existing resources and buildings for delta comparison
    const existingResources = getExistingResources_();
    const existingBuildings = getExistingBuildings_();

    // --- Aircraft ---
    const aircraftResults = syncAircraftResources_(existingResources, existingBuildings);
    summary.buildings.created  += aircraftResults.buildingsCreated;
    summary.buildings.updated  += aircraftResults.buildingsUpdated;
    summary.buildings.errors   += aircraftResults.buildingErrors;
    summary.aircraft.created   = aircraftResults.created;
    summary.aircraft.updated   = aircraftResults.updated;
    summary.aircraft.skipped   = aircraftResults.skipped;
    summary.aircraft.errors    = aircraftResults.errors;

    // Refresh building cache after aircraft pass (new airport buildings may exist)
    const refreshedBuildings = getExistingBuildings_();

    // --- Vehicles ---
    const vehicleResults = syncVehicleResources_(existingResources, refreshedBuildings);
    summary.buildings.created  += vehicleResults.buildingsCreated;
    summary.buildings.updated  += vehicleResults.buildingsUpdated;
    summary.buildings.errors   += vehicleResults.buildingErrors;
    summary.vehicles.created   = vehicleResults.created;
    summary.vehicles.updated   = vehicleResults.updated;
    summary.vehicles.errors    = vehicleResults.errors;

  } catch (err) {
    Logger.error('Resource sync failed', err);
    throw err;
  }

  summary.endTime  = new Date().toISOString();
  summary.duration = new Date() - start;

  Logger.info('Resource sync completed', summary);
  return summary;
}

// ============================================================================
// AIRCRAFT SYNC
// ============================================================================

/**
 * Syncs aircraft from Aircraft.txt to Google Workspace resources.
 *
 * @param {Object} existingResources - Map of resourceEmail → resource object
 * @param {Object} existingBuildings - Map of buildingId → building object
 * @returns {Object} Counts of created/updated/skipped/errors
 */
function syncAircraftResources_(existingResources, existingBuildings) {
  const results = {
    buildingsCreated: 0, buildingsUpdated: 0, buildingErrors: 0,
    created: 0, updated: 0, skipped: 0, errors: 0
  };

  const aircraftData = parseFile('Aircraft');
  const dutyMap   = buildDutyOfficerMap_();
  const squadrons = getSquadrons();

  // Collect all unique airport IDs from active, FAA-registered aircraft first,
  // then resolve them all in one API pass before the main loop.
  const uniqueAirportIds = [
    ...new Set(
      aircraftData
        .filter(row => {
          const tailno = (row[AIRCRAFT_COL.TAILNO] || '').trim().toUpperCase();
          const status = (row[AIRCRAFT_COL.STATUS] || '').trim();
          return /^N[0-9]{1,5}[A-HJ-NP-Z]{0,2}$/.test(tailno) && status.toLowerCase() === 'active';
        })
        .map(row => normalizeAirportId(row[AIRCRAFT_COL.AIRPORT_ID] || ''))
        .filter(isValidAirportId)
    )
  ];

  const airportInfo = fetchAirportInfo_(uniqueAirportIds);

  // Load wing-specific aircraft overrides from Drive JSON.
  // Use this for aircraft whose CAPWATCH AirPortID is a placeholder (e.g. "ASSIGN")
  // but whose actual base and owning unit are known.
  //
  // File format: { "N28FA": { "airportId": "KWHP", "orgid": "663" }, ... }
  // airportId must match an existing airport building ID (ICAO code).
  // orgid must match the owning squadron's CAPWATCH orgid.
  const AIRCRAFT_OVERRIDES = loadJsonOverride_('AircraftOverrides.json');

  for (let i = 0; i < aircraftData.length; i++) {
    const row = aircraftData[i];

    // Only process FAA-registered aircraft matching the N-number format:
    // N + 1-5 digits + 0-2 letters (excluding I and O)
    // Examples: N445CP, N9183E, N172CA, N1
    const tailno = (row[AIRCRAFT_COL.TAILNO] || '').trim().toUpperCase();
    if (!tailno || !/^N[0-9]{1,5}[A-HJ-NP-Z]{0,2}$/.test(tailno)) {
      results.skipped++;
      continue;
    }

    // Only process active aircraft
    const status = (row[AIRCRAFT_COL.STATUS] || '').trim();
    if (status.toLowerCase() !== 'active') {
      results.skipped++;
      Logger.info('Skipping inactive aircraft', { tailno: tailno, status: status });
      continue;
    }

    const acCode    = (row[AIRCRAFT_COL.AC_CODE]    || '').trim();
    const yrManf    = (row[AIRCRAFT_COL.YR_MANF]    || '').trim();

    // Apply manual override if present, otherwise use CAPWATCH data
    const override  = AIRCRAFT_OVERRIDES[tailno];
    const airportId = override
      ? override.airportId
      : normalizeAirportId(row[AIRCRAFT_COL.AIRPORT_ID] || '');
    const orgid     = override
      ? override.orgid
      : (row[AIRCRAFT_COL.ORGID] || '').trim();
    if (override) {
      Logger.info('Applying manual aircraft override', { tailno, airportId, orgid });
    }

    const model = resolveAircraftModel(acCode);
    const resourceName = `${tailno} - ${model} (${yrManf})`;

    // Resolve airport name and coordinates from cache/API
    const ap          = airportId ? (airportInfo[airportId] || {}) : {};
    const apNameRaw   = ap.name  || airportId;
    // Deduplicate slash-separated segments (e.g. "Camarillo/Camarillo" → "Camarillo")
    const apNameDeduped = apNameRaw
      .replace(/\/\s*$/, '').trim()
      .split('/')
      .map(s => s.trim())
      .filter((s, i, arr) => arr.indexOf(s) === i)
      .join('/');
    const apName      = unitTitleCase(apNameDeduped);
    const apCity      = ap.city  ? `${unitTitleCase(ap.city)}${ap.state ? ', ' + ap.state : ''}` : '';
    const apLabel     = apCity   ? `${apName} (${airportId}) — ${apCity}` : `${apName} (${airportId})`;
    const apLat       = ap.lat   || '';
    const apLng       = ap.lng   || '';

    // Maintenance Officer for this unit
    const maintOfficer = dutyMap[orgid] && dutyMap[orgid].maintenance
      ? dutyMap[orgid].maintenance
      : null;
    const maintContact = maintOfficer
      ? `Maintenance Officer: ${maintOfficer.name} (${maintOfficer.email})`
      : 'Maintenance Officer: unassigned';

    const org     = squadrons[orgid] || null;
    const orgName = org ? unitTitleCase(org.name) : `Org ${orgid}`;

    const description            = `CAP Aircraft | Model: ${model} | Year: ${yrManf} | Airport: ${apLabel} | Unit: ${orgName} (${orgid})`;
    const userVisibleDescription = `${model} (${yrManf}) based at ${apLabel} | ${orgName} | ${maintContact}`;

    // Ensure building exists for this airport, using resolved name and coordinates
    if (airportId && isValidAirportId(airportId)) {
      const airportAddress = apCity ? {
        regionCode:         'US',
        administrativeArea: ap.state || '',
        locality:           ap.city  || '',
        addressLines:       [apName]
      } : null;
      const buildingResult = ensureBuilding_(
        airportId,
        apName,
        airportAddress,
        apLat,
        apLng,
        existingBuildings
      );
      if (buildingResult === 'created') results.buildingsCreated++;
      else if (buildingResult === 'updated') results.buildingsUpdated++;
      else if (buildingResult === 'error') results.buildingErrors++;
    }

    // Upsert the resource — skip if airport ID is invalid (e.g. ASSIGN)
    if (!airportId || !isValidAirportId(airportId)) {
      Logger.warn('Skipping aircraft resource - invalid airportId', { resourceName, airportId });
      results.skipped++;
      continue;
    }
    const resourceResult = upsertResource_(
      resourceName,
      'Aircraft',
      airportId,
      description,
      userVisibleDescription,
      existingResources
    );

    if (resourceResult === 'created') results.created++;
    else if (resourceResult === 'updated') results.updated++;
    else if (resourceResult === 'error') results.errors++;
  }

  Logger.info('Aircraft sync completed', results);
  return results;
}

// ============================================================================
// VEHICLE SYNC
// ============================================================================

/**
 * Syncs vehicles from Vehicles.txt to Google Workspace resources.
 * Uses the unit's MEETING address from OrgAddresses.txt as the building.
 *
 * @param {Object} existingResources - Map of resourceEmail → resource object
 * @param {Object} existingBuildings - Map of buildingId → building object
 * @returns {Object} Counts of created/updated/errors
 */
function syncVehicleResources_(existingResources, existingBuildings) {
  const results = {
    buildingsCreated: 0, buildingsUpdated: 0, buildingErrors: 0,
    created: 0, updated: 0, errors: 0
  };

  const vehicleData = parseFile('Vehicles');
  const addressMap  = buildOrgAddressMap_();
  const dutyMap     = buildDutyOfficerMap_();
  const squadrons   = getSquadrons();

  // --- Pass 1: ensure a building exists for every squadron in the wing,
  // regardless of whether it has any vehicles. This makes all units
  // available as locations when members create calendar events.
  for (const [orgid, org] of Object.entries(squadrons)) {
    const charter = org.charter;
    if (!charter) continue;

    // Skip units without fixed meeting locations:
    // x-000 = Wing/Region HQ administrative placeholder
    // x-999 = Legislative Squadron (state legislators and staff; no regular meeting location)
    const unitNum = org.unit ? String(org.unit).padStart(3, '0') : '';
    if (unitNum === '000' || unitNum === '999') continue;

    const addr = addressMap[orgid];
    if (!addr) continue;  // no address on record - skip silently

    const buildingName = unitTitleCase(org.name);
    const streetLines  = [addr.addr1, addr.addr2].filter(Boolean);
    const address = {
      regionCode:         'US',
      postalCode:         addr.zip   || '',
      administrativeArea: addr.state || '',
      locality:           addr.city  || '',
      addressLines:       streetLines
    };

    const result = ensureBuilding_(charter, buildingName, address, addr.lat, addr.lng, existingBuildings);
    if (result === 'created') results.buildingsCreated++;
    else if (result === 'updated') results.buildingsUpdated++;
    else if (result === 'error')   results.buildingErrors++;
  }

  // --- Pass 2: sync vehicle resources (buildings already in cache from pass 1)
  for (let i = 0; i < vehicleData.length; i++) {
    const row = vehicleData[i];

    const orgid   = (row[VEHICLE_COL.ORGID]   || '').trim();
    const make    = (row[VEHICLE_COL.MAKE]     || '').trim();
    const yrMfgr  = (row[VEHICLE_COL.YR_MFGR] || '').trim();
    const vehType = (row[VEHICLE_COL.VEH_TYPE] || '').trim();
    const capId   = (row[VEHICLE_COL.CAP_ID]   || '').trim();

    const org = squadrons[orgid] || null;
    const orgName = org ? unitTitleCase(org.name) : `Org ${orgid}`;

    const resourceName = `${capId} - ${yrMfgr} ${make} ${vehType}`;
    const description  = `CAP Vehicle | ID: ${capId} | Type: ${vehType} | Make: ${make} | Year: ${yrMfgr} | Unit: ${orgName} (${orgid})`;

    // Transportation Officer for this unit
    const transOfficer = dutyMap[orgid] && dutyMap[orgid].transportation
      ? dutyMap[orgid].transportation
      : null;
    const transContact = transOfficer
      ? `Transportation Officer: ${transOfficer.name} (${transOfficer.email})`
      : 'Transportation Officer: unassigned';

    const userVisibleDescription = `${yrMfgr} ${make} ${vehType} | ${orgName} | ${transContact}`;

    // Look up unit's address for building; use charter as stable building ID
    const addr = addressMap[orgid];
    const charter = org ? org.charter : null;
    let buildingId = charter ? charter : `ORG-${orgid}`;

    if (addr) {
      const buildingName = org
        ? unitTitleCase(org.name)
        : addr.city
          ? `${unitTitleCase(addr.city)}, ${addr.state}`
          : `Unit ${addr.unit}`;
      const streetLines = [addr.addr1, addr.addr2].filter(Boolean);
      const vehicleAddress = {
        regionCode:         'US',
        postalCode:         addr.zip   || '',
        administrativeArea: addr.state || '',
        locality:           addr.city  || '',
        addressLines:       streetLines
      };

      const buildingResult = ensureBuilding_(
        buildingId,
        buildingName,
        vehicleAddress,
        addr.lat,
        addr.lng,
        existingBuildings
      );
      if (buildingResult === 'created') results.buildingsCreated++;
      else if (buildingResult === 'updated') results.buildingsUpdated++;
      else if (buildingResult === 'error') {
        results.buildingErrors++;
        buildingId = null;
      }
    } else {
      Logger.warn('No MEETING address found for org', { orgid: orgid, orgName: orgName });
      buildingId = null;
    }

    // Skip resource if we have no valid building
    if (!buildingId) {
      Logger.warn('Skipping vehicle resource - no valid building', { capId, orgid });
      results.errors++;
      continue;
    }

    // Upsert the resource
    const resourceResult = upsertResource_(
      resourceName,
      'Vehicle',
      buildingId,
      description,
      userVisibleDescription,
      existingResources
    );

    if (resourceResult === 'created') results.created++;
    else if (resourceResult === 'updated') results.updated++;
    else if (resourceResult === 'error') results.errors++;
  }

  Logger.info('Vehicle sync completed', results);
  return results;
}

// ============================================================================
// DUTY OFFICER HELPER
// ============================================================================

/**
 * Builds a map of ORGID -> { maintenanceOfficer, transportationOfficer }
 * by scanning DutyPosition.txt for MAINT and TRANS duty codes.
 *
 * Each value is an object { capid, email } where email is derived from the
 * member's Workspace account (firstname.lastname@domain), looked up via
 * the members object from getMembers().
 *
 * Only primary (non-assistant) duty holders are used (Asst == '0').
 *
 * @returns {Object} Map of orgid -> { maintenance: {capid, name, email}|null,
 *                                     transportation: {capid, name, email}|null }
 */
function buildDutyOfficerMap_() {
  const dutyData = parseFile('DutyPosition');
  const members  = getMembers(CONFIG.MEMBER_TYPES.ACTIVE, false);
  const map = {};

  for (let i = 0; i < dutyData.length; i++) {
    const row   = dutyData[i];
    const duty  = (row[DUTY_COL.DUTY]  || '').trim();
    const asst  = (row[DUTY_COL.ASST]  || '').trim();
    const orgid = (row[DUTY_COL.ORGID] || '').trim();
    const capid = String(row[DUTY_COL.CAPID] || '').trim();

    // Skip assistants
    if (asst !== '0') continue;

    if (duty !== DUTY_MAINTENANCE_OFFICER && duty !== DUTY_TRANSPORTATION_OFFICER) continue;

    const member = members[capid] || members[Number(capid)];
    if (!member) continue;

    const name  = `${member.rank ? member.rank + ' ' : ''}${member.firstName} ${member.lastName}`.trim();
    const email = member.email ||
      `${member.firstName.toLowerCase()}.${member.lastName.toLowerCase()}${CONFIG.EMAIL_DOMAIN}`;

    if (!map[orgid]) {
      map[orgid] = { maintenance: null, transportation: null };
    }

    if (duty === DUTY_MAINTENANCE_OFFICER && !map[orgid].maintenance) {
      map[orgid].maintenance = { capid, name, email };
    }
    if (duty === DUTY_TRANSPORTATION_OFFICER && !map[orgid].transportation) {
      map[orgid].transportation = { capid, name, email };
    }
  }

  Logger.info('Duty officer map built', { orgsWithOfficers: Object.keys(map).length });
  return map;
}

// ============================================================================
// ADDRESS HELPER
// ============================================================================

/**
 * Builds a map of ORGID → MEETING address from OrgAddresses.txt.
 * If a unit has multiple MEETING addresses, the one with the lowest Priority
 * value (highest priority) is used.
 *
 * @returns {Object} Map of orgid → address object
 */
/**
 * Returns true if the given address line looks like a PO box.
 * @param {string} addr
 * @returns {boolean}
 */
function isPoBox_(addr) {
  return /^\s*P\.?\s*O\.?\s*(BOX|B\.?O\.?X\.?)/i.test(addr || '');
}

/**
 * Builds a map of orgid → best available address from OrgAddresses.txt.
 * Prefers MEETING address type; falls back to non-PO-box MAIL address.
 *
 * @returns {Object} Map of orgid (string) → { addr1, addr2, city, state, zip, lat, lng }
 */
function buildOrgAddressMap_() {
  const addressData = parseFile('OrgAddresses');

  // Collect best MEETING and best MAIL address per org separately
  const meeting = {};
  const mail    = {};

  for (let i = 0; i < addressData.length; i++) {
    const row  = addressData[i];
    const type = (row[ADDR_COL.TYPE] || '').trim().toUpperCase();

    if (type !== 'MEETING' && type !== 'MAIL') continue;

    const orgid    = (row[ADDR_COL.ORGID]    || '').trim();
    const priority = parseInt(row[ADDR_COL.PRIORITY] || '99', 10);
    const addr1    = (row[ADDR_COL.ADDR1] || '').trim();

    const entry = {
      priority: priority,
      unit:     (row[ADDR_COL.UNIT]  || '').trim(),
      addr1:    addr1,
      addr2:    (row[ADDR_COL.ADDR2] || '').trim(),
      city:     (row[ADDR_COL.CITY]  || '').trim(),
      state:    (row[ADDR_COL.STATE] || '').trim(),
      zip:      (row[ADDR_COL.ZIP]   || '').trim(),
      lat:      (row[ADDR_COL.LAT]   || '').trim(),
      lng:      (row[ADDR_COL.LONG]  || '').trim()
    };

    if (type === 'MEETING') {
      if (!meeting[orgid] || priority < meeting[orgid].priority) {
        meeting[orgid] = entry;
      }
    } else if (type === 'MAIL' && !isPoBox_(addr1)) {
      if (!mail[orgid] || priority < mail[orgid].priority) {
        mail[orgid] = entry;
      }
    }
  }

  // Merge: prefer MEETING, fall back to MAIL
  const map = {};
  const allOrgids = new Set([...Object.keys(meeting), ...Object.keys(mail)]);
  allOrgids.forEach(orgid => {
    if (meeting[orgid]) {
      map[orgid] = { ...meeting[orgid], addressType: 'MEETING' };
    } else if (mail[orgid]) {
      map[orgid] = { ...mail[orgid], addressType: 'MAIL' };
      Logger.info('Using MAIL address fallback for org', { orgid });
    }
  });

  Logger.info('Org address map built', { count: Object.keys(map).length });
  return map;
}

// ============================================================================
// GOOGLE WORKSPACE RESOURCE HELPERS
// ============================================================================

/**
 * Retrieves all existing calendar resources in the domain.
 * Returns a map keyed by resource name (lowercased) for lookup.
 *
 * @returns {Object} Map of lower-cased resource name → resource object
 */
function getExistingResources_() {
  const resources = {};
  let nextPageToken = '';

  do {
    const page = executeWithRetry(() =>
      AdminDirectory.Resources.Calendars.list('my_customer', {
        maxResults: 500,
        pageToken: nextPageToken
      })
    );

    nextPageToken = page.nextPageToken || '';

    if (page.items) {
      page.items.forEach(resource => {
        resources[resource.resourceName.toLowerCase()] = resource;
      });
    }
  } while (nextPageToken);

  Logger.info('Existing resources loaded', { count: Object.keys(resources).length });
  return resources;
}

/**
 * Retrieves all existing buildings in the domain.
 * Returns a map keyed by buildingId (lowercased).
 *
 * @returns {Object} Map of lower-cased buildingId → building object
 */
function getExistingBuildings_() {
  const buildings = {};
  let nextPageToken = '';

  do {
    const page = executeWithRetry(() =>
      AdminDirectory.Resources.Buildings.list('my_customer', {
        maxResults: 500,
        pageToken: nextPageToken
      })
    );

    nextPageToken = page.nextPageToken || '';

    if (page.buildings) {
      page.buildings.forEach(building => {
        buildings[building.buildingId.toLowerCase()] = building;
      });
    }
  } while (nextPageToken);

  const buildingIds = Object.keys(buildings).slice(0, 5);
  Logger.info('Existing buildings loaded', { count: Object.keys(buildings).length, sampleIds: buildingIds });
  return buildings;
}

/**
 * Creates or updates a Google Workspace building.
 *
 * @param {string} buildingId   - Unique building identifier (e.g. ICAO code or ORG-12345)
 * @param {string} buildingName - Human-readable name shown in Calendar UI
 * @param {string} description  - Address or description
 * @param {string} lat          - Latitude (string)
 * @param {string} lng          - Longitude (string)
 * @param {Object} existingBuildings - Current building map (mutated on create)
 * @returns {string} 'created' | 'updated' | 'unchanged' | 'error'
 */
/**
 * @param {string} buildingId
 * @param {string} buildingName
 * @param {Object|null} address - Structured address: { lines, city, state, postalCode, countryCode }
 * @param {string} lat
 * @param {string} lng
 * @param {Object} existingBuildings
 */
function ensureBuilding_(buildingId, buildingName, address, lat, lng, existingBuildings) {
  const key = buildingId.toLowerCase();
  const existing = existingBuildings[key];

  // Skip buildings that already failed this run to avoid per-vehicle retries
  if (existing && existing._unavailable) {
    return 'error';
  }

  const buildingBody = {
    buildingId:   buildingId,
    buildingName: buildingName,
    floorNames:   ['1']
  };

  // Structured address field (supported by Admin SDK)
  if (address) {
    buildingBody.address = address;
  }

  // Add coordinates if available
  if (lat && lng && !isNaN(parseFloat(lat)) && !isNaN(parseFloat(lng))) {
    buildingBody.coordinates = {
      latitude:  parseFloat(lat),
      longitude: parseFloat(lng)
    };
  }

  try {
    if (existing) {
      // Only update if something meaningful changed
      const existingAddr = JSON.stringify(existing.address || null);
      const newAddr      = JSON.stringify(address || null);
      if (existing.buildingName === buildingName && existingAddr === newAddr) {
        return 'unchanged';
      }
      executeWithRetry(() =>
        AdminDirectory.Resources.Buildings.patch(buildingBody, 'my_customer', encodeURIComponent(buildingId))
      );
      existingBuildings[key] = Object.assign({}, existing, buildingBody);
      Logger.info('Building updated', { buildingId, buildingName });
      return 'updated';
    } else {
      try {
        executeWithRetry(() =>
          AdminDirectory.Resources.Buildings.insert(buildingBody, 'my_customer')
        );
        existingBuildings[key] = buildingBody;
        Logger.info('Building created', { buildingId, buildingName });
        return 'created';
      } catch (insertErr) {
        // 409 means the building already exists under this ID (e.g. created in a prior
        // run or the local cache was stale). Fall back to patch so the run succeeds.
        // The buildingId must be URL-encoded in the patch path since it may contain hyphens.
        if (insertErr.details && insertErr.details.code === 409) {
          Logger.info('Building already exists, trying update instead', { buildingId, buildingName });
          try {
            executeWithRetry(() =>
              AdminDirectory.Resources.Buildings.update(buildingBody, 'my_customer', buildingId)
            );
            existingBuildings[key] = buildingBody;
            Logger.info('Building updated (post-409)', { buildingId, buildingName });
            return 'updated';
          } catch (patchErr) {
            // If patch also fails, the building may truly not exist yet under this ID
            // but something else conflicts. Mark as existing anyway to avoid re-attempting
            // on subsequent vehicles for the same unit this run.
            Logger.warn('Patch after 409 failed — treating building as unavailable this run', {
              buildingId,
              buildingName,
              errorMessage: patchErr.message
            });
            existingBuildings[key] = { buildingId, buildingName, _unavailable: true };
            return 'error';
          }
        }
        throw insertErr; // re-throw non-409 errors to outer catch
      }
    }
  } catch (err) {
    Logger.error('Failed to upsert building', {
      buildingId,
      buildingName,
      errorMessage: err.message,
      errorCode: err.details?.code
    });
    return 'error';
  }
}

/**
 * Creates or updates a Google Workspace calendar resource.
 *
 * Resource IDs are derived from a sanitized version of the resource name
 * to keep them stable across runs.
 *
 * @param {string} resourceName     - Display name for the resource
 * @param {string} resourceCategory - 'Aircraft' or 'Vehicle'
 * @param {string} buildingId       - ID of the building this resource belongs to
 * @param {string} description      - Additional details about the resource
 * @param {Object} existingResources - Current resource map (mutated on create)
 * @returns {string} 'created' | 'updated' | 'unchanged' | 'error'
 */
function upsertResource_(resourceName, resourceCategory, buildingId, description, userVisibleDescription, existingResources) {
  const key = resourceName.toLowerCase();
  const existing = existingResources[key];

  // Derive a stable resource ID from the name
  const resourceId = sanitizeResourceId_(resourceName);

  const resourceBody = {
    resourceId:              resourceId,
    resourceName:            resourceName,
    resourceCategory:        'OTHER',              // Required field; category info goes in resourceType
    resourceType:            resourceCategory,      // 'Aircraft' or 'Vehicle'
    buildingId:              buildingId,
    resourceDescription:     description,
    userVisibleDescription:  userVisibleDescription,
    floorName:               '1'
  };

  try {
    if (existing) {
      // Check if update is needed
      if (
        existing.buildingId             === buildingId &&
        existing.resourceDescription    === description &&
        existing.userVisibleDescription === userVisibleDescription &&
        existing.resourceType           === resourceCategory
      ) {
        return 'unchanged';
      }
      executeWithRetry(() =>
        AdminDirectory.Resources.Calendars.patch(
          resourceBody,
          'my_customer',
          existing.resourceId
        )
      );
      existingResources[key] = Object.assign({}, existing, resourceBody);
      Logger.info('Resource updated', { resourceName, resourceCategory });
      return 'updated';
    } else {
      executeWithRetry(() =>
        AdminDirectory.Resources.Calendars.insert(resourceBody, 'my_customer')
      );
      existingResources[key] = resourceBody;
      Logger.info('Resource created', { resourceName, resourceCategory });
      return 'created';
    }
  } catch (err) {
    Logger.error('Failed to upsert resource', {
      resourceName,
      resourceCategory,
      buildingId,
      errorMessage: err.message,
      errorCode: err.details?.code
    });
    return 'error';
  }
}

/**
 * Converts a resource name into a safe, stable resource ID.
 * Strips non-alphanumeric characters and limits to 64 characters.
 *
 * @param {string} name - Resource display name
 * @returns {string} Sanitized ID string
 */
/**
 * Converts a CAP unit name from ALL-CAPS to title case, preserving
 * known acronyms (HQ, CAP) in full caps.
 *
 * @param {string} name - Unit name as it appears in CAPWATCH (e.g. "FRESNO COMPOSITE SQUADRON")
 * @returns {string} Title-cased name (e.g. "Fresno Composite Squadron")
 */
/**
 * Normalizes a raw AirPortID value from Aircraft.txt into a clean identifier
 * suitable for use as a building ID and display label.
 *
 * Handles:
 *   - Standard ICAO codes:         "KPAO"           → "KPAO"
 *   - FAA LIDs with name:          "C83 Byron"      → "C83"
 *   - FAA LIDs with name:          "L18 Fallbrook Airpark" → "L18"
 *   - ICAO with extra info after /: "KPAO/ Hangar side..." → "KPAO"
 *
 * The first whitespace- or slash-delimited token is treated as the identifier.
 *
 * @param {string} raw - Raw AirPortID string from Aircraft.txt
 * @returns {string} Cleaned airport identifier (uppercased), or empty string if blank
 */
// ============================================================================
// AIRPORT INFO LOOKUP
// ============================================================================

/**
 * Filename used to persist the airport info cache in the CAPWATCH data folder.
 * Avoids redundant API calls across runs.
 */
const AIRPORT_CACHE_FILENAME = 'AirportCache.json';

/**
 * Loads the persisted airport info cache from Drive.
 * Returns an empty object if the file doesn't exist yet.
 *
/**
 * Loads a wing-specific JSON override file from the CAPWATCH Drive folder.
 * Returns an empty object if the file doesn't exist, so the script runs
 * safely with no overrides configured.
 *
 * @param {string} filename - Name of the JSON file in the CAPWATCH data folder
 * @returns {Object} Parsed JSON object, or {} if file not found or invalid
 */
function loadJsonOverride_(filename) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
    const files  = folder.getFilesByName(filename);
    if (!files.hasNext()) {
      Logger.debug(`Override file not found, skipping: ${filename}`);
      return {};
    }
    const content = files.next().getBlob().getDataAsString();
    return JSON.parse(content);
  } catch (err) {
    Logger.warn(`Failed to load override file ${filename}`, { error: err.message });
    return {};
  }
}

/**
 * Loads the airport info cache from Drive (AirportCache.json).
 * Returns an empty object if the file doesn't exist yet.
 *
 * @returns {Object} Map of airport identifier (uppercased) → { name, city, state, lat, lng }
 */
function loadAirportCache_() {
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files  = folder.getFilesByName(AIRPORT_CACHE_FILENAME);
  if (files.hasNext()) {
    try {
      return JSON.parse(files.next().getBlob().getDataAsString());
    } catch (e) {
      Logger.warn('Could not parse airport cache - starting fresh', { error: e.message });
    }
  }
  return {};
}

/**
 * Saves the airport info cache to Drive, creating or overwriting the file.
 *
 * @param {Object} cache - Map of airport identifier → airport info object
 */
function saveAirportCache_(cache) {
  const folder  = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const content = JSON.stringify(cache, null, 2);
  const files   = folder.getFilesByName(AIRPORT_CACHE_FILENAME);
  if (files.hasNext()) {
    files.next().setContent(content);
  } else {
    folder.createFile(AIRPORT_CACHE_FILENAME, content, MimeType.PLAIN_TEXT);
  }
  Logger.info('Airport cache saved', { count: Object.keys(cache).length });
}

/**
 * Fetches name and location for a set of airport identifiers using the
 * Aviation Weather Center API (https://aviationweather.gov/api/data/airport).
 * Supports both ICAO codes (e.g. KPAO) and FAA LIDs (e.g. C83, L18).
 * No API key required.
 *
 * Results are cached in Drive (AirportCache.json) and only missing entries
 * are fetched, keeping API calls minimal across runs.
 *
 * @param {string[]} airportIds - Array of normalized airport identifiers
 * @returns {Object} Map of identifier (uppercased) → { name, city, lat, lng }
 */
function fetchAirportInfo_(airportIds) {
  const cache = loadAirportCache_();

  // Load wing-specific airport overrides from Drive JSON.
  // These are FAA LIDs that the AWC API doesn't resolve (e.g. non-ICAO identifiers).
  // File format: { "C83": { "name": "Byron Airport", "city": "Byron", "state": "CA",
  //                          "lat": 37.8284, "lng": -121.6258 }, ... }
  const airportOverrides = loadJsonOverride_('AirportOverrides.json');
  Object.entries(airportOverrides).forEach(([id, info]) => {
    if (!cache[id.toUpperCase()]) cache[id.toUpperCase()] = info;
  });

  const missing = airportIds.filter(id => isValidAirportId(id) && !cache[id.toUpperCase()]);

  if (missing.length === 0) {
    Logger.info('All airports found in cache', { count: airportIds.length });
    return cache;
  }

  Logger.info('Fetching airport info from AWC API', { missing: missing.length });

  // Fetch in batches of 5 to avoid AWC API timeouts on large requests
  const BATCH_SIZE = 5;
  for (let b = 0; b < missing.length; b += BATCH_SIZE) {
    const batch = missing.slice(b, b + BATCH_SIZE);
    const ids   = batch.join(',');
    const url   = `https://aviationweather.gov/api/data/airport?ids=${encodeURIComponent(ids)}&format=json`;

  try {
    const response = executeWithRetry(() =>
      UrlFetchApp.fetch(url, { muteHttpExceptions: true })
    );

    const code = response.getResponseCode();
    if (code !== 200) {
      Logger.warn('AWC airport API returned non-200', { code: code, ids: ids });
      continue;
    }

    const results = JSON.parse(response.getContentText());

    if (!Array.isArray(results)) {
      Logger.warn('AWC airport API returned unexpected format', { ids: ids });
      continue;
    }

    results.forEach(ap => {
      // AWC returns 'icaoId' (may be FAA LID for non-ICAO airports) and 'site' for name
      const key  = (ap.icaoId || ap.lid || '').toUpperCase();
      const lid  = (ap.lid   || '').toUpperCase();
      if (!key) return;

      const info = {
        name: ap.site   || ap.name || key,
        city: ap.city   || '',
        state: ap.state || '',
        lat:  ap.lat    || null,
        lng:  ap.lon    || null
      };

      cache[key] = info;
      // Also index by FAA LID if different from ICAO key (e.g. cache['C83'] as well as 'KC83')
      if (lid && lid !== key) {
        cache[lid] = info;
      }
    });

    // Mark any IDs in this batch that came back with no data so we don't retry them
    batch.forEach(id => {
      const upper = id.toUpperCase();
      if (!cache[upper]) {
        Logger.warn('Airport not found in AWC API', { id: upper });
        cache[upper] = { name: upper, city: '', state: '', lat: null, lng: null };
      }
    });

  } catch (e) {
    Logger.error('Failed to fetch airport info from AWC', {
      errorMessage: e.message,
      ids: ids
    });
  }
  } // end batch loop

  // Second pass: retry any unfound IDs with a K prefix (FAA LID → ICAO attempt)
  // e.g. C83 → KC83, L18 → KL18
  const retryBatch = missing.filter(id => {
    const upper = id.toUpperCase();
    return cache[upper] && !cache[upper].name || // marked as not found
           (cache[upper] && cache[upper].name === upper); // fallback placeholder
  }).map(id => ({ original: id.toUpperCase(), prefixed: 'K' + id.toUpperCase() }))
    .filter(pair => !cache[pair.prefixed]);

  if (retryBatch.length > 0) {
    Logger.info('Retrying unfound LIDs with K prefix', { count: retryBatch.length });
    const retryIds  = retryBatch.map(p => p.prefixed).join(',');
    const retryUrl  = `https://aviationweather.gov/api/data/airport?ids=${encodeURIComponent(retryIds)}&format=json`;
    try {
      const retryResp = executeWithRetry(() =>
        UrlFetchApp.fetch(retryUrl, { muteHttpExceptions: true })
      );
      if (retryResp.getResponseCode() === 200) {
        const retryResults = JSON.parse(retryResp.getContentText());
        if (Array.isArray(retryResults)) {
          retryResults.forEach(ap => {
            const icao = (ap.icaoId || '').toUpperCase();
            const lid  = (ap.lid   || '').toUpperCase();
            if (!icao) return;
            const info = {
              name:  ap.site  || ap.name || icao,
              city:  ap.city  || '',
              state: ap.state || '',
              lat:   ap.lat   || null,
              lng:   ap.lon   || null
            };
            // Store under both the ICAO code and the original bare LID
            cache[icao] = info;
            if (lid && lid !== icao) cache[lid] = info;
            // Also update the original bare key if it was a placeholder
            const origPair = retryBatch.find(p => p.prefixed === icao || p.prefixed === 'K' + lid);
            if (origPair) cache[origPair.original] = info;
          });
        }
      }
    } catch (e) {
      Logger.warn('K-prefix retry failed', { error: e.message });
    }
  }

  saveAirportCache_(cache);

  return cache;
}

/**
 * Normalizes a raw AirPortID value from CAPWATCH by extracting the first
 * token before any slash, comma, or whitespace, and uppercasing it.
 * Example: "KSNA / John Wayne" → "KSNA"
 *
 * @param {string} raw - Raw airport identifier from CAPWATCH
 * @returns {string} Normalized identifier, or empty string if blank
 */
function normalizeAirportId(raw) {
  if (!raw || !raw.trim()) return '';
  // Split on slash, comma, or whitespace; take the first token
  return raw.trim().split(/[\/,\s]+/)[0].toUpperCase();
}

/**
 * Returns true if the given string looks like a valid airport identifier:
 * - ICAO: 3-4 uppercase letters (e.g. KPAO, KBFL, C83 is FAA LID)
 * - FAA LID: 2-5 alphanumeric characters beginning with a letter or digit
 *
 * Rejects placeholder values like "ASSIGN", "NONE", "TBD", "N/A", or anything
 * longer than 5 characters after normalization.
 *
 * @param {string} id - Normalized airport identifier
 * @returns {boolean}
 */
function isValidAirportId(id) {
  if (!id) return false;
  // Must be 2-5 alphanumeric characters, no spaces (already normalized)
  return /^[A-Z0-9]{2,5}$/.test(id);
}

/**
 * Converts a unit name to title case while preserving known CAP acronyms
 * (HQ, CAP) in uppercase.
 *
 * @param {string} name - Unit name to convert
 * @returns {string} Title-cased name with acronyms preserved
 */
function unitTitleCase(name) {
  const PRESERVE = new Set(['HQ', 'CAP']);
  return (name || '')
    .toLowerCase()
    .replace(/\b\w+/g, word => {
      const upper = word.toUpperCase();
      return PRESERVE.has(upper) ? upper : word[0].toUpperCase() + word.slice(1);
    });
}

/**
 * Sanitizes a string for use as a Calendar resource ID by replacing
 * non-alphanumeric characters with underscores and truncating to 64 chars.
 *
 * @param {string} name - Raw name to sanitize
 * @returns {string} Sanitized resource ID string
 */
function sanitizeResourceId_(name) {
  return name
    .replace(/[^a-zA-Z0-9\-_]/g, '_')
    .replace(/_+/g, '_')
    .substring(0, 64);
}

// ============================================================================
// TEST FUNCTIONS
// ============================================================================

/**
 * Preview aircraft that would be synced without making any changes.
 * Logs count and a sample entry.
 *
 * @returns {Array} Array of aircraft objects that would be processed
 */
function previewAircraftResources() {
  clearCache();
  const aircraftData = parseFile('Aircraft');
  const preview = [];

  for (let i = 0; i < aircraftData.length; i++) {
    const row = aircraftData[i];
    const tailno = (row[AIRCRAFT_COL.TAILNO] || '').trim();
    const status = (row[AIRCRAFT_COL.STATUS] || '').trim();

    if (!/^N[0-9]{1,5}[A-HJ-NP-Z]{0,2}$/.test(tailno.toUpperCase())) continue;
    if (status !== 'Active') continue;

    preview.push({
      tailno:    tailno,
      model:     resolveAircraftModel(row[AIRCRAFT_COL.AC_CODE]),
      year:      row[AIRCRAFT_COL.YR_MANF],
      airport:   row[AIRCRAFT_COL.AIRPORT_ID],
      orgid:     row[AIRCRAFT_COL.ORGID],
      status:    status
    });
  }

  console.log(`\n=== AIRCRAFT PREVIEW: ${preview.length} aircraft found ===\n`);
  if (preview.length > 0) {
    console.log('Sample entry:', JSON.stringify(preview[0], null, 2));
  }

  Logger.info('Aircraft preview completed', { count: preview.length });
  return preview;
}

/**
 * Preview vehicles that would be synced without making any changes.
 * Logs count and a sample entry.
 *
 * @returns {Array} Array of vehicle objects that would be processed
 */
function previewVehicleResources() {
  clearCache();
  const vehicleData = parseFile('Vehicles');
  const addressMap  = buildOrgAddressMap_();
  const preview = [];

  for (let i = 0; i < vehicleData.length; i++) {
    const row = vehicleData[i];
    const orgid = (row[VEHICLE_COL.ORGID] || '').trim();
    preview.push({
      capId:    row[VEHICLE_COL.CAP_ID],
      make:     row[VEHICLE_COL.MAKE],
      year:     row[VEHICLE_COL.YR_MFGR],
      type:     row[VEHICLE_COL.VEH_TYPE],
      orgid:    orgid,
      address:  addressMap[orgid] || null
    });
  }

  console.log(`\n=== VEHICLE PREVIEW: ${preview.length} vehicles found ===\n`);
  if (preview.length > 0) {
    console.log('Sample entry:', JSON.stringify(preview[0], null, 2));
  }

  Logger.info('Vehicle preview completed', { count: preview.length });
  return preview;
}

/**
 * Runs the full resource sync with a dry-run flag to test ACCode resolution.
 * Does not call Google APIs — only logs what would be created.
 *
 * @returns {void}
 */
function testAcCodeResolution() {
  const testCodes = [
    'C172', 'C172S', 'C172R',
    'C182', 'C182T', 'T182T',
    'U206G', 'C206H', 'TU206G',
    'C206U', 'P210N', 'UNKNOWN'
  ];

  console.log('\n=== ACCODE RESOLUTION TEST ===\n');
  testCodes.forEach(code => {
    console.log(`${code.padEnd(10)} → ${resolveAircraftModel(code)}`);
  });
}