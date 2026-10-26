/**
 * RegionStaffMembers.gs
 * Version: 2.0.0
 *
 * Finds region-HQ members (primary unit PCR-PCR-001) who ALSO hold a duty
 * position in one of THIS wing's units — the people who today only get a
 * cawgcap.org account because someone added them to the Manual Members sheet.
 *
 * WHY THEY ARE INVISIBLE TO THE NORMAL ROSTER
 *
 * getMembers() builds the roster from the wing's own Member.txt, which is scoped
 * to members whose primary unit is in the wing. A region-HQ member is not in it,
 * however many wing duties they hold. The region-wide extract
 * (CONFIG.REGION_CAPWATCH_DATA_FOLDER_ID) has both halves of the join: the
 * member's primary org (Member.txt) and every duty row (DutyPosition.txt).
 *
 * TWO PARTS
 *
 *   previewRegionStaffMembers() - READ-ONLY report of who qualifies and how that
 *   compares with the Manual Members sheet. Works whether or not the flag is on.
 *
 *   mergeDerivedRegionStaff_() - PROVISIONING. Gated by CONFIG.DERIVE_REGION_STAFF
 *   (Script Property TENANT_DERIVE_REGION_STAFF=true; OFF by default). Called by
 *   getMembers(), so derived members flow through every consumer exactly as
 *   Manual Members rows do. Rules:
 *     - They are placed in the region HQ org (/PCR-001 OU), not the duty unit's.
 *     - The Manual Members sheet and the wing's own roster both WIN over a derived
 *       entry, so enabling this changes nothing for anyone already covered.
 *     - Unlike manual rows they are NOT exempt from the lifecycle: they count as
 *       active (getActiveMembers) only while the rule holds, so an account is
 *       suspended on the normal grace clock once the wing duty ends. That is the
 *       whole point of deriving them.
 *     - FAIL-SAFE: if the region extract cannot be read, suspendExpiredMembers()
 *       aborts rather than treat every derived member as lapsed.
 *     - The Level I gate still applies, judged on the REGION extract's records.
 *
 * Changelog:
 *   2.0.0 - Provisioning behind CONFIG.DERIVE_REGION_STAFF: mergeDerivedRegionStaff_,
 *           loadDerivedRegionStaff_, derivedRegionStaffExpirations_,
 *           addRegionLevel1Completions_, addDerivedRegionStaffCapids_.
 *   1.2.0 — Preview flags manual rows the wing's own roster already covers
 *           (wingCoverageForManual_), i.e. rows that are redundant.
 *   1.1.0 — Placement is always the HQ org; preview reports WHY each manual row
 *           was not derived (explainRegionStaffExclusions_).
 *   1.0.0 — Initial: pure matcher (deriveRegionStaffMembers_) + preview.
 */

/** PCR-PCR-001's CAPWATCH ORGID (same value UnitVisitReport.gs uses). */
const REGION_STAFF_HQ_ORGID = '434';

/**
 * PURE. Joins region Member rows to region DutyPosition rows.
 *
 * Qualifies when the member is ACTIVE, of an eligible type, has the region HQ as
 * primary org, and holds >=1 duty at an org in `squadrons` (this wing's units) that
 * is not an excluded holding unit.
 *
 * @param {Array<Array>} memberRows - region Member.txt data rows (no header)
 * @param {Array<Array>} dutyRows - region DutyPosition.txt data rows (no header)
 * @param {Object} squadrons - wing org lookup indexed by orgid ({charter,name,orgPath})
 * @param {{hqOrgId?:string, types:string[], excludedOrgIds?:string[]}} opts
 * @returns {Object} qualifying members indexed by CAPID:
 *   { capid, lastName, firstName, rank, type, duties:[{id,assistant,orgid,charter}],
 *     placementOrgid } — placement is ALWAYS the region HQ org (their home unit), so a
 *     derived member lands in the same /PCR-001 OU the Manual Members rows use today,
 *     not in the OU of the wing unit where they hold the duty.
 */
function deriveRegionStaffMembers_(memberRows, dutyRows, squadrons, opts) {
  const hq = String((opts && opts.hqOrgId) || REGION_STAFF_HQ_ORGID);
  const types = (opts && opts.types) || [];
  const excluded = ((opts && opts.excludedOrgIds) || []).map(String);

  const dutiesByCapid = {};
  (dutyRows || []).forEach(row => {
    const capid = String((row || [])[0] || '').trim();
    const orgid = String((row || [])[7] || '').trim();
    const dutyId = String((row || [])[1] || '').trim();
    if (!capid || !dutyId || !squadrons[orgid] || excluded.indexOf(orgid) > -1) return;
    (dutiesByCapid[capid] = dutiesByCapid[capid] || []).push({
      id: dutyId,
      assistant: row[4] == '1',
      orgid: orgid,
      charter: squadrons[orgid].charter
    });
  });

  const result = {};
  (memberRows || []).forEach(row => {
    const capid = String((row || [])[0] || '').trim();
    if (!capid || !dutiesByCapid[capid]) return;
    if (String(row[11] || '').trim() !== hq) return;
    if (row[24] !== 'ACTIVE' || types.indexOf(row[21]) === -1) return;

    result[capid] = {
      capid: capid,
      lastName: row[2] || '',
      firstName: row[3] || '',
      rank: row[14] || '',
      type: row[21],
      duties: dutiesByCapid[capid],
      placementOrgid: hq
    };
  });
  return result;
}

/**
 * PURE. For CAPIDs the rule did NOT derive, says why — the first failing test in the
 * order the rule applies them. Used to tell a legitimate manual exception from a
 * data gap in the region extract.
 *
 * @param {string[]} capids
 * @param {Array<Array>} memberRows - region Member.txt data rows
 * @param {Array<Array>} dutyRows - region DutyPosition.txt data rows
 * @param {Object} squadrons - wing org lookup (see deriveRegionStaffMembers_)
 * @param {{hqOrgId?:string, types:string[], excludedOrgIds?:string[]}} opts
 * @returns {Object} capid -> reason string, e.g. 'not in region extract',
 *   'primary org 100 (not HQ)', 'status EXPIRED', 'type CADET not eligible',
 *   'no duty in a wing unit (N duties elsewhere)'
 */
function explainRegionStaffExclusions_(capids, memberRows, dutyRows, squadrons, opts) {
  const hq = String((opts && opts.hqOrgId) || REGION_STAFF_HQ_ORGID);
  const types = (opts && opts.types) || [];
  const excluded = ((opts && opts.excludedOrgIds) || []).map(String);

  const memberByCapid = {};
  (memberRows || []).forEach(r => { memberByCapid[String((r || [])[0] || '').trim()] = r; });
  const dutiesByCapid = {};
  (dutyRows || []).forEach(r => {
    const capid = String((r || [])[0] || '').trim();
    (dutiesByCapid[capid] = dutiesByCapid[capid] || []).push(r);
  });

  const reasons = {};
  capids.forEach(capid => {
    const row = memberByCapid[capid];
    if (!row) { reasons[capid] = 'not in region extract'; return; }
    const org = String(row[11] || '').trim();
    if (org !== hq) { reasons[capid] = 'primary org ' + org + ' (not HQ)'; return; }
    if (row[24] !== 'ACTIVE') { reasons[capid] = 'status ' + row[24]; return; }
    if (types.indexOf(row[21]) === -1) { reasons[capid] = 'type ' + row[21] + ' not eligible'; return; }
    const duties = dutiesByCapid[capid] || [];
    const wing = duties.filter(d => {
      const o = String(d[7] || '').trim();
      return squadrons[o] && excluded.indexOf(o) === -1;
    });
    reasons[capid] = wing.length
      ? 'qualifies (rule should have derived it)'
      : 'no duty in a wing unit (' + duties.length + ' duties elsewhere)';
  });
  return reasons;
}

/**
 * PURE. For manual-sheet CAPIDs, says whether the wing's OWN Member.txt already
 * covers them — i.e. whether getMembers() would include them without the sheet.
 * Same eligibility test as shouldProcessMember(): ACTIVE, eligible type, not in an
 * excluded holding unit. A 'covered' row is redundant and safe to delete.
 *
 * @param {string[]} capids
 * @param {Array<Array>} wingMemberRows - this wing's Member.txt data rows
 * @param {{types:string[], excludedOrgIds?:string[]}} opts
 * @returns {Object} capid -> 'covered by wing roster' | 'not in wing file' |
 *   'status X' | 'type X not eligible' | 'in excluded org X'
 */
function wingCoverageForManual_(capids, wingMemberRows, opts) {
  const types = (opts && opts.types) || [];
  const excluded = ((opts && opts.excludedOrgIds) || []).map(String);
  const byCapid = {};
  (wingMemberRows || []).forEach(r => { byCapid[String((r || [])[0] || '').trim()] = r; });

  const out = {};
  capids.forEach(capid => {
    const row = byCapid[capid];
    if (!row) { out[capid] = 'not in wing file'; return; }
    if (row[24] !== 'ACTIVE') { out[capid] = 'status ' + row[24]; return; }
    if (types.indexOf(row[21]) === -1) { out[capid] = 'type ' + row[21] + ' not eligible'; return; }
    const org = String(row[11] || '').trim();
    if (excluded.indexOf(org) > -1) { out[capid] = 'in excluded org ' + org; return; }
    out[capid] = 'covered by wing roster';
  });
  return out;
}

// One load per execution: getMembers() and the lifecycle both ask.
let _derivedRegionStaffCache = undefined;

/** Builds a getMembers()-shaped member object for one derived entry. */
function regionStaffToMember_(entry, memberRow, org) {
  return {
    capsn: entry.capid,
    lastName: entry.lastName,
    firstName: entry.firstName,
    middleName: memberRow[4] || '',
    suffix: memberRow[5] || '',
    orgid: org.orgid,
    group: '',
    charter: org.charter,
    orgName: org.name,
    rank: entry.rank,
    type: entry.type,
    status: memberRow[24],
    joined: memberRow[15],
    modified: memberRow[19],
    orgPath: org.orgPath,
    email: null,
    derivedFromRegion: true,
    dutyPositions: [],
    dutyPositionIds: [],
    dutyPositionIdsAndLevel: []
  };
}

/**
 * Loads the derived members for this execution (cached).
 *
 * @param {Object} [squadrons] - getSquadrons()
 * @returns {Object|null} { derived, memberRowsByCapid, org, dutyRows, contactRows };
 *   an empty `derived` when the flag is off; or NULL when the flag is on but the
 *   region extract (or the HQ org's OU path) cannot be read. Callers must treat
 *   null as "unknown", never as "nobody".
 */
function loadDerivedRegionStaff_(squadrons) {
  if (!CONFIG.DERIVE_REGION_STAFF) {
    return { derived: {}, memberRowsByCapid: {}, dutyRows: [], contactRows: [] };
  }
  if (_derivedRegionStaffCache !== undefined) return _derivedRegionStaffCache;

  const folderId = String(CONFIG.REGION_CAPWATCH_DATA_FOLDER_ID || '').trim();
  const memberRows = folderId ? parseCapwatchFromFolder_(folderId, 'Member') : [];
  if (!memberRows.length) {
    Logger.error('Derived region staff: region extract unreadable', { folderId: folderId });
    return (_derivedRegionStaffCache = null);
  }
  squadrons = squadrons || getSquadrons();
  const dutyRows = parseCapwatchFromFolder_(folderId, 'DutyPosition');

  let orgPath = '';
  parseFile('OrgPaths').forEach(r => {
    if (String(r[0] || '').trim() === REGION_STAFF_HQ_ORGID) orgPath = String(r[1] || '').trim();
  });
  if (!orgPath) {
    Logger.error('Derived region staff: no OrgPaths entry for the region HQ org',
      { orgid: REGION_STAFF_HQ_ORGID });
    return (_derivedRegionStaffCache = null);
  }
  const org = { orgid: REGION_STAFF_HQ_ORGID, orgPath: orgPath, name: 'Region HQ', charter: '' };
  parseCapwatchFromFolder_(folderId, 'Organization').forEach(r => {
    if (String(r[0] || '').trim() !== REGION_STAFF_HQ_ORGID) return;
    org.name = String(r[5] || '').trim() || org.name;
    try { org.charter = Utilities.formatString('%s-%s-%03d', r[1], r[2], r[3]); } catch (ignored) {}
  });

  const derived = deriveRegionStaffMembers_(memberRows, dutyRows, squadrons,
    { types: CONFIG.MEMBER_TYPES.ACTIVE, excludedOrgIds: CONFIG.EXCLUDED_ORG_IDS });
  const memberRowsByCapid = {};
  memberRows.forEach(r => { memberRowsByCapid[String((r || [])[0] || '').trim()] = r; });

  return (_derivedRegionStaffCache = {
    derived: derived,
    memberRowsByCapid: memberRowsByCapid,
    org: org,
    dutyRows: dutyRows,
    contactRows: parseCapwatchFromFolder_(folderId, 'MbrContact')
  });
}

/**
 * Adds derived region staff to a getMembers() roster. Called before the Manual
 * Members merge, so a manual row for the same CAPID overwrites the derived one.
 * Anyone already on the roster (from the wing's own file) is left alone.
 *
 * @param {Object} members - roster indexed by CAPID (mutated)
 * @param {Object} squadrons
 * @returns {number} members added
 */
function mergeDerivedRegionStaff_(members, squadrons) {
  const loaded = loadDerivedRegionStaff_(squadrons);
  if (!loaded || !Object.keys(loaded.derived).length) return 0;

  const added = {};
  Object.keys(loaded.derived).forEach(capid => {
    if (members[capid]) return;
    added[capid] = regionStaffToMember_(loaded.derived[capid], loaded.memberRowsByCapid[capid], loaded.org);
  });
  const ids = Object.keys(added);
  if (!ids.length) return 0;

  // Same readers the wing roster uses, fed from the region extract and scoped to
  // the added members only (both skip CAPIDs that are not in the object they get).
  addContactInfo(added, loaded.contactRows);
  addDutyPositions(added,
    loaded.dutyRows.filter(r => squadrons[String((r || [])[7] || '').trim()]), squadrons);

  ids.forEach(capid => { members[capid] = added[capid]; });
  Logger.info('Derived region staff merged', { count: ids.length });
  return ids.length;
}

/**
 * CAPID -> expiration for derived members, for getActiveMembers(). Empty when the
 * flag is off OR the region extract is unreadable (use regionExtractReadable_()
 * to tell those apart before suspending anyone).
 */
function derivedRegionStaffExpirations_() {
  const loaded = loadDerivedRegionStaff_();
  const out = {};
  if (!loaded) return out;
  Object.keys(loaded.derived).forEach(capid => {
    out[capid] = (loaded.memberRowsByCapid[capid] || [])[16];
  });
  return out;
}

/** False only when the flag is ON and the region extract could not be read. */
function regionExtractReadable_() {
  return !CONFIG.DERIVE_REGION_STAFF || loadDerivedRegionStaff_() !== null;
}

/** Adds region-extract Level I completions to a Set (derived members are not in the wing's files). */
function addRegionLevel1Completions_(completed) {
  if (!CONFIG.DERIVE_REGION_STAFF) return completed;
  const folderId = String(CONFIG.REGION_CAPWATCH_DATA_FOLDER_ID || '').trim();
  if (!folderId) return completed;
  parseCapwatchFromFolder_(folderId, 'MbrAchievements').forEach(r => {
    if (String(r[1]) === '96' && r[2] === 'ACTIVE') completed.add(String(r[0]));
  });
  parseCapwatchFromFolder_(folderId, 'SeniorLevel').forEach(r => {
    if (String(r[1]) === 'LV1') completed.add(String(r[0]));
  });
  return completed;
}

/** Adds derived CAPIDs to an allow-list object ({capid: 1}), e.g. the audit's. */
function addDerivedRegionStaffCapids_(allowList) {
  Object.keys(derivedRegionStaffExpirations_()).forEach(c => { allowList[c] = 1; });
  return allowList;
}

/**
 * READ-ONLY. Logs who would be derived and how they compare with what is already
 * on the Manual Members sheet / the wing's own roster. Run by hand.
 *
 * @returns {{derived:Object, alreadyManual:string[], newToRoster:string[],
 *            manualNotDerived:string[], manualNotDerivedWhy:Object, manualWingCoverage:Object,
 *            inWingRoster:string[]}}
 */
function previewRegionStaffMembers() {
  const folderId = String(CONFIG.REGION_CAPWATCH_DATA_FOLDER_ID || '').trim();
  if (!folderId) {
    Logger.warn('Region staff preview needs TENANT_REGION_CAPWATCH_DATA_FOLDER_ID');
    return null;
  }

  const squadrons = getSquadrons();
  const regionMembers = parseCapwatchFromFolder_(folderId, 'Member');
  const regionDuties = parseCapwatchFromFolder_(folderId, 'DutyPosition');
  const ruleOpts = { types: CONFIG.MEMBER_TYPES.ACTIVE, excludedOrgIds: CONFIG.EXCLUDED_ORG_IDS };
  const derived = deriveRegionStaffMembers_(regionMembers, regionDuties, squadrons, ruleOpts);

  const manual = loadManualMembers(squadrons);
  const wingMemberRows = parseFile('Member');
  const wingRoster = {};
  wingMemberRows.forEach(r => { wingRoster[String(r[0])] = true; });

  const derivedIds = Object.keys(derived);
  const out = {
    derived: derived,
    alreadyManual: derivedIds.filter(id => manual[id]),
    newToRoster: derivedIds.filter(id => !manual[id]),
    // Manual rows the rule would NOT cover: genuine exceptions (NHQ, no wing duty...).
    manualNotDerived: Object.keys(manual).filter(id => !derived[id]),
    // Should be empty: a region-HQ primary member should not be in the wing's own file.
    inWingRoster: derivedIds.filter(id => wingRoster[id])
  };
  out.manualNotDerivedWhy = explainRegionStaffExclusions_(
    out.manualNotDerived, regionMembers, regionDuties, squadrons, ruleOpts);
  // Every manual row, not only the underived ones: a 'covered' row is redundant.
  out.manualWingCoverage = wingCoverageForManual_(Object.keys(manual), wingMemberRows, ruleOpts);

  // CAPIDs only: member names stay out of the logs.
  Logger.info('Region staff preview', {
    derivedCount: derivedIds.length,
    alreadyManual: out.alreadyManual.length,
    newToRoster: out.newToRoster,
    manualNotDerived: out.manualNotDerivedWhy,
    inWingRoster: out.inWingRoster,
    manualWingCoverage: out.manualWingCoverage
  });
  return out;
}
