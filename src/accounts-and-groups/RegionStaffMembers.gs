/**
 * RegionStaffMembers.gs
 * Version: 1.0.0
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
 * THIS FILE IS READ-ONLY. It only reports who would qualify and how that compares
 * with the Manual Members sheet. Nothing is provisioned from it yet.
 *
 * Entry point: previewRegionStaffMembers()
 *
 * Changelog:
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
 *     placementOrgid } — placement is the unit of the member's first PRIMARY duty
 *     (assistant duties only if they have no primary one), lowest charter on ties.
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

    const duties = dutiesByCapid[capid];
    const ranked = duties.slice().sort((a, b) =>
      (a.assistant - b.assistant) || String(a.charter).localeCompare(String(b.charter)));
    result[capid] = {
      capid: capid,
      lastName: row[2] || '',
      firstName: row[3] || '',
      rank: row[14] || '',
      type: row[21],
      duties: duties,
      placementOrgid: ranked[0].orgid
    };
  });
  return result;
}

/**
 * READ-ONLY. Logs who would be derived and how they compare with what is already
 * on the Manual Members sheet / the wing's own roster. Run by hand.
 *
 * @returns {{derived:Object, alreadyManual:string[], newToRoster:string[],
 *            manualNotDerived:string[], inWingRoster:string[]}}
 */
function previewRegionStaffMembers() {
  const folderId = String(CONFIG.REGION_CAPWATCH_DATA_FOLDER_ID || '').trim();
  if (!folderId) {
    Logger.warn('Region staff preview needs TENANT_REGION_CAPWATCH_DATA_FOLDER_ID');
    return null;
  }

  const squadrons = getSquadrons();
  const derived = deriveRegionStaffMembers_(
    parseCapwatchFromFolder_(folderId, 'Member'),
    parseCapwatchFromFolder_(folderId, 'DutyPosition'),
    squadrons,
    { types: CONFIG.MEMBER_TYPES.ACTIVE, excludedOrgIds: CONFIG.EXCLUDED_ORG_IDS });

  const manual = loadManualMembers(squadrons);
  const wingRoster = {};
  parseFile('Member').forEach(r => { wingRoster[String(r[0])] = true; });

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

  // CAPIDs only: member names stay out of the logs.
  Logger.info('Region staff preview', {
    derivedCount: derivedIds.length,
    alreadyManual: out.alreadyManual.length,
    newToRoster: out.newToRoster,
    manualNotDerived: out.manualNotDerived,
    inWingRoster: out.inWingRoster
  });
  return out;
}
