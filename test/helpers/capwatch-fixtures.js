/**
 * CAPWATCH fixtures.
 *
 * The Member.txt header here is the real one, copied verbatim from a live
 * extract on 2026-07-15. That matters: LSCodeNotify.gs resolves its columns by
 * header name, so a fixture with an invented header would test nothing — it
 * would pass whatever order the fake happened to use. Keep this in sync with the
 * live file if CAP ever reshapes the extract; the module fails loudly on a
 * missing column, and this fixture is what proves that path.
 */

const MEMBER_HEADER = [
  'CAPID', 'SSN', 'NameLast', 'NameFirst', 'NameMiddle', 'NameSuffix', 'Gender', 'DOB',
  'Profession', 'EducationLevel', 'Citizen', 'ORGID', 'Wing', 'Unit', 'Rank', 'Joined',
  'Expiration', 'OrgJoined', 'UsrID', 'DateMod', 'LSCode', 'Type', 'RankDate', 'Region',
  'MbrStatus', 'PicStatus', 'PicDate', 'CdtWaiver', 'Ethnicity'
];

const COL = {};
MEMBER_HEADER.forEach((name, i) => { COL[name] = i; });

/**
 * Builds one Member.txt row.
 *
 * @param {Object} spec - { capid, orgid, lscode, type, status }
 * @returns {string} A quoted CSV line
 */
function memberRow(spec) {
  const cells = new Array(MEMBER_HEADER.length).fill('');
  cells[COL.CAPID] = spec.capid;
  cells[COL.NameLast] = 'Last' + spec.capid;
  cells[COL.NameFirst] = 'First' + spec.capid;
  cells[COL.Gender] = 'MALE';
  cells[COL.DOB] = '1/01/1970';
  cells[COL.Citizen] = 'US Citizen';
  cells[COL.ORGID] = spec.orgid;
  cells[COL.Wing] = 'CA';
  cells[COL.Unit] = '070';
  cells[COL.Rank] = 'Capt';
  cells[COL.Joined] = '01/01/2000';
  cells[COL.Expiration] = '12/31/2026';
  cells[COL.OrgJoined] = '01/01/2000';
  cells[COL.UsrID] = 'someuser';
  cells[COL.DateMod] = '01/01/2026';
  cells[COL.LSCode] = spec.lscode === undefined ? '' : spec.lscode;
  cells[COL.Type] = spec.type || 'SENIOR';
  cells[COL.RankDate] = '01/01/2010';
  cells[COL.Region] = 'PCR';
  cells[COL.MbrStatus] = spec.status || 'ACTIVE';
  cells[COL.PicStatus] = 'NONE';
  cells[COL.PicDate] = '01/01/1900';
  cells[COL.Ethnicity] = 'WHITE';
  return cells.map(v => '"' + String(v) + '"').join(',');
}

/**
 * @param {Array<Object>} specs - Member specs (see memberRow)
 * @returns {string} A complete Member.txt
 */
function memberFile(specs) {
  return MEMBER_HEADER.join(',') + '\n' + specs.map(memberRow).join('\n') + '\n';
}

/** A header-only Member.txt, as a truncated extract would look. */
function emptyMemberFile() {
  return MEMBER_HEADER.join(',') + '\n';
}

// Two squadrons, each with a commander who has a reachable email.
const SQUADRONS = {
  '100': { orgid: '100', name: 'Alpha Squadron', charter: 'PCR-CA-070' },
  '200': { orgid: '200', name: 'Bravo Squadron', charter: 'PCR-CA-080' }
};

// Commanders.txt column order, verified against a live extract:
// ORGID,Region,Wing,Unit,CAPID,DateAsg,UsrID,DateMod,NameLast,NameFirst,NameMiddle,NameSuffix,Rank
const COMMANDER_ROWS = [
  ['100', 'PCR', 'CA', '070', '900001', '01/01/2025', 'usr', '01/01/2025', 'Alpha', 'Ada', '', '', 'Maj'],
  ['200', 'PCR', 'CA', '080', '900002', '01/01/2025', 'usr', '01/01/2025', 'Bravo', 'Ben', '', '', 'Maj']
];

const COMMANDER_EMAILS = {
  '900001': 'ada.alpha@cawgcap.org',
  '900002': 'ben.bravo@cawgcap.org'
};

module.exports = {
  MEMBER_HEADER, COL, memberRow, memberFile, emptyMemberFile,
  SQUADRONS, COMMANDER_ROWS, COMMANDER_EMAILS
};
