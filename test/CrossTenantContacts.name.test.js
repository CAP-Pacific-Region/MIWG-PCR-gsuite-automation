/**
 * CrossTenantContacts.gs — the directory display name and the GAL sort key, in
 * isolation. These pin the v0.2.0 sort fix: Domain Shared Contacts sort in the
 * GAL by gd:givenName (not fullName, and not familyName the way directory users
 * do), so the whole "Last Suffix, First M Grade" string must go into givenName
 * for the list to sort by last name. A regression here silently reverts the sort
 * to first-name order, so assert it directly.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker, Utilities } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'cross-tenant-contacts', 'CrossTenantContacts.gs');
const { section, check, done } = makeChecker();

// xtHash_ reaches for Utilities.computeDigest; the digest value is irrelevant to
// these assertions, so a deterministic stub is enough.
const U = Object.assign({}, Utilities, {
  DigestAlgorithm: { SHA_256: 'SHA_256' },
  Charset: { UTF_8: 'UTF_8' },
  computeDigest: () => [0, 1, 2, 3]
});

const m = loadModule(MODULE, {
  Logger: makeLogger().logger,
  Utilities: U,
  toTitleCase: s => String(s || '')
}, ['xtDisplayName_', 'xtBuildContactXml_', 'xtMakeContact_']);

// ---------------------------------------------------------------------------
section('1. xtDisplayName_ — "Last Suffix, First M Grade", matching native accounts');
{
  check('full name with suffix, middle, grade',
    m.xtDisplayName_({ familyName: 'Smith', suffix: 'Jr', givenName: 'John', middleName: 'Quincy', grade: 'C/Amn' }),
    'Smith Jr, John Q C/Amn');
  check('middle initial only (first letter of middle name)',
    m.xtDisplayName_({ familyName: 'Doe', givenName: 'Jane', middleName: 'Marie' }),
    'Doe, Jane M');
  check('no middle / no suffix / no grade',
    m.xtDisplayName_({ familyName: 'Doe', givenName: 'Jane' }),
    'Doe, Jane');
  check('grade but no middle',
    m.xtDisplayName_({ familyName: 'Roe', givenName: 'Rick', grade: 'Lt Col' }),
    'Roe, Rick Lt Col');
}

// ---------------------------------------------------------------------------
section('2. xtBuildContactXml_ — sort key (givenName) leads with the last name');
{
  const c = {
    capid: '123456', email: 'a@b.org',
    displayName: 'Smith, John Q C/Amn',
    dutyTitle: 'No Duty Assignment', department: 'Alpha Squadron',
    notes: 'CADET, PCR-CA-070', syncHash: 'deadbeef'
  };
  const xml = m.xtBuildContactXml_(c, { wing: 'CA' });

  check('givenName carries the full last-name-first display string',
    /<gd:givenName>Smith, John Q C\/Amn<\/gd:givenName>/.test(xml), true);
  check('fullName is the same display string',
    /<gd:fullName>Smith, John Q C\/Amn<\/gd:fullName>/.test(xml), true);
  check('no separate familyName element (would break the card First/Last split)',
    /<gd:familyName>/.test(xml), false);
  check('title is the display string',
    /<title>Smith, John Q C\/Amn<\/title>/.test(xml), true);
  check('managed marker is the wing',
    /<gd:orgName>CA<\/gd:orgName>/.test(xml), true);
}

// ---------------------------------------------------------------------------
section('3. xtMakeContact_ — builds the display from CAPWATCH fields end to end');
{
  const member = {
    capsn: '654321', firstName: 'Mary', middleName: 'Ann', lastName: 'OConnor',
    suffix: '', rank: 'C/CMSgt', orgName: 'Bravo Squadron', charter: 'PCR-CA-071',
    dutyPositions: []
  };
  const contact = m.xtMakeContact_(member, { email: 'mary@x.org', source: 'capwatch' }, { peerLabel: 'CADET' });
  check('displayName folds grade + middle initial', contact.displayName, 'OConnor, Mary A C/CMSgt');
  check('carries the resolved email', contact.email, 'mary@x.org');
  check('capid is the CAPWATCH serial', contact.capid, '654321');
  const xml = m.xtBuildContactXml_(contact, { wing: 'CA' });
  check('sort key leads with last name',
    /<gd:givenName>OConnor, Mary A C\/CMSgt<\/gd:givenName>/.test(xml), true);
}

done();
