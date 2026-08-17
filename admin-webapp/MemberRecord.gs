/***********************************************
 * File: MemberRecord.gs
 * Description: Reads one member out of the tenant's CAPWATCH extract. Read-only
 * throughout — nothing in this file writes anywhere.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-08-16
 * Changes: 1.0.0 — initial version.
 ***********************************************/

/**
 * WHY THIS FILE EXISTS INSTEAD OF A CALL TO getMembers()
 *
 * src/accounts-and-groups/UpdateMembers.gs builds EVERY member in the wing —
 * five CAPWATCH files, thousands of objects, manual-member merge, manager
 * lookups — because a wing-wide push needs all of them. A help desk looking up
 * one member needs one, so this reads the same files and assembles the fields
 * the app actually uses. Same approach, and much of the same code, as
 * signature-webapp/MemberRecord.gs.
 *
 * THE FIELDS ARE NOT ARBITRARY. member.email and member.secondaryEmail are the
 * two addresses the welcome email is mailed to, and they are derived here
 * EXACTLY as addContactInfo() derives them in src/ — PRIMARY sets email (unless
 * DoNotContact), SECONDARY sets secondaryEmail. Getting that wrong would not
 * show up as an error; it would show up as credentials quietly mailed somewhere
 * else. test/AdminWebApp.test.js pins it.
 *
 * A MANUAL MEMBER (the Manual Members sheet in src/) is invisible here. That is
 * a deliberate limitation, not an oversight: reading it would mean a spreadsheet
 * dependency and a second merge path for a population of a handful. A manual
 * member with no CAPWATCH row reads as "no CAPWATCH record", and the account
 * card still shows their Workspace account in full.
 */

/** Per-execution memo. A web app request is short-lived, so this is enough. */
const ADM_FILE_MEMO = {};

/**
 * Parses one CAPWATCH file out of the tenant's extract folder, header dropped —
 * the same contract as parseFile() in src/utils.gs.
 *
 * @param {string} baseName - e.g. 'Member' for Member.txt
 * @returns {Array<Array<string>>} rows, or [] when the file is absent or empty
 */
function admParseCapwatchFile_(baseName) {
  const folderId = ADMIN_CONFIG.CAPWATCH_DATA_FOLDER_ID;
  const key = folderId + '/' + baseName;
  if (ADM_FILE_MEMO[key]) return ADM_FILE_MEMO[key];

  let rows = [];
  try {
    const files = DriveApp.getFolderById(folderId).getFilesByName(baseName + '.txt');
    if (!files.hasNext()) {
      Logger.warn('CAPWATCH file not found', { fileName: baseName + '.txt', folderId: folderId });
    } else {
      const content = files.next().getBlob().getDataAsString();
      if (content) rows = Utilities.parseCsv(content).slice(1);
      else Logger.warn('CAPWATCH file is empty', { fileName: baseName + '.txt' });
    }
  } catch (err) {
    // The extract being unreadable does not stop the app: every directory-side
    // panel still works, and the member card says CAPWATCH could not be read
    // rather than "no such member", which would be a lie with consequences.
    Logger.warn('CAPWATCH folder could not be read; continuing without it', {
      folderId: folderId, fileName: baseName + '.txt', errorMessage: err.message
    });
  }

  ADM_FILE_MEMO[key] = rows;
  return rows;
}

/**
 * Wing org map: orgid -> { name, charter, scope }. Mirrors getSquadrons()'s wing
 * filter and its charter format, so a charter shown here is the one src/ writes
 * to the account's costCenter.
 *
 * @returns {Object}
 */
function admOrgMap_() {
  const rows = admParseCapwatchFile_('Organization');
  const wing = String(ADMIN_CONFIG.WING || '').trim().toUpperCase();
  const orgs = {};
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    const orgid = String(row[0] || '').trim();
    if (!orgid) continue;
    if (String(row[2] || '').trim().toUpperCase() !== wing) continue;

    const unit = parseInt(String(row[3] || '').trim(), 10);
    orgs[orgid] = {
      name: String(row[5] || '').trim(),
      charter: isNaN(unit)
        ? ''
        : Utilities.formatString('%s-%s-%03d', String(row[1] || '').trim(), String(row[2] || '').trim(), unit),
      scope: String(row[9] || '').trim()
    };
  }
  return orgs;
}

/**
 * The contact addresses CAPWATCH holds for a member, derived the way
 * addContactInfo() derives them in src/accounts-and-groups/UpdateMembers.gs.
 *
 * The DoNotContact rule is the subtle one and it is copied verbatim: a PRIMARY
 * address flagged DoNotContact is NOT used as the member's contact address, but
 * it is still reported (primaryEmailValue) because it remains a recovery
 * candidate. Members routinely list a personal address as PRIMARY.
 *
 * @param {string} capid
 * @returns {{email: string|null, secondaryEmail: string|null,
 *   primaryEmailValue: string|null, primaryEmailDNC: boolean,
 *   secondaryEmailDNC: boolean}}
 */
function admContactEmails_(capid) {
  const rows = admParseCapwatchFile_('MbrContact');
  const out = {
    email: null, secondaryEmail: null,
    primaryEmailValue: null, primaryEmailDNC: false, secondaryEmailDNC: false
  };

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    if (String(row[0] || '').trim() !== capid) continue;
    if (String(row[1] || '').toUpperCase() !== 'EMAIL') continue;

    const value = String(row[3] || '').trim();
    if (!value || value.indexOf('@') === -1) continue;
    const priority = String(row[2] || '').toUpperCase();
    const doNotContact = String(row[6] || '').toUpperCase() === 'TRUE';

    if (priority === 'PRIMARY') {
      if (!doNotContact) out.email = value;
      out.primaryEmailValue = value;
      out.primaryEmailDNC = doNotContact;
    } else if (priority === 'SECONDARY') {
      out.secondaryEmail = value;
      out.secondaryEmailDNC = doNotContact;
    }
  }
  return out;
}

/**
 * One member's CAPWATCH record.
 *
 * Unlike the signature app this does NOT filter to ACTIVE members: a help desk
 * is most often called about someone whose membership just lapsed, and "expired
 * on 2026-06-30" is the answer to their question. The status is returned and the
 * page shows it; the actions that must not run for an expired member say so
 * themselves.
 *
 * @param {string} capid
 * @returns {Object|null} null when no member carries that CAPID
 */
function admBuildMemberRecord_(capid) {
  const wanted = String(capid || '').trim();
  if (!ADM_CAPID_RE.test(wanted)) return null;

  const rows = admParseCapwatchFile_('Member');
  let row = null;
  for (let i = 0; i < rows.length; i++) {
    if (String((rows[i] || [])[0] || '').trim() === wanted) { row = rows[i]; break; }
  }
  if (!row) return null;

  const orgs = admOrgMap_();
  const homeOrg = orgs[String(row[11] || '').trim()];
  const contacts = admContactEmails_(wanted);

  return {
    capsn: wanted,
    rank: String(row[14] || '').trim(),
    firstName: String(row[3] || '').trim(),
    middleName: String(row[4] || '').trim(),
    lastName: String(row[2] || '').trim(),
    suffix: String(row[5] || '').trim(),
    type: String(row[21] || '').trim(),
    status: String(row[24] || '').trim(),
    expiration: String(row[16] || '').trim(),
    orgName: homeOrg ? homeOrg.name : '',
    charter: homeOrg ? homeOrg.charter : '',
    email: contacts.email,
    secondaryEmail: contacts.secondaryEmail,
    primaryEmailValue: contacts.primaryEmailValue,
    primaryEmailDNC: contacts.primaryEmailDNC,
    secondaryEmailDNC: contacts.secondaryEmailDNC
  };
}

/**
 * CAPIDs matching a name fragment, for the admin who has a name and no number.
 *
 * Scans Member.txt rather than the directory on purpose: a member with no
 * Workspace account at all is exactly the case a help desk is called about, and
 * a directory search cannot find them.
 *
 * It also means the results include members this tenant does not provision —
 * every cadet in the wing, on a seniors tenant. They are RETURNED rather than
 * filtered, and flagged: a help desk that searched a name and got nothing back
 * would conclude the member does not exist, when the truth is that they are on
 * the other Workspace. `offTenant` is what the page turns into that sentence.
 *
 * @param {string} query - a fragment of a first or last name
 * @param {number} limit
 * @returns {Array<{capid: string, name: string, type: string, status: string,
 *   offTenant: boolean}>}
 */
function admSearchMembersByName_(query, limit) {
  const needle = String(query || '').trim().toLowerCase();
  const out = [];
  if (needle.length < 2) return out;

  const rows = admParseCapwatchFile_('Member');
  for (let i = 0; i < rows.length && out.length < limit; i++) {
    const row = rows[i] || [];
    const last = String(row[2] || '').trim();
    const first = String(row[3] || '').trim();
    const full = (first + ' ' + last).toLowerCase();
    if (full.indexOf(needle) === -1 && (last + ', ' + first).toLowerCase().indexOf(needle) === -1) continue;

    const type = String(row[21] || '').trim();
    out.push({
      capid: String(row[0] || '').trim(),
      name: [String(row[14] || '').trim(), first, last].filter(Boolean).join(' '),
      type: type,
      status: String(row[24] || '').trim(),
      offTenant: !admTenantProvisionsType_(type)
    });
  }
  return out;
}

/** 'Maj Dana Okonkwo', from a record built above. */
function admMemberDisplayName_(member) {
  if (!member) return '';
  return [member.rank, member.firstName, member.lastName, member.suffix]
    .filter(Boolean).join(' ');
}
