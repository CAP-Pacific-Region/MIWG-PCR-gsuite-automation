/***********************************************
 * File: MemberRecord.gs
 * Description: Resolves the signed-in caller to their CAPWATCH record — the only
 * source of anything that appears in a signature. Read-only throughout.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-07-28
 * Changes: 1.0.0 — initial version.
 ***********************************************/

/**
 * WHY THIS FILE EXISTS INSTEAD OF A CALL TO getMembers()
 *
 * src/accounts-and-groups/UpdateMembers.gs builds EVERY member in the wing — five
 * CAPWATCH files, thousands of objects, manual-member merge, manager lookups —
 * because a wing-wide push needs all of them. A member waiting on a web page needs
 * exactly one, so this reads the same files and assembles the same fields for a
 * single CAPID.
 *
 * The fields it produces are precisely those SignatureTemplate.gs reads, and each
 * is derived the same way the main project derives it. The three places that
 * matters most, all of them load-bearing for parity:
 *
 *   1. The org map is filtered to THIS WING, exactly as getSquadrons() does. A
 *      duty held at an org outside the wing therefore has no orgName and falls
 *      back to the member's home unit — which looks like a bug until you notice
 *      the main project does the same, and that matching it is the point.
 *   2. A cadet NEVER gets a phone number, by the rule in addContactInfo(): cadet
 *      phone numbers are not published to the directory or to signatures. On a
 *      cadet tenant the "include my phone" option is therefore moot, and the UI
 *      says so rather than offering a toggle that does nothing.
 *   3. DoNotContact contact rows are skipped, and only a CELL/MOBILE row can
 *      supply the number, PRIMARY beating SECONDARY.
 *
 * getSquadrons() also synthesizes an AEM org from CONFIG.SPECIAL_ORGS.AEM_UNIT.
 * That property is empty on every tenant (an inert entry keyed by ''), so it is
 * deliberately not ported; if a tenant ever runs AEM automation, port it here too.
 */

/** Per-execution memo. A web app request is short-lived, so this is just enough. */
const SIG_FILE_MEMO = {};

/**
 * Parses one CAPWATCH file out of the tenant's extract folder, header dropped —
 * the same contract as parseFile() in src/utils.gs.
 *
 * @param {string} baseName - e.g. 'Member' for Member.txt
 * @returns {Array<Array<string>>} rows, or [] when the file is absent or empty
 */
function sigParseCapwatchFile_(baseName) {
  if (SIG_FILE_MEMO[baseName]) return SIG_FILE_MEMO[baseName];

  const folderId = SIG_CONFIG.CAPWATCH_DATA_FOLDER_ID;
  if (!folderId) {
    throw new Error('This tenant is not configured for signatures yet ' +
      '(TENANT_CAPWATCH_DATA_FOLDER_ID is unset).');
  }

  const files = DriveApp.getFolderById(folderId).getFilesByName(baseName + '.txt');
  if (!files.hasNext()) {
    Logger.warn('CAPWATCH file not found', { fileName: baseName + '.txt', folderId: folderId });
    SIG_FILE_MEMO[baseName] = [];
    return SIG_FILE_MEMO[baseName];
  }

  const content = files.next().getBlob().getDataAsString();
  if (!content) {
    Logger.warn('CAPWATCH file is empty', { fileName: baseName + '.txt' });
    SIG_FILE_MEMO[baseName] = [];
    return SIG_FILE_MEMO[baseName];
  }

  SIG_FILE_MEMO[baseName] = Utilities.parseCsv(content).slice(1);
  return SIG_FILE_MEMO[baseName];
}

/**
 * CAPIDs marking a Directory account as live, ignoring the retired-twin marker.
 * Ported from provisioningCapidsFromUser_() in DuplicateAccountGuard.gs.
 *
 * @param {Object} user - AdminDirectory user resource (projection: full)
 * @returns {Array<string>}
 */
function sigCapidsFromUser_(user) {
  const found = {};
  ((user && user.externalIds) || []).forEach(function (id) {
    // 'duplicate_retired_capid' parks a CAPID on a dead twin; it is not an account
    // the member uses, so it must not resolve here.
    if (id && id.type === 'custom' && id.customType === 'duplicate_retired_capid') return;
    const v = String(id && id.value != null ? id.value : '').trim();
    if (SIG_CAPID_RE.test(v)) found[v] = true;
  });
  const emp = String(user && user.employeeId != null ? user.employeeId : '').trim();
  if (SIG_CAPID_RE.test(emp)) found[emp] = true;
  return Object.keys(found);
}

/**
 * The CAPID on the caller's own account.
 *
 * @param {string} email - the authenticated caller, never a client-supplied value
 * @returns {string} '' when the account carries no usable CAPID
 * @throws {Error} when the account cannot be read at all
 */
function sigCapidForAccount_(email) {
  let user;
  try {
    user = AdminDirectory.Users.get(email, { projection: 'full' });
  } catch (err) {
    Logger.error('Could not read the caller\'s directory account', {
      user: email, errorMessage: err.message
    });
    throw new Error('Your account could not be read from the directory. ' +
      sigSupportSentence_());
  }

  const capids = sigCapidsFromUser_(user);
  if (capids.length > 1) {
    // Two live CAPIDs on one account is a data problem no member can fix, and
    // guessing would put someone else's grade and duty on their signature.
    Logger.warn('Account carries more than one CAPID; refusing to guess', {
      user: email, capids: capids
    });
    throw new Error('Your account is linked to more than one CAPID, so we cannot ' +
      'tell which record is yours. ' + sigSupportSentence_());
  }
  return capids[0] || '';
}

/** Wing org map: orgid -> { name, scope }. Mirrors getSquadrons()'s wing filter. */
function sigOrgMap_() {
  const rows = sigParseCapwatchFile_('Organization');
  const orgs = {};
  const wing = String(SIG_CONFIG.WING || '').trim().toUpperCase();
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    const orgid = String(row[0] || '').trim();
    if (!orgid) continue;
    if (String(row[2] || '').trim().toUpperCase() !== wing) continue;
    orgs[orgid] = { name: String(row[5] || '').trim(), scope: String(row[9] || '').trim() };
  }
  return orgs;
}

/**
 * The member's directory phone, or '' — the number the signature would print.
 * Ported from the phone half of addContactInfo(); see this file's header for the
 * three rules that matter.
 *
 * @param {string} capid
 * @param {string} memberType - CAPWATCH Member.txt `Type`; 'CADET' means no phone
 * @returns {string} '+15551234567', or ''
 */
function sigDirectoryPhone_(capid, memberType) {
  if (String(memberType || '').trim().toUpperCase() === 'CADET') return '';

  const rows = sigParseCapwatchFile_('MbrContact');
  let phone = '';
  let haveOne = false;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    if (String(row[0] || '').trim() !== capid) continue;

    const type = String(row[1] || '').toUpperCase();
    const priority = String(row[2] || '').toUpperCase();
    const contact = String(row[3] || '').trim();
    const doNotContact = String(row[6] || '').toUpperCase() === 'TRUE';
    if (doNotContact) continue;

    const digits = contact.replace(/\D/g, '');
    if (digits.length < 10) continue;

    const isMemberCell = type.indexOf('PARENT') === -1 &&
      (type === 'CELL PHONE' || type === 'MOBILE PHONE' || type.indexOf('CELL') !== -1);
    if (!isMemberCell) continue;

    // The first cell row wins, and any later PRIMARY row displaces it — the exact
    // precedence addContactInfo() applies through its phoneSource marker.
    if (!haveOne || priority === 'PRIMARY') {
      phone = '+1' + digits.slice(-10);
      haveOne = true;
    }
  }
  return phone;
}

/** Duty positions for one CAPID, senior and cadet files alike. */
function sigDutyPositions_(capid, orgs) {
  const out = [];
  ['DutyPosition', 'CadetDutyPositions'].forEach(function (fileName) {
    const rows = sigParseCapwatchFile_(fileName);
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i] || [];
      if (String(row[0] || '').trim() !== capid) continue;

      const org = orgs[String(row[7] || '').trim()];
      out.push({
        id: String(row[1] || '').trim(),
        level: String(row[3] || '').trim(),
        assistant: String(row[4] || '').trim() === '1',
        // The org the duty is held AT, which is not necessarily the member's home
        // unit. Blank when the org is outside this wing — see the header note.
        orgName: org ? org.name : '',
        orgScope: org ? org.scope : ''
      });
    }
  });
  return out;
}

/**
 * The caller's CAPWATCH record, reduced to what a signature needs.
 *
 * @param {string} capid
 * @returns {Object|null} null when no ACTIVE member carries that CAPID
 */
function sigBuildMemberRecord_(capid) {
  const rows = sigParseCapwatchFile_('Member');
  let row = null;
  for (let i = 0; i < rows.length; i++) {
    if (String((rows[i] || [])[0] || '').trim() === capid) { row = rows[i]; break; }
  }
  if (!row) return null;
  if (String(row[24] || '').trim().toUpperCase() !== 'ACTIVE') return null;

  const orgs = sigOrgMap_();
  const homeOrg = orgs[String(row[11] || '').trim()];
  const type = String(row[21] || '').trim();

  return {
    capsn: capid,
    rank: String(row[14] || '').trim(),
    firstName: String(row[3] || '').trim(),
    middleName: String(row[4] || '').trim(),
    lastName: String(row[2] || '').trim(),
    suffix: String(row[5] || '').trim(),
    type: type,
    orgName: homeOrg ? homeOrg.name : '',
    phone: sigDirectoryPhone_(capid, type),
    dutyPositions: sigDutyPositions_(capid, orgs)
  };
}

/**
 * The caller's record, cached for SIG_CACHE_TTL_SECONDS.
 *
 * ⚠️ THE KEY IS THE ONLY THING SEPARATING ONE MEMBER FROM ANOTHER HERE. Do not
 * simplify it away.
 *
 * getUserCache() sounds per-visitor and is not: this app runs `executeAs:
 * USER_DEPLOYING`, so the "user" Apps Script scopes the cache (and User
 * Properties) to is the DEPLOYER — one shared cache behind every member who opens
 * the page. Keying each entry by the AUTHENTICATED address is what keeps one
 * member's grade and duties out of another's signature. A key of, say, 'member'
 * would serve whoever loaded the page first to everyone else for ten minutes.
 *
 * What it holds is the resolved CAPWATCH record only — never rendered HTML, never
 * a client choice — so a cache hit and a miss are indistinguishable in the output.
 *
 * @param {string} actorEmail - from requireMember_()
 * @returns {{capid: string, member: Object}}
 * @throws {Error} with a sentence written for the member when we cannot proceed
 */
function sigResolveMemberForActor_(actorEmail) {
  const cache = CacheService.getUserCache();
  const key = 'sigmember:' + actorEmail;

  const hit = cache.get(key);
  if (hit) {
    try {
      return JSON.parse(hit);
    } catch (err) {
      // A corrupt entry must never be the reason a member cannot fix their
      // signature; fall through and rebuild.
      Logger.warn('Discarding an unreadable cache entry', { user: actorEmail });
    }
  }

  const capid = sigCapidForAccount_(actorEmail);
  if (!capid) {
    throw new Error('Your Workspace account is not linked to a CAPID, so we cannot ' +
      'look up your membership record. ' + sigSupportSentence_());
  }

  const member = sigBuildMemberRecord_(capid);
  if (!member) {
    throw new Error('CAPID ' + capid + ' is not an active member in today\'s CAPWATCH ' +
      'extract, so there is nothing to build a signature from. ' + sigSupportSentence_());
  }

  const resolved = { capid: capid, member: member };
  cache.put(key, JSON.stringify(resolved), SIG_CACHE_TTL_SECONDS);
  return resolved;
}

/** "Email it@example.org if you think this is wrong." — or nothing, if unset. */
function sigSupportSentence_() {
  return SIG_CONFIG.SUPPORT_EMAIL
    ? 'Contact ' + SIG_CONFIG.SUPPORT_EMAIL + ' if you think this is wrong.'
    : 'Contact your wing IT director if you think this is wrong.';
}
