/***********************************************
 * File: MemberRecord.gs
 * Description: Resolves the signed-in caller to their CAPWATCH record — the only
 * source of anything that appears in a signature. Read-only throughout.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.1.0
 * Date: 2026-08-03
 * Changes: 1.1.0 — reads an OPTIONAL region-wide CAPWATCH extract for duties this
 *            tenant's own pull cannot see (region and national billets), for
 *            CAPIDs already on this tenant's roster and orgs outside this wing
 *            only. Adds previewOutOfWingDuties() to diagnose it, and logs whether
 *            the supplement was consulted. Self-disabling when unset.
 *          1.0.0 — initial version.
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
  const folderId = SIG_CONFIG.CAPWATCH_DATA_FOLDER_ID;
  if (!folderId) {
    throw new Error('This tenant is not configured for signatures yet ' +
      '(TENANT_CAPWATCH_DATA_FOLDER_ID is unset).');
  }
  return sigParseCapwatchFromFolder_(folderId, baseName);
}

/**
 * The same, from any folder. Missing folder, missing file and empty file all
 * yield [] rather than throwing: the supplementary region extract is optional,
 * and a member must still get their signature when it is absent.
 *
 * @param {string} folderId
 * @param {string} baseName
 * @returns {Array<Array<string>>} rows without the header, or []
 */
function sigParseCapwatchFromFolder_(folderId, baseName) {
  const key = folderId + '/' + baseName;
  if (SIG_FILE_MEMO[key]) return SIG_FILE_MEMO[key];

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
    Logger.warn('CAPWATCH folder could not be read; continuing without it', {
      folderId: folderId, fileName: baseName + '.txt', errorMessage: err.message
    });
  }

  SIG_FILE_MEMO[key] = rows;
  return rows;
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

  return out.concat(sigOutOfWingDutyPositions_(capid, orgs));
}

/**
 * Duties this tenant's own CAPWATCH extract CANNOT SEE.
 *
 * CAPWATCH scopes an extract to the echelon it was downloaded as: a wing pull
 * carries the duties its members hold at wing or below, and a member's REGION or
 * NATIONAL assignment is absent from it entirely — not filtered, not empty,
 * absent. So a member holding a region billet saw a signature that quietly did
 * not mention it, with nothing in any log to explain why.
 *
 * The region tenant already downloads a region-wide extract, so this reads that
 * folder (shared read-only) purely to fill the gap. Two rules, matching
 * addOutOfWingDutyPositions_() in src/ exactly:
 *
 *   - only for the CAPID already resolved from this tenant's own roster, so the
 *     region extract can never introduce a member this tenant does not have;
 *   - only orgs OUTSIDE this wing, because our own pull is authoritative for
 *     ours and the region extract carries those too.
 *
 * Org names come from the supplementary extract's own Organization.txt, which is
 * the only place a region or national org's name exists.
 *
 * @param {string} capid
 * @param {Object} wingOrgs - this wing's org map, to tell "ours" from "not"
 * @returns {Array<Object>}
 */
function sigOutOfWingDutyPositions_(capid, wingOrgs) {
  const folderId = String(SIG_CONFIG.REGION_CAPWATCH_DATA_FOLDER_ID || '').trim();
  if (!folderId) {
    // Said once per request, because "the supplement is switched off" and "the
    // supplement found nothing" look identical from the outside and are not.
    Logger.info('No supplementary CAPWATCH folder configured; region and national ' +
      'duties will not appear', { property: 'TENANT_REGION_CAPWATCH_DATA_FOLDER_ID' });
    return [];
  }

  const foreignOrgs = {};
  sigParseCapwatchFromFolder_(folderId, 'Organization').forEach(function (row) {
    const orgid = String((row || [])[0] || '').trim();
    if (orgid) {
      foreignOrgs[orgid] = {
        name: String(row[5] || '').trim(),
        scope: String(row[9] || '').trim()
      };
    }
  });

  const out = [];
  ['DutyPosition', 'CadetDutyPositions'].forEach(function (fileName) {
    sigParseCapwatchFromFolder_(folderId, fileName).forEach(function (row) {
      if (String((row || [])[0] || '').trim() !== capid) return;

      const orgid = String(row[7] || '').trim();
      if (wingOrgs[orgid]) return;               // ours; our own extract has it

      const dutyId = String(row[1] || '').trim();
      if (!dutyId) return;

      const org = foreignOrgs[orgid];
      out.push({
        id: dutyId,
        level: String(row[3] || '').trim(),
        assistant: String(row[4] || '').trim() === '1',
        orgName: org ? org.name : '',
        orgScope: org ? org.scope : '',
        outOfWing: true
      });
    });
  });

  // Counted, not just returned: a zero here with a healthy org count means the
  // extract was read and simply holds nothing for this member, which is a very
  // different problem from an unreadable folder.
  Logger.info('Supplementary CAPWATCH consulted for out-of-wing duties', {
    folderId: folderId,
    orgsInSupplement: Object.keys(foreignOrgs).length,
    dutiesFound: out.length
  });

  return out;
}

/**
 * Run-input for previewOutOfWingDuties(). Apps Script cannot pass arguments to a
 * function you Run from the editor, so set the CAPID here first — the same
 * pattern as SIGNATURE_PREVIEW_RUN_INPUTS in src/.
 */
const SIGNATURE_DIAGNOSTIC_RUN_INPUTS = {
  CAPID: ''   // e.g. '123456'
};

/**
 * Answers "why is my region duty not showing?" in one run. Writes NOTHING, and
 * reads no Gmail or Directory data at all.
 *
 * Run this from the EDITOR, not the deployed app: the editor runs the project's
 * current code, so it tells you about the code you just pushed even when the live
 * deployment is still an older version. That distinction is usually the answer.
 *
 * It walks the same path sigOutOfWingDutyPositions_() takes and reports where it
 * stops: property unset, folder unreadable, file absent, no rows for the CAPID,
 * or rows found and what they say.
 */
function previewOutOfWingDuties() {
  const capid = String(SIGNATURE_DIAGNOSTIC_RUN_INPUTS.CAPID || '').trim();
  if (!capid) {
    Logger.error('Set SIGNATURE_DIAGNOSTIC_RUN_INPUTS.CAPID in MemberRecord.gs, then Run again.');
    return;
  }

  const folderId = String(SIG_CONFIG.REGION_CAPWATCH_DATA_FOLDER_ID || '').trim();
  Logger.info('1. Script Property', {
    TENANT_REGION_CAPWATCH_DATA_FOLDER_ID: folderId || '(NOT SET — this alone explains it)'
  });
  if (!folderId) return;

  let folderName = '';
  try {
    folderName = DriveApp.getFolderById(folderId).getName();
    Logger.info('2. Folder is readable by this project', { folderId: folderId, name: folderName });
  } catch (err) {
    Logger.error('2. Folder is NOT readable — check the share, and that the id is a FOLDER id', {
      folderId: folderId, errorMessage: err.message
    });
    return;
  }

  ['Organization', 'DutyPosition', 'CadetDutyPositions'].forEach(function (name) {
    const rows = sigParseCapwatchFromFolder_(folderId, name);
    Logger.info('3. ' + name + '.txt', {
      rows: rows.length,
      note: rows.length ? '' : 'absent or empty — is this the folder the extract actually lands in?'
    });
  });

  const wingOrgs = sigOrgMap_();
  const foreign = {};
  sigParseCapwatchFromFolder_(folderId, 'Organization').forEach(function (row) {
    const orgid = String((row || [])[0] || '').trim();
    if (orgid) foreign[orgid] = { name: String(row[5] || '').trim(), scope: String(row[9] || '').trim() };
  });

  const seen = [];
  ['DutyPosition', 'CadetDutyPositions'].forEach(function (name) {
    sigParseCapwatchFromFolder_(folderId, name).forEach(function (row) {
      if (String((row || [])[0] || '').trim() !== capid) return;
      const orgid = String(row[7] || '').trim();
      const org = foreign[orgid];
      seen.push({
        duty: String(row[1] || '').trim(),
        level: String(row[3] || '').trim(),
        assistant: String(row[4] || '').trim() === '1',
        orgid: orgid,
        org: org ? org.scope + ' ' + org.name : '(no org row in this extract)',
        verdict: wingOrgs[orgid]
          ? 'SKIPPED — this org is in our own wing extract, which is authoritative'
          : 'TAKEN — this is a duty our own pull cannot see'
      });
    });
  });

  Logger.info('4. Rows for this CAPID in the supplementary extract', {
    capid: capid,
    found: seen.length,
    rows: seen,
    note: seen.length ? '' :
      'The extract was read but holds no duty rows for this CAPID. Either it is not ' +
      'the region-wide pull, or it is stale.'
  });
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
