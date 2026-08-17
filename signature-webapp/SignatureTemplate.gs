/***********************************************
 * File: SignatureTemplate.gs
 * Description: The CAP email signature, rendered from a CAPWATCH member record.
 * A port of the generator in src/accounts-and-groups/UpdateMembers.gs — see the
 * parity note below before changing a single character of it.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.1.1
 * Date: 2026-07-28
 * Changes: 1.1.1 — sigSignatureDutyTitle_(): a chosen assistant duty prints as
 *            "Assistant Supply Officer" rather than claiming the billet. Mirrors
 *            UpdateMembers.gs 1.21.1.
 *          1.1.0 — sigDutyBlock_() honors member.selectedDutyKeys: a member may
 *            choose which of their own duties appear, assistants included, and the
 *            defaults (assistants out, one per echelon) give way to that choice.
 *            Echelon order and the two-line cap never do. Assistants now sort after
 *            principals WITHIN an echelon. Mirrors UpdateMembers.gs 1.21.0 exactly.
 *          1.0.1 — dropped the leading <br /> and cut the logo's bottom margin
 *            from 20px to 5px. Mirrors UpdateMembers.gs 1.20.1 exactly.
 *          1.0.0 — initial version, ported from UpdateMembers.gs 1.8.x.
 ***********************************************/

/**
 * DUPLICATED LOGIC — THE OUTPUT MUST MATCH src/ BYTE FOR BYTE.
 *
 * Ported from src/accounts-and-groups/UpdateMembers.gs:
 *   CAP_SIGNATURE_LOGO_URL, DUTY_LEVEL_ORDER, SIGNATURE_MAX_DUTIES,
 *   DUTY_TITLE_OVERRIDES, ORG_NAME_EXPANSIONS, dutyLevelRank_, dutyTitleRank_,
 *   getDutyBlock, formatDutyTitle_, formatOrgName_, getSignatureName,
 *   isUngradedRank_, getPublicRank, formatPhone, getWingCode, and
 *   generateEmailSignature itself; plus toTitleCase from src/utils.gs.
 *
 * Copied rather than shared because this is a separate script project (see
 * Config.gs). src/ writes the same members' signatures — at provisioning, and
 * whenever pushAllSignatures() is run by hand — so any divergence between the two
 * generators means a member's signature changes depending on which one touched it
 * last, and the next bulk push would quietly undo what the member approved here.
 *
 * test/SignatureWebApp.test.js loads BOTH files and asserts identical output over
 * a spread of synthetic members. Keeping that test green is the whole reason the
 * duplication is tolerable. If you change the template, change it in src/ first.
 *
 * Every function here is private (trailing underscore) so that google.script.run,
 * which dispatches by name to any global, cannot reach the renderer directly.
 */

/**
 * Logo used in the email signature, hot-linked from every member's mail client.
 * The 2000x415 master rendered down by the <img> attributes, so it stays sharp on
 * high-DPI displays. Verify with a HEAD request before changing it: a dead URL is
 * invisible in the logs and shows up only as a broken image in already-sent mail.
 */
const SIG_LOGO_URL = 'https://cap-brand-tools.netlify.app/signature-generator/LogoNoAux.png';

/** Organizational level, highest first — the CAP style guide's required order. */
const SIG_DUTY_LEVEL_ORDER = { NAT: 0, NHQ: 0, REGION: 1, WING: 2, GROUP: 3, UNIT: 4 };

/** The CAP style guide caps the signature's duty block at two assignments. */
const SIG_MAX_DUTIES = 2;

/** Retired titles still present in the CAPWATCH feed, shown as the ICL renamed them. */
const SIG_DUTY_TITLE_OVERRIDES = {
  'RECRUITING & RETENTION OFFICER': 'Recruiting Officer',
  'RECRUITING AND RETENTION OFFICER': 'Recruiting Officer',
  'DIRECTOR OF RECRUITING & RETENTION': 'Director of Recruiting',
  'DIRECTOR OF RECRUITING AND RETENTION': 'Director of Recruiting'
};

/** CAPWATCH unit-name abbreviations, expanded for display. ORG NAMES ONLY. */
const SIG_ORG_NAME_EXPANSIONS = {
  SQ: 'Squadron',
  SQD: 'Squadron',
  SQDN: 'Squadron',
  GP: 'Group',
  GRP: 'Group',
  CDT: 'Cadet',
  COMP: 'Composite',
  SR: 'Senior',
  CALIF: 'California',
  HQ: 'HQ'
};

function sigDutyLevelRank_(level) {
  const rank = SIG_DUTY_LEVEL_ORDER[String(level || '').trim().toUpperCase()];
  return rank === undefined ? 99 : rank;
}

/** Seniority of a duty title WITHIN one echelon; lower sorts first. */
function sigDutyTitleRank_(title) {
  const t = String(title || '').trim().toUpperCase();
  if (t === 'COMMANDER') return 0;
  if (t.indexOf('VICE COMMANDER') === 0 || t.indexOf('DEPUTY COMMANDER') === 0) return 1;
  if (t.indexOf('DIRECTOR OF ') === 0) return 2;
  return 3;
}

/** Title Case, breaking on hyphens/apostrophes/periods. Ported from src/utils.gs. */
function sigTitleCase_(str) {
  if (!str || typeof str !== 'string') return '';
  return str.toLowerCase().replace(/\b\w+/g, t => t[0].toUpperCase() + t.substring(1));
}

/**
 * A stable identifier for one duty assignment — how the browser names a duty
 * without being trusted to hand back a duty object. Content-derived, because the
 * record is rebuilt from CAPWATCH between calls and row order is not a promise.
 *
 * @param {Object} dp - one entry of member.dutyPositions
 * @returns {string}
 */
/**
 * Every duty a signature may show: the tenant's own extract, plus the out-of-wing
 * ones (region, national) that a wing CAPWATCH pull cannot see.
 *
 * The only place the two are merged. In src/ that separation is load-bearing —
 * `dutyPositions` there also feeds the Workspace directory title and duty-group
 * matching, and a region billet belongs in neither — so the port keeps the same
 * shape rather than inventing a different one.
 *
 * @param {Object} member
 * @returns {Array<Object>}
 */
function sigAllDuties_(member) {
  return (member.dutyPositions || []).concat(member.outOfWingDutyPositions || []);
}

function sigDutyKey_(dp) {
  return [
    String((dp && dp.id) || '').trim(),
    String((dp && dp.level) || '').trim().toUpperCase(),
    String((dp && dp.orgName) || '').trim().toUpperCase(),
    (dp && dp.assistant) ? 'A' : 'P'
  ].join('|');
}

/**
 * The signature's duty line(s): highest organizational level first, at most two.
 * Returns '' when there is nothing to show, and the template then omits the
 * element entirely rather than emitting an empty heading.
 *
 * By default assistants are excluded and at most one duty per echelon is taken.
 * When `member.selectedDutyKeys` is present — the member said which of their own
 * duties should appear — both of those DEFAULTS give way: they exist to guess well
 * in the member's absence. The echelon ordering and the two-line cap never give
 * way, whoever chose the contents.
 *
 * @param {Object} member - CAPWATCH member object, optionally carrying
 *   selectedDutyKeys (an array of sigDutyKey_() values)
 * @returns {string} HTML, or '' when there is no duty to show
 */
function sigDutyBlock_(member) {
  const all = sigAllDuties_(member);
  const chosen = Array.isArray(member.selectedDutyKeys) ? member.selectedDutyKeys : null;

  const positions = chosen
    ? all.filter(dp => chosen.indexOf(sigDutyKey_(dp)) !== -1)
    : all.filter(dp => !dp.assistant);
  if (positions.length === 0) return '';

  // Echelon is the OUTER key — the style guide asks for highest organizational
  // level first, so a wing assistant still precedes a squadron principal. The
  // assistant term only orders duties within one echelon, and is inert on the
  // default path where assistants are already gone.
  const sorted = positions.slice().sort((a, b) =>
    sigDutyLevelRank_(a.level) - sigDutyLevelRank_(b.level) ||
    (a.assistant ? 1 : 0) - (b.assistant ? 1 : 0) ||
    sigDutyTitleRank_(a.id) - sigDutyTitleRank_(b.id)
  );

  // An explicit selection is already the answer to "which two"; applying the
  // one-per-echelon heuristic to it would drop a duty the member asked for.
  if (chosen) return sigDutyLines_(sorted.slice(0, SIG_MAX_DUTIES), member);

  // At most one duty per echelon first, so two same-level roles cannot crowd out
  // a lower command; then fill any spare slot from the remainder.
  const seenLevel = {};
  const oncePerLevel = [];
  const remainder = [];
  sorted.forEach(dp => {
    const lvl = String(dp.level || '').trim().toUpperCase();
    if (seenLevel[lvl]) {
      remainder.push(dp);
    } else {
      seenLevel[lvl] = true;
      oncePerLevel.push(dp);
    }
  });

  const picked = oncePerLevel.slice(0, SIG_MAX_DUTIES);
  for (let i = 0; picked.length < SIG_MAX_DUTIES && i < remainder.length; i++) {
    picked.push(remainder[i]);
  }

  return sigDutyLines_(picked, member);
}

/** Renders chosen duties as the block's HTML — one writer for both paths. */
function sigDutyLines_(picked, member) {
  return picked
    .map(dp => {
      // dp.level is the same vocabulary as an org's scope and comes straight off
      // the duty row, so it stands in when the org's own scope is missing —
      // otherwise the echelon trim silently does nothing and the raw name ships.
      const org = dp.orgName
        ? sigFormatOrgName_(dp.orgName, dp.orgScope || dp.level)
        : sigFormatOrgName_(member.orgName, member.orgScope);
      return `${org} ${sigSignatureDutyTitle_(dp)}`;
    })
    .join('<br />');
}

/**
 * The duty title as a signature prints it. CAPWATCH keeps the assistant flag in
 * its own column, so the `Duty` value reads "Supply Officer" whether the member
 * holds the billet or assists it — printing it bare claims the billet.
 *
 * Kept out of sigFormatDutyTitle_() for the same reason src/ keeps it out of
 * formatDutyTitle_(): that one normalizes titles for MATCHING elsewhere in the
 * codebase, and display and matching want different strings.
 *
 * @param {Object} dp - one entry of member.dutyPositions
 * @returns {string}
 */
function sigSignatureDutyTitle_(dp) {
  const title = sigFormatDutyTitle_(dp && dp.id);
  if (!title || !(dp && dp.assistant)) return title;
  return /^assistant\b/i.test(title) ? title : 'Assistant ' + title;
}

/** A duty title as displayed: verbatim from CAPWATCH bar the renames above. */
function sigFormatDutyTitle_(dutyId) {
  const title = String(dutyId || '').trim().replace(/\s+/g, ' ');
  return SIG_DUTY_TITLE_OVERRIDES[title.toUpperCase()] || title;
}

/**
 * Title-cases a CAPWATCH unit name and expands its abbreviations:
 * "SAN JOSE SR SQDN 80" -> "San Jose Senior Squadron 80". Wing and region orgs
 * are cut back to the echelon ("CALIFORNIA WING HQ" -> "California Wing").
 *
 * @param {string} orgName - Raw CAPWATCH unit name
 * @param {string} [scope] - CAPWATCH org scope: UNIT / GROUP / WING / REGION
 * @returns {string} Display form
 */
function sigFormatOrgName_(orgName, scope) {
  let name = String(orgName || '').trim();
  const s = String(scope || '').trim().toUpperCase();

  if (s === 'WING' || s === 'REGION') {
    const m = name.toUpperCase().match(/\b(WING|REGION)\b/);
    if (m) name = name.slice(0, m.index + m[1].length);
  }

  // A trailing "CAP" is the organization's own initials ("PACIFIC REGION CAP"),
  // which title-cases into a word — "Pacific Region Cap Director of Safety" — and
  // repeats what the line below it already says in full. After the echelon trim,
  // so it also catches a duty whose org scope did not survive the extract.
  name = name.replace(/\s+CAP$/i, '');

  return sigTitleCase_(name)
    .split(/\s+/)
    .filter(Boolean)
    .map(word => {
      const key = word.replace(/\.$/, '').toUpperCase();
      return SIG_ORG_NAME_EXPANSIONS[key] || word;
    })
    .join(' ');
}

/** True for a member with no grade to display: CAPWATCH 'SM', or a blank Rank. */
function sigIsUngradedRank_(rank) {
  const r = String(rank || '').trim().toUpperCase();
  return r === '' || r === 'SM';
}

/**
 * The signature's name line. An ungraded senior ('SM') is shown by name with a
 * middle initial, because the style guide does not permit 'SM' as a grade.
 *
 * @param {Object} member - CAPWATCH member object
 * @returns {string} e.g. "Maj. Jane Doe", or ungraded, "Jane M. Doe"
 */
function sigName_(member) {
  const first = String(member.firstName || '').trim();
  const last = String(member.lastName || '').trim();
  const suffix = String(member.suffix || '').trim();

  if (sigIsUngradedRank_(member.rank)) {
    const initial = String(member.middleName || '').trim().charAt(0);
    return [first, initial ? initial + '.' : '', last, suffix].filter(Boolean).join(' ');
  }

  return [sigPublicRank_(member.rank), first, last, suffix].filter(Boolean).join(' ');
}

/** A CAPWATCH grade in the form CAP's style guide prints it. */
function sigPublicRank_(rank) {
  const MAP = {
    // Senior members.
    "SSgt": "Staff Sgt.",
    "TSgt": "Tech. Sgt.",
    "MSgt": "Master Sgt.",
    "SMSgt": "Senior Master Sgt.",
    "CMSgt": "Chief Master Sgt.",
    "FO": "Flight Officer",
    "TFO": "Tech. Flight Officer",
    "SFO": "Senior Flight Officer",
    "2d Lt": "2nd Lt.",
    "1st Lt": "1st Lt.",
    "Capt": "Capt.",
    "Maj": "Maj.",
    "Lt Col": "Lt. Col.",
    "Col": "Col.",
    "Brig Gen": "Brig. Gen.",
    "Maj Gen": "Maj. Gen.",

    // Cadets. NB the CAPWATCH spellings are NOT the senior ones with a "C/" glued
    // on — the cadet forms have no internal space ("C/2dLt", not "C/2d Lt").
    // "CADET" is CAPWATCH's representation of C/AB, the entry grade.
    "CADET": "Cadet Airman Basic",
    "C/Amn": "Cadet Airman",
    "C/A1C": "Cadet Airman 1st Class",
    "C/SrA": "Cadet Senior Airman",
    "C/SSgt": "Cadet Staff Sgt.",
    "C/TSgt": "Cadet Tech. Sgt.",
    "C/MSgt": "Cadet Master Sgt.",
    "C/SMSgt": "Cadet Senior Master Sgt.",
    "C/CMSgt": "Cadet Chief Master Sgt.",
    "C/2dLt": "Cadet 2nd Lt.",
    "C/1stLt": "Cadet 1st Lt.",
    "C/Capt": "Cadet Capt.",
    "C/Maj": "Cadet Maj.",
    "C/LtCol": "Cadet Lt. Col.",
    "C/Col": "Cadet Col."
  };
  return MAP[rank] || rank || '';
}

/** '+15551234567' -> '555.123.4567'; '' when there is nothing usable. */
function sigFormatPhone_(phone) {
  if (!phone) return '';
  const digits = phone.replace(/\D/g, '').slice(-10);
  return digits ? digits.replace(/(\d{3})(\d{3})(\d{4})/, '$1.$2.$3') : '';
}

/** The wing website host: 'CA' -> 'cawg' -> cawg.cap.gov. Ported verbatim. */
function sigWingCode_() {
  return String(SIG_CONFIG.WING || '').toLowerCase() + 'wg';
}

/**
 * Generates the HTML email signature for a member using the standard CAP template.
 *
 * @param {Object} member - A resolved CAPWATCH member record (see MemberRecord.gs)
 * @returns {string} HTML signature block
 */
function sigGenerateEmailSignature_(member) {
  const nameLine = sigName_(member);
  const duty = sigDutyBlock_(member);
  const wingCode = sigWingCode_();
  const phoneDigits = member.phone ? member.phone.replace(/\D/g, '').slice(-10) : '';
  const phoneFormatted = sigFormatPhone_(member.phone);

  return `
<!DOCTYPE html>
<html>
<body>
<h1 style="font-size: 12px; line-height: 12px;
           font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
           color: #001871; font-weight: bold; margin: 0 0 5px;">
  ${nameLine}
</h1>

${duty ? `<h2 style="font-size: 12px; line-height: 14px;
           font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
           color: #000000; font-weight: normal; margin: 0 0 5px;">
  ${duty}
</h2>` : ''}

<p style="font-size: 12px; line-height: 12px;
          font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
          color: #000000; font-weight: bold; margin: 0 0 5px;">
  Civil Air Patrol, U.S. Air Force Auxiliary
</p>

${phoneFormatted ? `<p style="font-size: 12px; line-height: 12px;
          font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
          color: #000000; font-weight: normal; margin: 0 0 5px;">
  (M) <a href="tel:+1${phoneDigits}" style="color: #000000; text-decoration: none;">${phoneFormatted}</a>
</p>` : ''}

<p style="font-size: 12px; line-height: 12px;
          font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
          color: #001871; font-weight: normal; margin: 0 0 5px;">
  <a href="https://www.GoCivilAirPatrol.com"
     style="color: #000000; text-decoration: underline;">
     GoCivilAirPatrol.com
  </a>
</p>

<p style="font-size: 12px; line-height: 12px;
        font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
        color: #001871; font-weight: normal; margin: 0 0 5px;">
 <a href="https://${wingCode}.cap.gov"
   style="color: #000000; text-decoration: underline;">
   ${wingCode}.cap.gov
 </a>
</p>

<a href="https://www.GoCivilAirPatrol.com">
  <img
    src="${SIG_LOGO_URL}"
    width="200"
    height="42"
    style="display:block; border:0; outline:none; text-decoration:none;
           width:200px; max-width:200px; height:42px; margin: 15px 0 5px 0;"
    alt="Civil Air Patrol Logo" />
</a>

<p style="font-size: 12px; line-height: 14px;
          font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
          color:#000000; font-weight: normal; font-style: italic;
          white-space: normal; margin: 0 0 5px;">
  Volunteers serving America&apos;s communities, saving lives, and shaping futures.
</p>

</body>
</html>
  `;
}
