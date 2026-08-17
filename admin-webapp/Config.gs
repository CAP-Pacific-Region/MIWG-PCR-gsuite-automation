/***********************************************
 * File: Config.gs
 * Description: Per-tenant configuration for the domain admin help-desk web app.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-08-16
 * Changes: 1.0.0 — initial version.
 ***********************************************/

/**
 * WHAT THIS APP IS, AND WHY IT IS A FOURTH SCRIPT PROJECT
 *
 * Wing IT staff hold Google's Help Desk Administrator role, which lets them reset
 * a member's password from the Admin console — and nothing else this wing's
 * automation does. Resending a welcome email, or parking a member in the 2SV
 * setup group, meant asking someone with access to the Apps Script editor to run
 * a function by hand. That is a bottleneck made of one person, and the editor is
 * not somewhere a help-desk volunteer should be: every function in src/ is one
 * mis-click away from a wing-wide run.
 *
 * So this app exposes the small set of per-member operations a help desk actually
 * needs, each one scoped to a single member, each one logged with the name of the
 * admin who ran it.
 *
 * It is a SEPARATE Apps Script project for the same reason signature-webapp/ is:
 * a project has exactly one doGet, and the three that exist are spoken for —
 * src/ on the FileMaker mission-provisioning webhook (anonymous), and
 * signature-webapp/ on member self-service. Separation also keeps the blast
 * radius honest: this project's OAuth scopes are the ones these four actions
 * need and no more, so a bug here cannot touch calendars, chat spaces, shared
 * contacts, or group settings.
 *
 * Consequence: the pieces it borrows from src/ are COPIES (the welcome email
 * template and its eligibility policy, the CAPID account lookup, the audit
 * ledger's format). Each copy is pinned by test/AdminWebApp.test.js against the
 * original, which is the price of the duplication — see that file's header.
 */

/**
 * Like src/config.gs, every tenant-specific value is a Script Property, never a
 * literal: `clasp push` overwrites source files but never touches properties. The
 * TENANT_* names are deliberately identical to the main project's, so the
 * canonical values in config-tenants/<tenant>.json copy straight across.
 */
function getAdminWebAppConfig_() {
  const p = PropertiesService.getScriptProperties();
  const get = function (key, fallback) {
    const v = p.getProperty(key);
    return (v === null || String(v).trim() === '') ? (fallback || '') : String(v).trim();
  };
  const list = function (key, fallback) {
    return get(key, fallback).split(',')
      .map(function (s) { return s.trim(); })
      .filter(Boolean);
  };

  const wing = get('TENANT_WING').toUpperCase();

  return {
    /** Primary mail domain, '@cawgcap.org'. */
    EMAIL_DOMAIN: get('TENANT_EMAIL_DOMAIN'),
    /** The same domain without the '@', as it appears in the welcome email's sign-in link. */
    DOMAIN: get('TENANT_DOMAIN'),
    /** Secondary mail domain, when the tenant has one. Blank on most. */
    SECONDARY_EMAIL_DOMAIN: get('TENANT_SECONDARY_EMAIL_DOMAIN'),
    /** Drive folder holding today's CAPWATCH extract. */
    CAPWATCH_DATA_FOLDER_ID: get('TENANT_CAPWATCH_DATA_FOLDER_ID'),
    /** Two-letter wing code, e.g. 'CA'. Appears in the welcome email. */
    WING: wing,

    /**
     * WHO MAY USE THIS APP.
     *
     * Directory ROLES, not a group — because the answer to "should this person be
     * able to reset a member's password?" is already recorded in the Admin
     * console, and a second list in Script Properties would be a copy that drifts.
     * A help-desk volunteer who loses the role loses this app the same day,
     * without anyone remembering to edit a group.
     *
     * Super admins are always allowed, and are not listed here (a super admin has
     * every privilege by definition; requiring a role assignment as well would
     * lock the wing's own admins out of a tool they can already do more than).
     *
     * Built-in Google roles carry a fixed roleName: the Help Desk Administrator
     * is `_HELP_DESK_ADMIN_ROLE`, which is the default. Custom roles are matched
     * on the name you gave them, case-insensitively.
     */
    ALLOWED_ROLES: list('WEBAPP_ADMIN_ROLES', '_HELP_DESK_ADMIN_ROLE'),

    /**
     * OPTIONAL extra allow-list, for someone who must use the app without holding
     * an admin role at all. Blank — the intended steady state — means roles are
     * the whole gate. It is additive: it can only widen access, never narrow it.
     */
    ADMIN_GROUP: get('WEBAPP_ADMIN_GROUP'),

    /**
     * The 2SV setup group — the one members are parked in while they enroll, and
     * the only group the app touches unless MANAGED_GROUPS adds more.
     *
     * Blank disables the group actions entirely rather than guessing at an
     * address: adding a member to the wrong group is not a mistake this app
     * should be capable of making.
     */
    TWO_SV_GROUP: get('WEBAPP_2SV_SETUP_GROUP'),

    /**
     * OPTIONAL further groups this app may add to and remove from, comma
     * separated. The 2SV group is always included.
     *
     * THIS LIST IS A SECURITY BOUNDARY, not a convenience. The browser names a
     * group in every membership request, and admGroupFromRequest_() refuses any
     * address that is not on this list — so a help-desk admin cannot use the app
     * to add themselves to, say, a group that grants Drive access. Widening it is
     * an Admin console edit, deliberately not a code change.
     */
    MANAGED_GROUPS: list('WEBAPP_MANAGED_GROUPS'),

    /**
     * Where the action log is written, in addition to Stackdriver. Optional: with
     * no spreadsheet the app still logs every action to the execution log, which
     * is the record of last resort. A sheet is worth setting because the people
     * who need to read this log are not the people with Cloud Logging access.
     */
    AUDIT_SPREADSHEET_ID: get('TENANT_AUTOMATION_SPREADSHEET_ID'),
    AUDIT_SHEET_NAME: get('WEBAPP_AUDIT_SHEET_NAME', 'Admin Web App Log'),

    /**
     * Which population this tenant provisions — 'seniors' | 'cadets' | 'region',
     * the same property and the same default as src/config.gs.
     *
     * It matters here because CAPWATCH is scoped to the WING, not to the tenant.
     * A seniors-tenant extract lists every cadet in the wing, so a name search
     * finds them — and none of them has an account here to act on. See
     * admTenantProvisionsType_().
     */
    PROFILE: get('TENANT_PROFILE', 'seniors').toLowerCase(),

    /**
     * Where to send an admin holding a member this tenant does not provision.
     * Blank falls back to naming the peer domain, which is still better than a
     * dead end.
     */
    CADET_TOOLS_URL: get('WEBAPP_CADET_TOOLS_URL'),
    CADETS_TENANT_DOMAIN: get('TENANT_CADETS_TENANT_DOMAIN'),

    /** Shown to an admin the app cannot help, and CC'd on credentials it mails. */
    SUPPORT_EMAIL: get('TENANT_ITSUPPORT_EMAIL'),

    ORG_LABEL: get('TENANT_WING_ABBREVIATION',
      (wing && wing.length === 2) ? wing + 'WG' : (wing || 'CAP'))
  };
}

const ADMIN_CONFIG = getAdminWebAppConfig_();

/**
 * Script Properties without which this app cannot do its job.
 *
 * Asked BEFORE identity, for the reason signature-webapp learned the hard way: an
 * unset TENANT_EMAIL_DOMAIN turns every caller away with a message about their
 * account, which is a configuration gap wearing the costume of an authorization
 * failure. Say what is actually wrong.
 *
 * The group properties are deliberately NOT here. A tenant with no 2SV group is
 * fully functional minus one panel, and refusing to serve the page over it would
 * take password resets away too.
 *
 * @returns {Array<string>} unset property names; [] when the app is usable
 */
function admMissingConfig_() {
  const missing = [];
  if (!ADMIN_CONFIG.EMAIL_DOMAIN) missing.push('TENANT_EMAIL_DOMAIN');
  if (!ADMIN_CONFIG.DOMAIN) missing.push('TENANT_DOMAIN');
  if (!ADMIN_CONFIG.WING) missing.push('TENANT_WING');
  if (!ADMIN_CONFIG.CAPWATCH_DATA_FOLDER_ID) missing.push('TENANT_CAPWATCH_DATA_FOLDER_ID');
  return missing;
}

/**
 * The CAPWATCH member types each tenant profile provisions accounts for.
 * Mirrors MEMBER_TYPES_ACTIVE in src/config.gs — keep the two in step; the test
 * compares them.
 *
 * Note the region profile provisions CADET as well as the senior types, which is
 * why this is a table rather than "cadets are somebody else's problem".
 */
const ADM_PROFILE_MEMBER_TYPES = {
  seniors: ['SENIOR', 'FIFTY YEAR', 'INDEFINITE', 'CADET SPONSOR'],
  cadets: ['CADET'],
  region: ['SENIOR', 'FIFTY YEAR', 'INDEFINITE', 'CADET']
};

/**
 * True when this tenant is the one that holds accounts for a member of this
 * CAPWATCH type.
 *
 * WHY THE APP NEEDS TO ASK. CAPWATCH is scoped to the WING, so the seniors
 * tenant's extract lists every cadet in the wing alongside every senior. A help
 * desk searching for "Okonkwo" gets both, and the cadets among them have no
 * account here — their accounts are on the cadet Workspace, which this app
 * cannot reach and should not pretend to. Saying so, and saying where to go
 * instead, is the difference between a dead end and an answer.
 *
 * An unknown type (a blank cell, a new CAPWATCH value) reads as NOT ours. That
 * is the safe way round: the worst case is an unnecessary "try the other site"
 * note on a member whose accounts are listed right below it, rather than a help
 * desk quietly failing to act on someone.
 *
 * Pure.
 *
 * @param {string} type - CAPWATCH Member.txt `Type`
 * @returns {boolean}
 */
function admTenantProvisionsType_(type) {
  const wanted = String(type || '').trim().toUpperCase();
  if (!wanted) return false;
  const types = ADM_PROFILE_MEMBER_TYPES[ADMIN_CONFIG.PROFILE] || ADM_PROFILE_MEMBER_TYPES.seniors;
  return types.indexOf(wanted) !== -1;
}

/**
 * Where an admin should go for a member this tenant does not provision, as a
 * sentence. Never a dead end: with no URL configured it names the peer domain,
 * and with neither it still says the account is not here.
 *
 * @param {string} type - the member's CAPWATCH type, for the wording
 * @returns {string}
 */
function admElsewhereSentence_(type) {
  const kind = String(type || '').trim().toUpperCase() === 'CADET' ? 'Cadet' : 'This member\'s';
  const where = ADMIN_CONFIG.CADET_TOOLS_URL
    ? 'Use the cadet tools site instead: ' + ADMIN_CONFIG.CADET_TOOLS_URL
    : (ADMIN_CONFIG.CADETS_TENANT_DOMAIN
      ? 'Their accounts live on ' + ADMIN_CONFIG.CADETS_TENANT_DOMAIN + ', which has its own admin page.'
      : 'Their accounts live on the cadet Workspace, which has its own admin page.');
  return kind + ' accounts are not on this Workspace, so nothing on this page can ' +
    'act on them. ' + where;
}

/** A CAPID on these tenants is a 5-7 digit number. Mirrors DUP_GUARD_CAPID_RE. */
const ADM_CAPID_RE = /^\d{5,7}$/;

/** Most a name/email search will return before it asks for something narrower. */
const ADM_SEARCH_LIMIT = 25;

/**
 * Structured logger. Named `Logger` on purpose: the whole codebase shadows the
 * built-in Apps Script Logger with this shape, so `Logger.info(msg, ctx)` means
 * the same thing here as in src/utils.gs. Never call Logger.log().
 */
const Logger = {
  info: function (message, data) { console.log(JSON.stringify({ level: 'INFO', timestamp: new Date().toISOString(), message: message, data: data || {} })); },
  warn: function (message, data) { console.warn(JSON.stringify({ level: 'WARN', timestamp: new Date().toISOString(), message: message, data: data || {} })); },
  error: function (message, data) { console.error(JSON.stringify({ level: 'ERROR', timestamp: new Date().toISOString(), message: message, data: data || {} })); }
};

/** "Contact it@example.org if …" — or a generic fallback when unset. */
function admSupportSentence_() {
  return ADMIN_CONFIG.SUPPORT_EMAIL
    ? 'Contact ' + ADMIN_CONFIG.SUPPORT_EMAIL + ' if you think this is wrong.'
    : 'Contact your wing IT director if you think this is wrong.';
}

/** HTML-escapes a value for the few places the server writes markup itself. */
function admEscape_(value) {
  return String(value == null ? '' : value).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

/**
 * True if an address sits on a domain this tenant hands out. Compares the full
 * domain, never a suffix: a lookalike @cawgcap.org.example.com must not pass.
 * Ported from isOnATenantDomain_() in signature-webapp/Auth.gs.
 *
 * @param {string} email
 * @returns {boolean}
 */
function admIsTenantAddress_(email) {
  const addr = String(email || '').trim().toLowerCase();
  const at = addr.lastIndexOf('@');
  if (at < 0) return false;
  const domain = addr.slice(at + 1);
  if (!domain) return false;

  return [ADMIN_CONFIG.EMAIL_DOMAIN, ADMIN_CONFIG.SECONDARY_EMAIL_DOMAIN, ADMIN_CONFIG.DOMAIN]
    .map(function (d) { return String(d || '').trim().toLowerCase().replace(/^@+/, ''); })
    .filter(Boolean)
    .indexOf(domain) !== -1;
}
