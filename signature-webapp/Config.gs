/***********************************************
 * File: Config.gs
 * Description: Per-tenant configuration for the member self-service email
 * signature web app.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-07-28
 * Changes: 1.0.0 — initial version.
 ***********************************************/

/**
 * WHY THIS PROJECT IS SEPARATE FROM src/ AND FROM webapp/
 *
 * An Apps Script project has exactly one doGet. src/ already spends its on the
 * FileMaker mission-provisioning webhook (ANYONE_ANONYMOUS + USER_DEPLOYING), and
 * webapp/ spends its on the Secondary Alias ADMIN interface. This app is the
 * opposite of that one in who may use it — every member of the domain, not a
 * handful of IT staff — so it gets its own project, its own manifest and its own
 * scope list rather than sharing a deployment with alias-mutation code.
 *
 * Consequence: the signature template itself is duplicated from
 * src/accounts-and-groups/UpdateMembers.gs (see SignatureTemplate.gs). That copy
 * is not free — it is pinned by test/SignatureWebApp.test.js, which runs BOTH
 * copies over the same members and fails on the first byte of divergence. Without
 * that test a drift would show up as members whose self-service signature differs
 * from the one the main project pushes, flip-flopping between the two.
 */

/**
 * Like src/config.gs, every tenant-specific value is a Script Property, never a
 * literal — `clasp push` overwrites source files but never touches properties.
 * The TENANT_* names are deliberately identical to the main project's so the
 * canonical values in config-tenants/<tenant>.json can be copied straight across.
 */
function getSignatureWebAppConfig_() {
  const p = PropertiesService.getScriptProperties();
  const get = function (key, fallback) {
    const v = p.getProperty(key);
    return (v === null || String(v).trim() === '') ? (fallback || '') : String(v).trim();
  };

  const wing = get('TENANT_WING').toUpperCase();

  return {
    /** Primary mail domain, '@cawgcap.org'. Signatures are only ever written here… */
    EMAIL_DOMAIN: get('TENANT_EMAIL_DOMAIN'),
    /** …and on the secondary domain, when the tenant has one. */
    SECONDARY_EMAIL_DOMAIN: get('TENANT_SECONDARY_EMAIL_DOMAIN'),
    /** Drive folder holding today's CAPWATCH extract. Read-only here. */
    CAPWATCH_DATA_FOLDER_ID: get('TENANT_CAPWATCH_DATA_FOLDER_ID'),
    /** Two-letter wing code. Feeds the wing website line — see sigWingCode_(). */
    WING: wing,

    /**
     * OPTIONAL Workspace group that restricts who may use the app.
     *
     * BLANK MEANS EVERY AUTHENTICATED DOMAIN USER, which is the intended steady
     * state: this is self-service, and the only thing a member can do with it is
     * set their OWN signature to the one src/ would give them anyway.
     * Set it to a group address to hold the app to a pilot cohort first; unset it
     * to open the doors. Either way the value is a Script Property, so widening or
     * narrowing access is an Admin console edit, not a redeploy.
     *
     * NOTE the reversed default versus webapp/'s WEBAPP_ALIAS_ADMIN_GROUP, which
     * fails closed when blank. That app hands out addresses; this one lets people
     * format their own name.
     */
    ALLOWED_GROUP: get('SIGNATURE_WEBAPP_ALLOWED_GROUP'),

    /** Shown to a member the app cannot help, so they know whom to ask. */
    SUPPORT_EMAIL: get('TENANT_ITSUPPORT_EMAIL'),

    ORG_LABEL: get('TENANT_WING_ABBREVIATION',
      (wing && wing.length === 2) ? wing + 'WG' : (wing || 'CAP')),

    /**
     * Presence only — the key itself is never carried in this object, only read at
     * the point of use in GmailSignature.gs.
     */
    HAS_SA_CREDENTIALS: !!(p.getProperty('SA_IMPERSONATION_EMAIL') && p.getProperty('SA_PRIVATE_KEY'))
  };
}

const SIG_CONFIG = getSignatureWebAppConfig_();

/**
 * CAP's own signature generator — the authority for this format, and the fallback
 * for anyone this page cannot help.
 *
 * Named on the page for both reasons. It is where the template came from, so a
 * member who wants to know why their signature looks the way it does has a source
 * that is not "because the wing's script says so"; and it is a way to get a
 * correct signature by hand when this app refuses — no CAPID on the account, no
 * active CAPWATCH record, the tenant not configured yet. Those refusals otherwise
 * end with nowhere to go.
 *
 * Same site as CAP_SIGNATURE_LOGO_URL in SignatureTemplate.gs; if one moves, check
 * the other.
 */
const SIG_GENERATOR_URL = 'https://cap-brand-tools.netlify.app/signature-generator/index.html';

/**
 * Script Properties without which this app cannot do its job, and the reason the
 * check exists at all.
 *
 * The Apps Script UI will not store a blank value, so an unset property is simply
 * ABSENT — `getProperty()` returns null, which the reader above turns into ''.
 * That is deliberate and correct for the optional ones (no allowed-group means
 * everyone; no support address means a generic "ask your wing IT director").
 *
 * It is NOT harmless for the required ones. With TENANT_EMAIL_DOMAIN unset,
 * isOnATenantDomain_() matches nothing, so every single member is turned away by
 * requireMember_() with a message about THEIR account — a configuration gap
 * wearing the costume of an authorization failure, which is the kind of thing
 * that gets debugged for an afternoon. So both entry points ask this first and
 * say what is actually wrong.
 *
 * @returns {Array<string>} unset property names; [] when the app is usable
 */
function sigMissingConfig_() {
  const missing = [];
  if (!SIG_CONFIG.EMAIL_DOMAIN) missing.push('TENANT_EMAIL_DOMAIN');
  if (!SIG_CONFIG.WING) missing.push('TENANT_WING');
  if (!SIG_CONFIG.CAPWATCH_DATA_FOLDER_ID) missing.push('TENANT_CAPWATCH_DATA_FOLDER_ID');
  // Without these the page can show a member their signature and then fail on the
  // one button it offers. Better to say so before they click it.
  if (!SIG_CONFIG.HAS_SA_CREDENTIALS) missing.push('SA_IMPERSONATION_EMAIL / SA_PRIVATE_KEY');
  return missing;
}

/** A CAPID on these tenants is a 5-7 digit number. Mirrors DUP_GUARD_CAPID_RE. */
const SIG_CAPID_RE = /^\d{5,7}$/;

/**
 * How long a resolved CAPWATCH record stays in the per-user cache.
 *
 * Preview and apply are two round trips over the same member, and each rebuild
 * parses three CAPWATCH files out of Drive (several seconds on a wing-sized
 * extract). Caching the RESOLVED RECORD — never the rendered HTML, never anything
 * the browser sent — keeps apply fast without letting the client influence what
 * gets written: a cache hit and a cache miss produce the same record from the
 * same source. Ten minutes is long enough for someone to read a preview and short
 * enough that a CAPWATCH correction is picked up the same session.
 */
const SIG_CACHE_TTL_SECONDS = 600;

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
