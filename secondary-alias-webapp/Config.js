/***********************************************
 * File: Config.gs
 * Description: Per-tenant configuration for the Secondary Alias admin web app.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-07-20
 * Changes: 1.0.0 — initial version.
 ***********************************************/

/**
 * WHY THIS PROJECT IS SEPARATE FROM src/
 *
 * The main automation project already owns doGet()/doPost() for the FileMaker
 * mission-provisioning webhook, deployed ANYONE_ANONYMOUS + USER_DEPLOYING. An
 * Apps Script project has exactly one doGet, and an admin UI behind an anonymous,
 * runs-as-the-deployer endpoint would hand alias-management powers to anyone who
 * learned the URL. So this is its own script project with its own manifest:
 * access DOMAIN (so there is a real authenticated identity to check) and a much
 * smaller scope list than src/.
 *
 * Consequence: a handful of helpers are duplicated from src/ rather than shared
 * (Apps Script libraries add a deploy-version step for two dozen lines). Each
 * duplicate names its source file so the two can be kept in step.
 */

/**
 * Like src/config.gs, every tenant-specific value is a Script Property, never a
 * literal — `clasp push` overwrites source files but never touches properties.
 * The TENANT_* names are deliberately identical to the main project's so the
 * canonical values in config-tenants/<tenant>.json can be copied straight across.
 */
function getWebAppConfig_() {
  const p = PropertiesService.getScriptProperties();
  const get = function (key, fallback) {
    const v = p.getProperty(key);
    return (v === null || String(v).trim() === '') ? (fallback || '') : String(v).trim();
  };
  return {
    EMAIL_DOMAIN: get('TENANT_EMAIL_DOMAIN'),
    SECONDARY_EMAIL_DOMAIN: get('TENANT_SECONDARY_EMAIL_DOMAIN'),
    AUTOMATION_SPREADSHEET_ID: get('TENANT_AUTOMATION_SPREADSHEET_ID'),
    /**
     * Workspace group whose members may use this app. Membership is the ONLY
     * grant — there is no hard-coded owner, so losing access means losing it for
     * everyone and the fix is a group edit in the Admin console, not a redeploy.
     * Blank locks the app to nobody, which is the correct fail-closed default for
     * an unconfigured tenant.
     */
    ADMIN_GROUP: get('WEBAPP_ALIAS_ADMIN_GROUP'),

    // ORG_LABEL ('CAWG') and WING_NAME ('California Wing') are DERIVED from
    // TENANT_WING ('CA'), mirroring src/config.gs, so a deployer only sets the one
    // canonical property. TENANT_WING_ABBREVIATION / TENANT_WING_NAME remain
    // optional overrides for a wing whose derivation is wrong. See webappWing*_.
    ORG_LABEL: webappWingAbbreviation_(get('TENANT_WING'), get('TENANT_WING_ABBREVIATION')),
    WING_NAME: webappWingName_(get('TENANT_WING'), get('TENANT_WING_NAME')),

    // Consumed by Notify.gs when it configures Send-As and emails the member on a
    // new alias. Same TENANT_* names as src/, so copy the values across from
    // config-tenants/<tenant>.json (SA_IMPERSONATION_EMAIL / SA_PRIVATE_KEY are
    // read directly from Script Properties by webappGetImpersonatedToken_).
    AUTOMATION_SENDER_EMAIL: get('TENANT_AUTOMATION_SENDER_EMAIL'),
    ITSUPPORT_EMAIL: get('TENANT_ITSUPPORT_EMAIL'),
    SENDER_NAME: get('TENANT_SENDER_NAME', 'CAP Information Technology')
  };
}

/**
 * Wing/region code -> proper name, mirrored from WING_NAMES_ in src/config.gs.
 * Any code not listed falls back to the abbreviation.
 */
const WEBAPP_WING_NAMES_ = {
  CA: 'California Wing',
  HI: 'Hawaii Wing',
  NV: 'Nevada Wing',
  OR: 'Oregon Wing',
  WA: 'Washington Wing',
  AK: 'Alaska Wing',
  PCR: 'Pacific Region'
};

/** 'CA' -> 'CAWG'; an explicit override wins. Mirrors WING_ABBREVIATION_ in src. */
function webappWingAbbreviation_(wing, override) {
  const explicit = String(override || '').trim();
  if (explicit) return explicit.toUpperCase();
  const w = String(wing || '').trim().toUpperCase();
  if (!w) return 'CAP';
  return (/WG$/.test(w) || w.length !== 2) ? w : w + 'WG';
}

/** 'CA' -> 'California Wing'; an explicit override wins. Mirrors WING_NAME_ in src. */
function webappWingName_(wing, override) {
  const explicit = String(override || '').trim();
  if (explicit) return explicit;
  const w = String(wing || '').trim().toUpperCase();
  return WEBAPP_WING_NAMES_[w] || webappWingAbbreviation_(wing, '');
}

const WEBAPP_CONFIG = getWebAppConfig_();

/** Tab this app reads and writes. Must match SECONDARY_ALIAS_SHEET in src/. */
const WEBAPP_ALIAS_SHEET = 'Secondary Aliases';

/** Append-only audit tab, created on first write. This project's own. */
const WEBAPP_AUDIT_SHEET = 'Alias Admin Log';

/** Status prefix. Must match SECONDARY_ALIAS_CONFLICT_PREFIX in src/. */
const WEBAPP_CONFLICT_PREFIX = 'CONFLICT';

/** A CAPID on these tenants is a 5-7 digit number. Mirrors DUP_GUARD_CAPID_RE. */
const WEBAPP_CAPID_RE = /^\d{5,7}$/;

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
