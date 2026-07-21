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
    ORG_LABEL: get('TENANT_WING_ABBREVIATION', 'CAP')
  };
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
