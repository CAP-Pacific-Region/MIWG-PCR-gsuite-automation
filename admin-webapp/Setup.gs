/***********************************************
 * File: Setup.gs
 * Description: One-time Script Property setup, run BY HAND from the editor.
 * Nothing calls it and no trigger fires it.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-08-16
 * Changes: 1.0.0 — initial version.
 ***********************************************/

/**
 * WHY THIS IS IN THE PROJECT RATHER THAN A LIST OF PROPERTIES TO TYPE
 *
 * Eleven properties typed by hand is eleven chances to paste a folder id one
 * character short, and the failure that produces — "no CAPWATCH member with that
 * CAPID", for every member — looks like a bug in the app rather than a typo.
 * The values below are copied from config-tenants/seniors.json, which is the
 * canonical non-secret record for this tenant.
 *
 * Running it also does something the properties do not: it triggers the OAuth
 * consent prompt. A web app deployed `executeAs: USER_DEPLOYING` serves nobody
 * until the deployer has authorized the project's scopes, and a Run from the
 * editor is the ordinary way to do that.
 *
 * SAFE TO LEAVE IN PLACE. It is not reachable from the web app — the browser can
 * only call the api* functions in AdminApi.gs — and it never runs on its own.
 * It is also safe to re-run: BLANK VALUES ARE SKIPPED, never written, so filling
 * in one of the FILL_IN entries below and running it again cannot clear anything
 * already set. That is the same behavior as setupTenantConfig() in src/.
 */

/**
 * The tenant's own values. Everything above the FILL_IN block comes from
 * config-tenants/seniors.json; keep the two in step.
 */
const ADMIN_WEBAPP_SETUP_VALUES = {
  TENANT_PROFILE: 'seniors',
  TENANT_DOMAIN: 'cawgcap.org',
  TENANT_EMAIL_DOMAIN: '@cawgcap.org',
  TENANT_SECONDARY_EMAIL_DOMAIN: '@cawg.cap.gov',
  TENANT_WING: 'CA',
  TENANT_CAPWATCH_DATA_FOLDER_ID: '10T0wBubqzUzHa_7nx__eNfuzhTpFRDs3',
  TENANT_ITSUPPORT_EMAIL: 'it@cawgcap.org',
  TENANT_CADETS_TENANT_DOMAIN: 'cawgcadets.org',

  /**
   * The audit log's home. This is the tenant's existing automation spreadsheet —
   * the app adds its own tab ('Admin Web App Log') on the first action and
   * touches nothing else in it.
   */
  TENANT_AUTOMATION_SPREADSHEET_ID: '1UqCc6aRMEYw-Y_bTcTDKXuaYLsQ6bQzkdoVG7rRsV9Q',

  /**
   * The 2SV setup group's address. BLANK HIDES THE GROUP PANEL ENTIRELY — the
   * app will not guess at a group address, because adding a member to the wrong
   * group is not a mistake it should be capable of making.
   *
   * Set on the seniors tenant 2026-08-16. Note it carries no `ca.` prefix,
   * unlike most groups on this domain — it is typed here exactly as it exists.
   */
  WEBAPP_2SV_SETUP_GROUP: '2sv-setup@cawgcap.org',

  // ---- FILL IN, then re-run. Blank is valid for both. ----

  /**
   * Linked when an admin opens a cadet, who has no account on this tenant.
   * Blank falls back to naming cawgcadets.org, which is still not a dead end.
   */
  WEBAPP_CADET_TOOLS_URL: '',

  /**
   * Further groups this app may add to and remove from, comma separated.
   * This list is a security boundary — see Config.gs. Blank is the right
   * default: the 2SV group above is already included.
   */
  WEBAPP_MANAGED_GROUPS: ''
};

/**
 * Writes the values above into Script Properties.
 *
 * Run this ONCE from the editor, then check the log: it prints what it set, what
 * it skipped, and whether the app is ready to serve.
 *
 * @returns {void}
 */
function setupAdminWebAppProperties() {
  const props = PropertiesService.getScriptProperties();
  const set = [];
  const skipped = [];

  Object.keys(ADMIN_WEBAPP_SETUP_VALUES).forEach(function (key) {
    const value = String(ADMIN_WEBAPP_SETUP_VALUES[key] || '').trim();
    if (!value) {
      // Skipped, not written as '': the Apps Script UI will not store a blank
      // anyway, and writing one would let a re-run clear a value someone had
      // set by hand in the console.
      skipped.push(key);
      return;
    }
    props.setProperty(key, value);
    set.push(key);
  });

  console.log('Set ' + set.length + ' properties: ' + set.join(', '));
  if (skipped.length) {
    console.log('Left alone (blank in ADMIN_WEBAPP_SETUP_VALUES, existing values untouched): ' +
      skipped.join(', '));
  }

  checkAdminWebAppSetup();
}

/**
 * READ-ONLY. Says whether the app can serve, and what is missing if not.
 *
 * Reads the properties back through the app's OWN config loader rather than the
 * literals above, so it reports what the running app will actually see — which
 * is the point when someone has edited a property in the console.
 *
 * @returns {void}
 */
function checkAdminWebAppSetup() {
  const config = getAdminWebAppConfig_();
  const props = PropertiesService.getScriptProperties();

  const missing = [];
  if (!config.EMAIL_DOMAIN) missing.push('TENANT_EMAIL_DOMAIN');
  if (!config.DOMAIN) missing.push('TENANT_DOMAIN');
  if (!config.WING) missing.push('TENANT_WING');
  if (!config.CAPWATCH_DATA_FOLDER_ID) missing.push('TENANT_CAPWATCH_DATA_FOLDER_ID');

  if (missing.length) {
    console.log('NOT READY — these are still unset: ' + missing.join(', '));
    return;
  }

  // Proving the folder is READABLE, not merely that an id is present: a folder
  // id that is right-looking and wrong presents as "no CAPWATCH member with that
  // CAPID" for every member, which reads as a broken app.
  let capwatch = 'unreadable';
  try {
    const folder = DriveApp.getFolderById(config.CAPWATCH_DATA_FOLDER_ID);
    capwatch = folder.getFilesByName('Member.txt').hasNext()
      ? 'readable, Member.txt present'
      : 'READABLE BUT Member.txt IS NOT IN IT — wrong folder?';
  } catch (err) {
    capwatch = 'NOT READABLE (' + err.message + ')';
  }

  console.log('Ready to serve.');
  console.log('  Tenant:          ' + config.ORG_LABEL + ' (' + config.PROFILE + ' profile)');
  console.log('  CAPWATCH folder: ' + capwatch);
  console.log('  Audit log:       ' + (config.AUDIT_SPREADSHEET_ID
    ? 'sheet "' + config.AUDIT_SHEET_NAME + '" (created on first action)'
    : 'execution log only — TENANT_AUTOMATION_SPREADSHEET_ID is unset'));
  console.log('  Group panel:     ' + (admManagedGroups_().length
    ? admManagedGroups_().join(', ')
    : 'HIDDEN — WEBAPP_2SV_SETUP_GROUP is unset'));
  console.log('  Cadet tools:     ' + (config.CADET_TOOLS_URL || '(not set — will name ' +
    (config.CADETS_TENANT_DOMAIN || 'the cadet Workspace') + ' instead)'));
  console.log('  May use it:      super admins, and holders of ' +
    config.ALLOWED_ROLES.join(' / ') +
    (config.ADMIN_GROUP ? ', plus members of ' + config.ADMIN_GROUP : ''));

  // Named last because it is the one thing this function cannot check: a
  // property is not a deployment.
  console.log('Now deploy from the editor: Deploy > New deployment > Web app, ' +
    'execute as Me, access Anyone in ' + props.getProperty('TENANT_DOMAIN') + '.');
}
