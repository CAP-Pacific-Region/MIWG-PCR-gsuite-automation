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
 * ONE FILE, TWO TENANTS — AND WHY THAT NEEDS A GUARD
 *
 * The same source is pushed to the seniors project and the cadets project, so
 * this file necessarily carries both tenants' values. Running the wrong one is
 * therefore possible, and it is the single most damaging mistake available here:
 * pointing the cadet tenant at the SENIORS CAPWATCH folder and the
 * `@cawgcap.org` domain would not throw. It would look like it worked, and then
 * every lookup would answer from the wrong wing's extract while every action ran
 * against addresses on a domain this tenant does not own.
 *
 * So each tenant gets its own named function, and both go through
 * admSetupApply_(), which REFUSES to write a profile different from the one
 * already on the project. The first run is unguarded (nothing to contradict);
 * every run after that is checked.
 *
 * Values come from config-tenants/<tenant>.json — keep the two in step.
 */
const ADMIN_WEBAPP_SETUP_VALUES = {
  seniors: {
    TENANT_PROFILE: 'seniors',
    TENANT_DOMAIN: 'cawgcap.org',
    TENANT_EMAIL_DOMAIN: '@cawgcap.org',
    TENANT_SECONDARY_EMAIL_DOMAIN: '@cawg.cap.gov',
    TENANT_WING: 'CA',
    TENANT_CAPWATCH_DATA_FOLDER_ID: '10T0wBubqzUzHa_7nx__eNfuzhTpFRDs3',
    TENANT_ITSUPPORT_EMAIL: 'it@cawgcap.org',

    /** The OTHER tenant, named when no peer admin URL is set. */
    XT_PEER_DOMAIN: 'cawgcadets.org',

    /**
     * The audit log's home — the tenant's existing automation spreadsheet. The
     * app adds its own tab ('Admin Web App Log') on the first action and
     * touches nothing else in it.
     */
    TENANT_AUTOMATION_SPREADSHEET_ID: '1UqCc6aRMEYw-Y_bTcTDKXuaYLsQ6bQzkdoVG7rRsV9Q',

    /**
     * The 2SV setup group. BLANK HIDES THE GROUP PANEL — the app will not guess
     * at a group address, because adding a member to the wrong group is not a
     * mistake it should be capable of making.
     *
     * Note it carries no `ca.` prefix, unlike most groups on this domain — it is
     * typed here exactly as it exists. src/ prunes this same group nightly
     * (TwoSvSetupGroup.gs) under the name TENANT_2SV_SETUP_GROUP; that project
     * needs its own copy of the value.
     */
    WEBAPP_2SV_SETUP_GROUP: '2sv-setup@cawgcap.org',

    /**
     * Linked when an admin opens a CADET here — the cadet tenant's own copy of
     * this app. Derived from that deployment's id; the two apps point at each
     * other, so a help desk that lands on the wrong one is one click from the
     * right one.
     */
    WEBAPP_PEER_ADMIN_URL: 'https://script.google.com/a/macros/cawgcadets.org/s/' +
      'AKfycbx7yMbuNp2So9PbKMzwrn7OL1X6oSddxaml5re_YkHWF6b-ZQEWGsC8-3P8N6mpXM2AeA/exec',

    /** Extra groups this app may change. A security boundary — see Config.gs. */
    WEBAPP_MANAGED_GROUPS: ''
  },

  cadets: {
    TENANT_PROFILE: 'cadets',
    TENANT_DOMAIN: 'cawgcadets.org',
    TENANT_EMAIL_DOMAIN: '@cawgcadets.org',
    // No secondary domain on this tenant; @cawg.cap.gov is seniors-only.
    TENANT_SECONDARY_EMAIL_DOMAIN: '',
    TENANT_WING: 'CA',
    TENANT_CAPWATCH_DATA_FOLDER_ID: '1Y2MmtJoyk4qCMncGmvoIe-rTotje_1dj',
    TENANT_ITSUPPORT_EMAIL: 'it@cawgcap.org',

    /** The OTHER tenant — the seniors domain, from this side. */
    XT_PEER_DOMAIN: 'cawgcap.org',

    TENANT_AUTOMATION_SPREADSHEET_ID: '1tsVoGIbTztl9esydyiFt5Gc6pxIJRlcGdV7oLhXZbF4',

    /**
     * The cadet tenant's OWN 2SV setup group, confirmed on the tenant and filled
     * in there by hand 2026-08-19. Deliberately NOT 2sv-setup@cawgcap.org: that
     * group lives on the other Workspace and cannot hold cadet accounts. Like
     * the seniors entry it carries no prefix — typed exactly as it exists.
     */
    WEBAPP_2SV_SETUP_GROUP: '2sv-setup@cawgcadets.org',

    /** Linked when an admin opens a SENIOR here: the seniors admin site. */
    WEBAPP_PEER_ADMIN_URL: 'https://script.google.com/a/macros/cawgcap.org/s/AKfycby1lqAHHsPq9hHU4ACfuGrZ1wCDOQYX38mA19s-LOtgirB1fP6bN4ZwbJ20nRaOaAcL/exec',

    WEBAPP_MANAGED_GROUPS: ''
  }
};

/**
 * SENIORS tenant. Run this from the editor of the seniors script project.
 * @returns {void}
 */
function setupSeniorsAdminWebApp() {
  admSetupApply_('seniors');
}

/**
 * CADETS tenant. Run this from the editor of the cadets script project.
 * @returns {void}
 */
function setupCadetsAdminWebApp() {
  admSetupApply_('cadets');
}

/**
 * Writes one tenant's values into Script Properties.
 *
 * Refuses when the project already carries a DIFFERENT TENANT_PROFILE — see this
 * file's header for why that mistake is worth a guard rather than a warning.
 *
 * @param {string} tenant - 'seniors' | 'cadets'
 * @returns {void}
 */
function admSetupApply_(tenant) {
  const props = PropertiesService.getScriptProperties();
  const values = ADMIN_WEBAPP_SETUP_VALUES[tenant];
  if (!values) throw new Error('No values for tenant "' + tenant + '".');

  const existing = String(props.getProperty('TENANT_PROFILE') || '').trim().toLowerCase();
  if (existing && existing !== tenant) {
    // Thrown, not logged: this is someone about to point one tenant's app at
    // another tenant's domain and CAPWATCH extract, which fails silently and
    // answers from the wrong wing.
    throw new Error('This project is already configured as the "' + existing + '" tenant, ' +
      'and you just ran the "' + tenant + '" setup. Refusing — that would point this app at ' +
      'the wrong domain and the wrong CAPWATCH extract. Run setup' +
      existing.charAt(0).toUpperCase() + existing.slice(1) + 'AdminWebApp() instead, or clear ' +
      'TENANT_PROFILE by hand if this project really is changing tenant.');
  }

  const set = [];
  const skipped = [];

  Object.keys(values).forEach(function (key) {
    const value = String(values[key] || '').trim();
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

  console.log('Configured this project as the ' + tenant.toUpperCase() + ' tenant.');
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
  console.log('  Other tenant:    ' + (config.PEER_ADMIN_URL || '(no URL set — will name ' +
    (config.PEER_DOMAIN || 'the other Workspace') + ' instead)'));
  console.log('  May use it:      super admins, and holders of ' +
    config.ALLOWED_ROLES.join(' / ') +
    (config.ADMIN_GROUP ? ', plus members of ' + config.ADMIN_GROUP : ''));

  // Named last because it is the one thing this function cannot check: a
  // property is not a deployment.
  console.log('Now deploy from the editor: Deploy > New deployment > Web app, ' +
    'execute as Me, access Anyone in ' + props.getProperty('TENANT_DOMAIN') + '.');
}
