/**
 * ONE-TIME Hawaii Wing tenant setup — paste this file into the target Apps Script
 * editor, fill in the FILL_IN values, select the matching function, and Run once.
 *
 * There are two Hawaii Workspace tenants, mirroring the California setup:
 *   - SENIORS  -> setupHiwgSeniorsScriptProperties()
 *   - CADETS   -> setupHiwgCadetsScriptProperties()
 *
 * Each writes the NON-SECRET tenant identity (canonical copies in
 * config-tenants/hiwg-seniors.json and hiwg-cadets.json) to Script Properties,
 * which the shared config.gs reads via getTenantConfig_(). Safe to re-run:
 * existing values are overwritten with the same data; blanks are skipped.
 *
 * WHY SCRIPT PROPERTIES (not config.gs): all tenants share one config.gs and a
 * `clasp push` overwrites it, so per-tenant values must live in Script Properties
 * (which a push never touches). See src/config.gs getTenantConfig_().
 *
 * WING LABELS ARE AUTOMATIC: with TENANT_WING='HI', config.gs derives 'HIWG'
 * (WING_ABBREVIATION_) and 'Hawaii Wing' (WING_NAMES_) on its own. The
 * TENANT_WING_ABBREVIATION / TENANT_WING_NAME properties below are only needed to
 * override those defaults, so they are left blank.
 *
 * NOT set here (configure separately / by hand):
 *   - Service account: SA_IMPERSONATION_EMAIL, SA_PRIVATE_KEY, SA_PRIVATE_KEY_ID
 *   - CAPWATCH_AUTHORIZATION
 *   - Cross-tenant peer SA (if used): XT_PEER_SA_EMAIL / XT_PEER_SA_SUBJECT via
 *     setupCrossTenantConfig(), XT_PEER_SA_KEY by hand.
 * See docs/ADMIN_GUIDE.md.
 *
 * After running, run validateTenantConfig() (present once the shared src/ is
 * pushed) to confirm the required properties are populated.
 */

/** Writes Hawaii Wing SENIORS tenant identity. Run once in the seniors project. */
function setupHiwgSeniorsScriptProperties() {
  const values = {
    TENANT_PROFILE: 'seniors',
    TENANT_WING: 'HI',
    TENANT_WING_ABBREVIATION: '',           // '' derives HIWG
    TENANT_WING_NAME: '',                   // '' derives 'Hawaii Wing'
    TENANT_REGION: '',
    TENANT_DOMAIN: '',                      // FILL_IN senior Workspace domain
    TENANT_EMAIL_DOMAIN: '',                // FILL_IN '@' + senior domain
    TENANT_SECONDARY_EMAIL_DOMAIN: '',      // optional verified secondary domain
    TENANT_CADETS_TENANT_DOMAIN: '',        // '' derives '<wing>wgcadets.org'; set if different
    TENANT_CAPWATCH_ORGID: '',              // FILL_IN Hawaii Wing CAPWATCH ORGID
    TENANT_CAPWATCH_DATA_FOLDER_ID: '',     // FILL_IN
    TENANT_AUTOMATION_FOLDER_ID: '',        // FILL_IN
    TENANT_AUTOMATION_SPREADSHEET_ID: '',   // FILL_IN
    TENANT_RETENTION_LOG_SPREADSHEET_ID: '',// FILL_IN or leave blank
    TENANT_RETENTION_EMAIL: '',             // FILL_IN
    TENANT_DIRECTOR_RECRUITING_EMAIL: '',   // FILL_IN
    TENANT_AUTOMATION_SENDER_EMAIL: '',     // FILL_IN
    TENANT_SENDER_NAME: 'HIWG Information Technology',
    TENANT_TEST_EMAIL: '',                  // FILL_IN
    TENANT_ITSUPPORT_EMAIL: ''              // FILL_IN
  };
  return applyHiwgProperties_(values);
}

/** Writes Hawaii Wing CADETS tenant identity. Run once in the cadets project. */
function setupHiwgCadetsScriptProperties() {
  const values = {
    TENANT_PROFILE: 'cadets',
    TENANT_WING: 'HI',
    TENANT_WING_ABBREVIATION: '',           // '' derives HIWG
    TENANT_WING_NAME: '',                   // '' derives 'Hawaii Wing' (transition email masthead)
    TENANT_REGION: '',
    TENANT_DOMAIN: '',                      // FILL_IN cadet Workspace domain
    TENANT_EMAIL_DOMAIN: '',                // FILL_IN '@' + cadet domain
    TENANT_CAPWATCH_ORGID: '',              // FILL_IN Hawaii Wing CAPWATCH ORGID
    TENANT_CAPWATCH_DATA_FOLDER_ID: '',     // FILL_IN
    TENANT_AUTOMATION_FOLDER_ID: '',        // FILL_IN
    TENANT_AUTOMATION_SPREADSHEET_ID: '',   // FILL_IN
    TENANT_RETENTION_LOG_SPREADSHEET_ID: '',// FILL_IN or leave blank
    TENANT_RETENTION_EMAIL: '',             // FILL_IN
    TENANT_DIRECTOR_RECRUITING_EMAIL: '',   // FILL_IN
    TENANT_AUTOMATION_SENDER_EMAIL: '',     // FILL_IN
    TENANT_SENDER_NAME: 'HIWG Information Technology',
    TENANT_TEST_EMAIL: '',                  // FILL_IN
    TENANT_ITSUPPORT_EMAIL: ''              // FILL_IN
  };
  return applyHiwgProperties_(values);
}

/** Shared writer: sets non-blank properties, skips blanks, logs, validates. */
function applyHiwgProperties_(values) {
  const props = PropertiesService.getScriptProperties();
  const applied = [];
  Object.keys(values).forEach(function (k) {
    const v = String(values[k] == null ? '' : values[k]).trim();
    if (v !== '') { props.setProperty(k, v); applied.push(k); }
  });

  console.log('Applied ' + applied.length + ' Script Properties: ' + JSON.stringify(applied));
  console.log('Blank fields were skipped (existing values untouched).');
  if (typeof validateTenantConfig === 'function') { validateTenantConfig(); }
  return applied;
}
