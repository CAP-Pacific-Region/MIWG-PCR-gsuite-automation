/**
 * ONE-TIME Pacific Region setup — paste into the "PCR Automation" Apps Script
 * editor, select setupPacificScriptProperties, and Run once.
 *
 * Writes the NON-SECRET tenant identity (parsed from the project's old config.gs,
 * canonical copy in config-tenants/pacific.json) to Script Properties, which the
 * refactored shared config.gs reads via getTenantConfig_(). Safe to re-run:
 * existing values are overwritten with the same data; blanks are skipped.
 *
 * NOT set here (leave your existing values / set by hand):
 *   SA_IMPERSONATION_EMAIL, SA_PRIVATE_KEY, SA_PRIVATE_KEY_ID  (already present —
 *   the live SA delegation works), and CAPWATCH_AUTHORIZATION.
 */
function setupPacificScriptProperties() {
  const values = {
    TENANT_PROFILE: "pacific",
    TENANT_DOMAIN: "pcr.cap.gov",
    TENANT_EMAIL_DOMAIN: "@pcr.cap.gov",
    TENANT_CAPWATCH_ORGID: "434",
    TENANT_WING: "PCR",
    TENANT_REGION: "PCR",
    TENANT_CAPWATCH_DATA_FOLDER_ID: "1F_M0gdbw0_fmsJiAEfUZYm_ekLEXN3ii",
    TENANT_REGION_CAPWATCH_DATA_FOLDER_ID: "1lU9yWHPf1Eij3AEQPmMR8ki7EpgslV9z",
    TENANT_UNIT_VISIT_SPREADSHEET_ID: "1kI6KGn6_bKJpL_aEQHnm6tJ98SeF2cIQq4FdvtiK838",
    TENANT_UNIT_VISIT_CALENDAR_ID: "c_3078a71e04860c7f7bbb12860b9733cb7d6e16bb724dec40836770dbd57cb2a2@group.calendar.google.com",
    TENANT_AUTOMATION_FOLDER_ID: "1JRIPKScCjnKBy7GKWLwj2l23ux0_opGb",
    TENANT_AUTOMATION_SPREADSHEET_ID: "1cSCw1aOwnT8j5-A7giKXFUQwAosMlnnuJLn2_Mbj-Ew",
    TENANT_RETENTION_EMAIL: "noel.luneau@pcr.cap.gov",
    TENANT_DIRECTOR_RECRUITING_EMAIL: "noel.luneau@pcr.cap.gov",
    TENANT_AUTOMATION_SENDER_EMAIL: "automation@pcr.cap.gov",
    TENANT_SENDER_NAME: "Noel Luneau, Lt Col, Director of Recruiting & Retention",
    TENANT_TEST_EMAIL: "noel.luneau@pcr.cap.gov",
    TENANT_ITSUPPORT_EMAIL: "it@pcr.cap.gov",
  };

  const props = PropertiesService.getScriptProperties();
  const applied = [];
  Object.keys(values).forEach(function (k) {
    const v = String(values[k] == null ? '' : values[k]).trim();
    if (v !== '') { props.setProperty(k, v); applied.push(k); }
  });

  console.log('Applied ' + applied.length + ' Script Properties: ' + JSON.stringify(applied));
  // If validateTenantConfig() is present (after the shared src/ is pushed), run it to confirm.
  if (typeof validateTenantConfig === 'function') { validateTenantConfig(); }
  return applied;
}
