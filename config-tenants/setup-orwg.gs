/**
 * ONE-TIME Oregon Wing tenant setup — paste this file into the Oregon Apps Script
 * editor, select setupOrwgScriptProperties, and Run once. Do this BEFORE the first
 * clasp push of the shared src/ (a push overwrites config.gs).
 *
 * Oregon is ONE composite tenant (orwgcap.org) holding cadets and seniors, so it uses
 * TENANT_PROFILE=composite — not 'seniors', under which cadets would never be
 * provisioned. Canonical copy of these values: config-tenants/orwg.json.
 *
 * Existing property values are OVERWRITTEN with the values below; blanks are
 * skipped. Secrets (SA_*), CAPWATCH_AUTHORIZATION are NOT set here. See
 * docs/ADMIN_GUIDE.md and docs/NEW_TENANT_SETUP.md.
 */
function setupOrwgScriptProperties() {
  const values = {
    TENANT_PROFILE: "composite",
    TENANT_WING: "OR",
    TENANT_WING_ABBREVIATION: "",
    TENANT_WING_NAME: "",
    TENANT_REGION: "",
    TENANT_DOMAIN: "orwgcap.org",
    TENANT_EMAIL_DOMAIN: "@orwgcap.org",
    TENANT_SECONDARY_EMAIL_DOMAIN: "",
    TENANT_CAPWATCH_ORGID: "930",
    TENANT_CAPWATCH_DATA_FOLDER_ID: "1etHX-2gacZfdWeHIawUuJ19TTx04pu9b",
    TENANT_AUTOMATION_FOLDER_ID: "1lX4y8N91J1Ga_7yJUJpaQJ3jRCSbWh6e",
    TENANT_AUTOMATION_SPREADSHEET_ID: "1ux3n7hbwXCmfc1nG5r90AkN5Vc9l5aXCbB4Zs8pgOwc",
    TENANT_RETENTION_LOG_SPREADSHEET_ID: "17bleHCDV9uoibomyTQw5pSxEKM-w8Xacdcd2guV5TPE",
    TENANT_RETENTION_EMAIL: "retention@orwgcap.org",
    TENANT_RECRUITING_MAILBOX: "",
    TENANT_2SV_SETUP_GROUP: "",
    TENANT_DIRECTOR_RECRUITING_EMAIL: "it@orwgcap.org",
    TENANT_DIRECTOR_RECRUITING_NAME: "",
    TENANT_AUTOMATION_SENDER_EMAIL: "automation@orwgcap.org",
    TENANT_SENDER_NAME: "ORWG Information Technology",
    TENANT_TEST_EMAIL: "it@orwgcap.org",
    TENANT_ITSUPPORT_EMAIL: "it@orwgcap.org",
    TENANT_HELP_GUIDE_URL: "",  // optional: welcome-email quick-reference page; blank omits it
    TENANT_SUPPORT_URL: ""      // optional: support-ticket portal; blank mails TENANT_ITSUPPORT_EMAIL
  };

  const props = PropertiesService.getScriptProperties();
  const applied = [];
  Object.keys(values).forEach(function (k) {
    const v = String(values[k] == null ? '' : values[k]).trim();
    if (v !== '') { props.setProperty(k, v); applied.push(k); }
  });

  console.log('Applied ' + applied.length + ' Script Properties: ' + JSON.stringify(applied));
  if (typeof validateTenantConfig === 'function') { validateTenantConfig(); }
  return applied;
}
