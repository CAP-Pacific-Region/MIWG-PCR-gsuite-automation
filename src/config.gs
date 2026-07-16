/***********************************************
 * File: config.gs
 * Description: Centralized configuration and constants for CAPWATCH automation scripts.
 * Provides organization-specific parameters, email domains, folder IDs, and time zone mapping.
 * Author: Noel Luneau
 * Contributors: Maj Isaac Wilson IV, California Wing (1.4.0, 1.5.0)
 * Version: 1.5.0
 * Date: 2026-07-15
 * Changes: Added LICENSE_CONFIG.MIN_MEMBER_ROWS, a floor under the CAPWATCH
 *   Member.txt row count. deleteIneligibleSuspendedUsers() reads a missing
 *   CAPWATCH record as proof a member lapsed long ago, so a truncated extract
 *   would make current members look deletable; the deletion path now refuses to
 *   run below the floor. Also rewrote DAYS_BEFORE_DELETE_INELIGIBLE's docs: it is
 *   measured from CAPWATCH Expiration, not from lastLoginTime as before.
 *   (1.4.0: added SECONDARY_EMAIL_DOMAIN (Script Property
 *   TENANT_SECONDARY_EMAIL_DOMAIN), consumed by
 *   accounts-and-groups/SecondaryDomainAliases.gs — the verified secondary domain
 *   a tenant hands out parallel addresses on. Blank on cadets/pacific, which
 *   disables that module there.
 *   1.3.0: added a per-profile CROSS_TENANT block (consumed by
 *   cross-tenant-contacts/CrossTenantContacts.gs) selecting cross-tenant shared-
 *   contact behavior — on for seniors/cadets, off for pacific.
 *   1.2.2: profile-driven SQUADRON_DISTRIBUTION_TOGGLES per tenant, so squadron
 *   list creation is tenant-aware; cadets = all-hands + cadets + parents (no
 *   .seniors / command-staff lists), prior toggles were hard-coded.
 *   1.2.1: same, but cadets initially excluded .all; .all kept by request.
 *   1.2.0: per-feature region flags + REGION_CAPWATCH_DATA_FOLDER_ID. 1.1.0:
 *   'pacific' profile + profile-driven EXCLUDED_ORG_IDS/AEM_UNIT/org sync. 1.0.0:
 *   per-tenant config to Script Properties; INDEFINITE not LIFE.)
 *   See PCR_CHANGELOG.md.
 ***********************************************/

/**
 * Service account credentials for per-user impersonation are stored in Script
 * Properties (never in source) and read directly by getImpersonatedToken_()
 * in UpdateMembers.gs, which signs the delegation JWT at runtime:
 *   - SA_IMPERSONATION_EMAIL  (service account client_email)
 *   - SA_PRIVATE_KEY          (private key PEM; may contain literal \n sequences)
 *   - SA_PRIVATE_KEY_ID       (optional)
 *
 * Each Workspace tenant uses its own dedicated service account, with
 * domain-wide delegation limited to the Gmail settings and Calendar scopes the
 * automation actually needs.
 */

// ============================================================================
// PER-TENANT IDENTITY (Script Properties — NOT literals in this file)
// ============================================================================
/**
 * Loads this project's per-tenant identity from Script Properties.
 *
 * WHY THIS EXISTS: config.gs is shared by all three tenants and is OVERWRITTEN on
 * every `clasp push` (all targets use rootDir: ../src, and .claspignore ships all
 * of src/). Tenant-specific values therefore CANNOT live as literals here — a push
 * would clobber them (this is what happened to the cadets project). They must live
 * in Script Properties, which `clasp push` never touches.
 *
 * Populate them once per project with setupTenantConfig() (or by hand in
 * Project Settings > Script Properties). Canonical non-secret values for each
 * tenant are version-controlled in config-tenants/<tenant>.json.
 *
 * There is intentionally NO fallback to another tenant's values: an unconfigured
 * project yields empty identity fields and fails loudly, rather than silently
 * operating on the wrong domain.
 */
function getTenantConfig_() {
  const p = PropertiesService.getScriptProperties();
  const get = function (key, fallback) {
    const v = p.getProperty(key);
    return (v === null || String(v).trim() === '') ? (fallback || '') : String(v).trim();
  };
  return {
    DOMAIN: get('TENANT_DOMAIN'),
    EMAIL_DOMAIN: get('TENANT_EMAIL_DOMAIN'),
    SECONDARY_EMAIL_DOMAIN: get('TENANT_SECONDARY_EMAIL_DOMAIN'),
    CAPWATCH_ORGID: get('TENANT_CAPWATCH_ORGID'),
    WING: get('TENANT_WING'),
    REGION: get('TENANT_REGION'),
    CAPWATCH_DATA_FOLDER_ID: get('TENANT_CAPWATCH_DATA_FOLDER_ID'),
    REGION_CAPWATCH_DATA_FOLDER_ID: get('TENANT_REGION_CAPWATCH_DATA_FOLDER_ID'),
    AUTOMATION_FOLDER_ID: get('TENANT_AUTOMATION_FOLDER_ID'),
    AUTOMATION_SPREADSHEET_ID: get('TENANT_AUTOMATION_SPREADSHEET_ID'),
    RETENTION_LOG_SPREADSHEET_ID: get('TENANT_RETENTION_LOG_SPREADSHEET_ID'),
    RETENTION_EMAIL: get('TENANT_RETENTION_EMAIL'),
    DIRECTOR_RECRUITING_EMAIL: get('TENANT_DIRECTOR_RECRUITING_EMAIL'),
    AUTOMATION_SENDER_EMAIL: get('TENANT_AUTOMATION_SENDER_EMAIL'),
    SENDER_NAME: get('TENANT_SENDER_NAME', 'CAP Information Technology'),
    TEST_EMAIL: get('TENANT_TEST_EMAIL'),
    ITSUPPORT_EMAIL: get('TENANT_ITSUPPORT_EMAIL')
  };
}

const TENANT = getTenantConfig_();

// ============================================================================
// PER-TENANT BEHAVIORAL PROFILE (also Script-Property driven)
// ============================================================================
/**
 * Some config is not just different *identity* per tenant, it is different
 * *behavior*: the cadets tenant processes only CADET members, runs cadet-lite,
 * and creates a smaller set of squadron groups. Those values live here as coded
 * profiles (so the structure stays version-controlled) and are selected by the
 * `TENANT_PROFILE` Script Property ('seniors' | 'cadets'; defaults to 'seniors').
 *
 * Like TENANT_*, the selector is a Script Property, so a shared-config `clasp
 * push` cannot clobber a tenant's behavior — set `TENANT_PROFILE=cadets` on the
 * cadets project and it keeps cadet behavior across every push.
 */
const TENANT_PROFILE = String(
  PropertiesService.getScriptProperties().getProperty('TENANT_PROFILE') || 'seniors'
).trim().toLowerCase();

const TENANT_PROFILES_ = {
  seniors: {
    MEMBER_TYPES_ACTIVE: ['', 'SENIOR', 'FIFTY YEAR', 'INDEFINITE', 'CADET SPONSOR', ''],
    CADET_LITE: false,
    TRANSITION_ROLE: 'destination',     // receives ex-cadets; only exempts them from Level I
    EXCLUDED_ORG_IDS: ['1297', '368'], // CA-000 (1297) + CA-999 (368) holding units
    AEM_UNIT: '',                      // no Aerospace Education Member unit
    SYNC_ORG_PATHS: true,              // multi-unit wing: auto-map new squadrons
    RUN_REGION_GROUP_CHATS: false,     // region-only feature (updateRegionGroupChats)
    RUN_UNIT_VISIT_REPORT: false,      // region-only feature (buildRegionUnitVisitReport)
    RUN_SHARED_CONTACTS: false,        // wing handles shared contacts in a separate project
    RUN_AUTOMATION_CHAT_SPACES: false, // automation + user-additions chat spaces (region)
    // Mail squadron commanders when a member's FBI background check (Member.txt
    // LSCode) is granted or lapses. On here: this is the tenant that carries the
    // fingerprinted members.
    RUN_LSCODE_NOTIFICATIONS: true,
    SQUADRON_ACCESS_GROUP_AUTO_CREATE: true,
    SQUADRON_PUBLIC_CONTACT_AUTO_CREATE: true,
    SQUADRON_DISTRIBUTION_TYPES: [
      { suffix: 'allhands', name: 'All Hands', description: 'All members (cadets and seniors)', includeTypes: ['CADET', 'SENIOR', 'FIFTY YEAR', 'INDEFINITE', 'CADET SPONSOR'] },
      { suffix: 'cadets', name: 'Cadets', description: 'Cadet members only', includeTypes: ['CADET'] },
      { suffix: 'seniors', name: 'Seniors', description: 'Senior members only', includeTypes: ['SENIOR', 'FIFTY YEAR', 'INDEFINITE', 'CADET SPONSOR'] },
      { suffix: 'parents', name: 'Parents & Guardians', description: 'Parent and guardian contacts for cadet members', isParentList: true }
    ],
    // Which squadron distribution/command-staff lists this tenant creates
    // (consumed by SquadronGroups.gs getSquadronDistributionToggles_). A full
    // composite wing: every list type.
    SQUADRON_DISTRIBUTION_TOGGLES: {
      PUBLIC_CONTACT: false,
      ALLHANDS: true,
      CADETS: true,
      SENIORS: true,
      PARENTS: true,
      COMMANDER: true,
      DEPUTY_COMMANDER: true,
      DEPUTY_COMMANDER_CADETS: true,
      DEPUTY_COMMANDER_SENIORS: true
    },
    // Cross-tenant Domain Shared Contacts (src/cross-tenant-contacts). The
    // seniors tenant publishes the CADET roster into its own GAL. Peer identity
    // + service-account creds are Script Properties (XT_PEER_*); see
    // validateCrossTenantConfig(). Sheet-free: roster from CAPWATCH, authoritative
    // email from the peer directory, fallback from CAPWATCH MbrContact.
    CROSS_TENANT: {
      RUN_INBOUND: true,
      RUN_PARENTS: true,                 // publish cadet-squadron *.parents groups into the senior GAL
      PEER_TYPES: ['CADET'],             // peer members to publish
      PEER_LABEL: 'CADET',               // notes tag on managed contacts
      EMIT_PLACEHOLDERS: true            // include no-email peers as do.not.contact sentinels
    }
  },
  cadets: {
    MEMBER_TYPES_ACTIVE: ['', 'CADET', ''],
    CADET_LITE: true,
    TRANSITION_ROLE: 'source',          // owns the transition lifecycle end to end
    EXCLUDED_ORG_IDS: ['1297', '368'], // same CA holding units as seniors
    AEM_UNIT: '',
    SYNC_ORG_PATHS: true,
    RUN_REGION_GROUP_CHATS: false,
    RUN_UNIT_VISIT_REPORT: false,
    RUN_SHARED_CONTACTS: false,
    RUN_AUTOMATION_CHAT_SPACES: false,
    // Off: cadet records carry no LSCode, so there is nothing here to report on.
    // This holds for cadets over 18 too — checked against CAPID 612148 (a cadet
    // aged 18+), whose LSCode is blank. That is worth knowing rather than
    // assuming: 18+ cadets do undergo a background check, so LSCode evidently
    // tracks the senior-side record rather than "has a check on file" for
    // everyone. Do not read the column as a universal background-check flag.
    RUN_LSCODE_NOTIFICATIONS: false,
    SQUADRON_ACCESS_GROUP_AUTO_CREATE: false,
    SQUADRON_PUBLIC_CONTACT_AUTO_CREATE: false,
    SQUADRON_DISTRIBUTION_TYPES: [
      { suffix: 'allhands', name: 'All Hands', description: 'All cadets', includeTypes: ['CADET'] },
      { suffix: 'cadets', name: 'Cadets', description: 'Cadet members only', includeTypes: ['CADET'] },
      { suffix: 'parents', name: 'Parents & Guardians', description: 'Parent and guardian contacts for cadet members', isParentList: true }
    ],
    // Cadets tenant = all-hands + cadets + parents lists. No senior-side lists:
    //  - SENIORS: no senior members exist on this tenant.
    //  - COMMANDER / DEPUTY_COMMANDER*: senior duty positions whose holders have
    //    wing accounts, not cadet-tenant accounts, so the lists would be empty.
    //  - ALLHANDS: kept on purpose. On a cadet-only tenant it is effectively a
    //    duplicate of CADETS (both = all cadets), but the .all lists already
    //    exist and are retained in case something references them.
    SQUADRON_DISTRIBUTION_TOGGLES: {
      PUBLIC_CONTACT: false,
      ALLHANDS: true,
      CADETS: true,
      SENIORS: false,
      PARENTS: true,
      COMMANDER: false,
      DEPUTY_COMMANDER: false,
      DEPUTY_COMMANDER_CADETS: false,
      DEPUTY_COMMANDER_SENIORS: false
    },
    // Cross-tenant Domain Shared Contacts: the cadets tenant publishes the
    // SENIOR roster into its own GAL. Mirror image of the seniors profile.
    CROSS_TENANT: {
      RUN_INBOUND: true,
      RUN_PARENTS: false,
      PEER_TYPES: ['SENIOR', 'FIFTY YEAR', 'INDEFINITE', 'CADET SPONSOR'],
      PEER_LABEL: 'SENIOR',
      EMIT_PLACEHOLDERS: true
    }
  },
  // Pacific Region — single-unit region HQ (PCR-PCR-001). Runs the shared code,
  // differentiated only by config. Region confirmed: no AEM automation, and all
  // senior members are typed INDEFINITE (legacy LIFE dropped). See docs/PACIFIC_DIFF.md.
  pacific: {
    MEMBER_TYPES_ACTIVE: ['', 'SENIOR', 'FIFTY YEAR', 'INDEFINITE', 'CADET', ''],
    CADET_LITE: false,
    TRANSITION_ROLE: '',          // single tenant holds both cadets and seniors: nothing to cross
    EXCLUDED_ORG_IDS: ['1345'],   // PCR holding unit
    AEM_UNIT: '',                 // region does not run AEM automation
    SYNC_ORG_PATHS: false,        // single unit: no subordinate orgs to auto-map
    RUN_REGION_GROUP_CHATS: true,      // region duty groups + duty chat spaces
    RUN_UNIT_VISIT_REPORT: true,       // region-wide unit visit report
    RUN_SHARED_CONTACTS: true,         // External Contacts sheet -> Domain Shared Contacts
    RUN_AUTOMATION_CHAT_SPACES: true,  // automation + user-additions chat spaces
    // Off pending a decision: this is a single-unit region HQ, so the only
    // "squadron commander" is the region commander, sitting alongside the ~50
    // members a digest would describe. Flip to true if PCR/CC wants it.
    RUN_LSCODE_NOTIFICATIONS: false,
    SQUADRON_ACCESS_GROUP_AUTO_CREATE: false,
    SQUADRON_PUBLIC_CONTACT_AUTO_CREATE: false,
    SQUADRON_DISTRIBUTION_TYPES: [], // no subordinate squadrons
    // Single-unit region HQ: squadron group sync is not triggered here, and there
    // are no subordinate squadrons, so every squadron list type is off.
    SQUADRON_DISTRIBUTION_TOGGLES: {
      PUBLIC_CONTACT: false,
      ALLHANDS: false,
      CADETS: false,
      SENIORS: false,
      PARENTS: false,
      COMMANDER: false,
      DEPUTY_COMMANDER: false,
      DEPUTY_COMMANDER_CADETS: false,
      DEPUTY_COMMANDER_SENIORS: false
    },
    // Region has no peer tenant to cross-publish; cross-tenant sync is off.
    CROSS_TENANT: {
      RUN_INBOUND: false,
      RUN_PARENTS: false,
      PEER_TYPES: [],
      PEER_LABEL: '',
      EMIT_PLACEHOLDERS: false
    }
  }
};

const PROFILE_ = TENANT_PROFILES_[TENANT_PROFILE] || TENANT_PROFILES_.seniors;

/**
 * ONE-TIME per project: fill in THIS tenant's values below, run once, done.
 * These are written to Script Properties, which survive every `clasp push`.
 * Blank fields are skipped (never overwrite an existing property), so it is safe
 * to re-run. See config-tenants/<tenant>.json for the canonical values to paste.
 */
function setupTenantConfig() {
  const values = {
    TENANT_DOMAIN: '',                     // e.g. cawgcap.org  (cadets: cawgcadets.org)
    TENANT_EMAIL_DOMAIN: '',               // e.g. @cawgcap.org
    TENANT_SECONDARY_EMAIL_DOMAIN: '',     // e.g. @cawg.cap.gov (bare 'cawg.cap.gov' also accepted); '' unless the tenant has a verified secondary domain
    TENANT_CAPWATCH_ORGID: '',             // e.g. 188
    TENANT_WING: '',                       // e.g. CA
    TENANT_REGION: '',                     // '' unless this project is a Region-level pull
    TENANT_CAPWATCH_DATA_FOLDER_ID: '',
    TENANT_REGION_CAPWATCH_DATA_FOLDER_ID: '', // Region-level all-members CAPWATCH folder (region tenants only)
    TENANT_AUTOMATION_FOLDER_ID: '',
    TENANT_AUTOMATION_SPREADSHEET_ID: '',
    TENANT_RETENTION_LOG_SPREADSHEET_ID: '',
    TENANT_RETENTION_EMAIL: '',
    TENANT_DIRECTOR_RECRUITING_EMAIL: '',
    TENANT_AUTOMATION_SENDER_EMAIL: '',
    TENANT_SENDER_NAME: '',
    TENANT_TEST_EMAIL: '',
    TENANT_ITSUPPORT_EMAIL: ''
  };

  const props = PropertiesService.getScriptProperties();
  const applied = [];
  Object.keys(values).forEach(function (k) {
    const v = String(values[k] == null ? '' : values[k]).trim();
    if (v !== '') { props.setProperty(k, v); applied.push(k); }
  });

  console.log('✅ Applied tenant properties: ' + JSON.stringify(applied));
  console.log('ℹ️  Blank fields were skipped (existing values untouched).');
  console.log('➡️  Now run validateTenantConfig() to confirm.');
  return applied;
}

/**
 * Reports whether this project's required tenant properties are set.
 * Run after setupTenantConfig() (or after entering properties by hand), and
 * ALWAYS before pushing the shared config to a freshly-created project.
 */
function validateTenantConfig() {
  const required = [
    'TENANT_DOMAIN', 'TENANT_EMAIL_DOMAIN', 'TENANT_CAPWATCH_ORGID', 'TENANT_WING',
    'TENANT_CAPWATCH_DATA_FOLDER_ID', 'TENANT_AUTOMATION_FOLDER_ID', 'TENANT_AUTOMATION_SPREADSHEET_ID'
  ];
  const p = PropertiesService.getScriptProperties();
  const missing = required.filter(function (k) {
    const v = p.getProperty(k);
    return v === null || String(v).trim() === '';
  });

  if (missing.length) {
    console.error('❌ Missing required tenant properties: ' + missing.join(', '));
    console.error('   Set them in Project Settings > Script Properties, or run setupTenantConfig().');
  } else {
    console.log('✅ Tenant config OK — DOMAIN=' + p.getProperty('TENANT_DOMAIN') +
                ', ORGID=' + p.getProperty('TENANT_CAPWATCH_ORGID') +
                ', WING=' + p.getProperty('TENANT_WING'));
  }
  return { ok: missing.length === 0, missing: missing };
}

const CONFIG = {
/**
 * Configuration Constants
 *
 * Centralized configuration for CAPWATCH automation system.
 * Update these values to match your organization's settings.
 */

/**
 * Per-tenant identity — sourced from Script Properties via getTenantConfig_().
 * Do NOT hard-code tenant values here; a `clasp push` overwrites this file.
 * Set them per project with setupTenantConfig(). See config-tenants/<tenant>.json.
 */

/** CAPWATCH organization ID for data download (this project's Wing/Region ORGID). */
CAPWATCH_ORGID: TENANT.CAPWATCH_ORGID,

/** Region three-letter abbreviation; '' for a Wing-level project. */
REGION: TENANT.REGION,

/** Wing abbreviation, used for building squadron identifiers. */
WING: TENANT.WING,

/** Email domain for CAP accounts (members get username + this). */
EMAIL_DOMAIN: TENANT.EMAIL_DOMAIN,

/**
 * Optional secondary domain, already verified in this tenant, used to give
 * accounts a parallel address with the same local part (see
 * accounts-and-groups/SecondaryDomainAliases.gs). '' disables that module.
 */
SECONDARY_EMAIL_DOMAIN: TENANT.SECONDARY_EMAIL_DOMAIN,

/** Google Workspace domain used for API calls. */
DOMAIN: TENANT.DOMAIN,

// ============================================================================
// ACCOUNTS AND GROUPS
// ============================================================================

/**
* Member type definitions
* Determines which member types are processed in different scenarios
*/
MEMBER_TYPES: {
  /**
   * Active member types processed for this tenant — selected by TENANT_PROFILE.
   * Seniors: SENIOR/FIFTY YEAR/INDEFINITE/CADET SPONSOR. Cadets: CADET only.
   * NOTE: 'INDEFINITE' (not 'LIFE') is the correct current CAPWATCH type.
   */
  ACTIVE: PROFILE_.MEMBER_TYPES_ACTIVE,
  /** Only Aerospace Education Members */
  AEM_ONLY: ['AEM']
},

/**
 * Cadet-lite mode (selected by TENANT_PROFILE):
 * When enabled, low grades (CADET, C/Amn, C/A1C, C/SrA) are excluded from
 * Workspace account creation. On for the cadets tenant, off for seniors.
 */
CADET_LITE: PROFILE_.CADET_LITE,

/**
 * Per-target resource management (aircraft & vehicles)
 * Set Script Property MANAGE_RESOURCES=false on any script that should NOT
 * create or update calendar resources (e.g. the Cadets script — resources
 * are managed exclusively by the Seniors script).
 * Omitting the property, or setting it to anything other than "false", enables
 * resource management (default behaviour).
 */

/**
 * Cadet grades that should NOT get Workspace accounts
 */
CADET_LITE_EXCLUDED_GRADES: [
    'CADET',
    'C/Amn',
    'C/A1C',
    'C/SrA'
],

  /**
   * Number of days to wait before suspending expired members
   * Members who expire will remain active for this many days before suspension
   */
  SUSPENSION_GRACE_DAYS: 7,

  /**
   * Organization IDs whose members are suspended and (via shouldProcessMember()
   * in UpdateMembers.gs) never eligible for accounts or group membership,
   * regardless of member type. Holding units differ by wing/region, so this is
   * per-tenant via TENANT_PROFILE. Seniors/cadets: 1297 = CALIF WING HQ SQ
   * (CA-000) + 368 = CALIF LEGISLATIVE SQ (CA-999). Pacific: 1345.
   */
  EXCLUDED_ORG_IDS: PROFILE_.EXCLUDED_ORG_IDS,

  /**
   * Special organization configurations
   */
  SPECIAL_ORGS: {
    /**
     * Artificial org ID for Aerospace Education Members (per-tenant, via
     * TENANT_PROFILE). Currently empty on every tenant — none run AEM
     * automation. When empty, an inert squadron keyed by '' is created and
     * matches no member (the existing behavior). Set it in a profile only if a
     * tenant starts running AEM automation.
     */
    AEM_UNIT: PROFILE_.AEM_UNIT
  },

  /**
   * When true, senior members without a completed Level I achievement in
   * MbrAchievements will not receive a new Workspace account. Existing
   * accounts are unaffected — only new provisioning is gated.
   *
   * Exception: a member transitioning up from the cadet tenant already has an
   * account and is exempt. See TRANSITION_CONFIG and updateAllMembers().
   */
  REQUIRE_LEVEL_I_FOR_SENIORS: true,

  /**
   * Number of members to process in each batch
   * Used by batchUpdateMembers() to group API calls
   */
  BATCH_SIZE: 50,

  /**
   * Maximum number of users to process in a single execution
   * Safety limit to prevent runaway processing
   *
   * If more users need processing, they'll be handled in the next run.
   * Set to a reasonable limit based on your user volume and script timeout.
   */
  MAX_BATCH_SIZE: 100,

  /**
   * Maximum number of retry attempts for API calls
   * Used by executeWithRetry() in utils.gs
   */
  /**
   * Maximum number of retry attempts for API calls
   * Used by executeWithRetry() in utils.gs
   */
  API_RETRY_ATTEMPTS: 3,

  /**
   * Delay (milliseconds) between each API call to prevent 403 quota errors
   */
  API_DELAY_MS: 250,

  /**
   * Base delay (milliseconds) for exponential backoff after quota errors
   */
  API_BACKOFF_BASE_MS: 3000,

  // ------------------------------------------------------
  // FOLDER + SHEET IDS
  // ------------------------------------------------------

  /**
   * Google Drive folder ID where CAPWATCH data files are stored (per tenant).
   * Downloaded files (Member.txt, Organization.txt, etc.) go here.
   */
  CAPWATCH_DATA_FOLDER_ID: TENANT.CAPWATCH_DATA_FOLDER_ID,

  /**
   * Region-level all-members CAPWATCH data folder (region tenants only).
   * Used by updateRegionGroupChats(); empty on wing tenants.
   */
  REGION_CAPWATCH_DATA_FOLDER_ID: TENANT.REGION_CAPWATCH_DATA_FOLDER_ID,

  /**
   * Google Drive folder ID for automation files (per tenant).
   * Contains configuration spreadsheets and logs.
   */
  AUTOMATION_FOLDER_ID: TENANT.AUTOMATION_FOLDER_ID,

  /**
   * Google Sheets ID for automation configuration (per tenant).
   * Contains 'Groups', 'User Additions', 'Error Emails' sheets.
   */
  AUTOMATION_SPREADSHEET_ID: TENANT.AUTOMATION_SPREADSHEET_ID,

  /**
   * Log level for automation scripts
   * Options:
   *   - 'INFO'  → show all logs (default)
   *   - 'WARN'  → warnings and errors only
   *   - 'ERROR' → errors only
   *   - 'NONE'  → suppress all logs
   */
  LOG_LEVEL: 'INFO',
};

/**
 * HTTP Error Code Constants
 * Standard error codes from Google Admin API
 * Used for consistent error handling across scripts
 */
const ERROR_CODES = {
  /** Invalid request (bad parameters, malformed data) */
  BAD_REQUEST: 400,

  /** Insufficient permissions */
  FORBIDDEN: 403,

  /** Resource not found (user, group, etc.) */
  NOT_FOUND: 404,

  /** Resource already exists or conflict */
  CONFLICT: 409,

  /** Google server error (usually transient) */
  SERVER_ERROR: 500
};

/**
 * Maximum number of members to retrieve per page from Admin API
 * Google's recommended value is 200 for optimal performance
 * Higher values may cause timeouts
 */
const GROUP_MEMBER_PAGE_SIZE = 200;


// ============================================================================
// RECRUITING AND RETENTION
// ============================================================================

/**
 * Configuration for Retention Automation Emails
 * These values are used by retention email scripts (if implemented)
 */

/* Per-tenant recruiting/retention values — sourced from Script Properties
 * via TENANT (see getTenantConfig_()). Do NOT hard-code tenant values here. */

/** Google Sheets ID for retention tracking log */
const RETENTION_LOG_SPREADSHEET_ID = TENANT.RETENTION_LOG_SPREADSHEET_ID;

/** Email address for retention Google Group */
const RETENTION_EMAIL = TENANT.RETENTION_EMAIL;

/** Email address for Director of Recruiting and Retention */
const DIRECTOR_RECRUITING_EMAIL = TENANT.DIRECTOR_RECRUITING_EMAIL;

/** Email alias to use as sender for automated emails */
const AUTOMATION_SENDER_EMAIL = TENANT.AUTOMATION_SENDER_EMAIL;

/** Display name for automated email sender */
const SENDER_NAME = TENANT.SENDER_NAME;

/** Test email address for development/testing */
const TEST_EMAIL = TENANT.TEST_EMAIL;

/** IT support mailbox for notifications */
const ITSUPPORT_EMAIL = TENANT.ITSUPPORT_EMAIL;

/**
 * Configuration for retention email system
 * Centralized constants for email subjects and thresholds
 */
const RETENTION_CONFIG = {
  /**
   * Email subject lines
   */
  SUBJECTS: {
    TURNING_18: 'Important Membership Update - Turning 18',
    TURNING_21: 'Important Membership Update - Turning 21',
    EXPIRING: 'Your CAP Membership Expires Soon'
  },

  /**
   * Age thresholds for email triggers
   */
  AGE_THRESHOLDS: {
    TRANSITION_TO_SENIOR: 18,
    CADET_AGE_OUT: 21
  },

  /**
   * Email rate limiting (milliseconds between sends)
   */
  EMAIL_DELAY_MS: 100,

  /**
   * Progress logging frequency (log every N emails)
   */
  PROGRESS_LOG_INTERVAL: 10
};

// ============================================================================
// CADET → SENIOR TENANT TRANSITION
// ============================================================================

/**
 * Cadet-to-senior tenant transition.
 *
 * A cadet who turns 21 (or converts voluntarily after 18) leaves the cadet
 * tenant and reappears on the senior tenant as a new CAPWATCH type. Without
 * intervention their cadet mailbox is suspended, then permanently deleted, and
 * the mail is gone — Archived User licenses are not available on this edition,
 * so deletion is the only way to free a seat and there is nothing to fall back
 * on. This config drives the module that carries their mail across first.
 *
 * ROLE (per profile):
 *   'source'      — cadets tenant. Owns the whole lifecycle: detect, migrate,
 *                   delete, forward. Reads the cadet mailbox with its own SA and
 *                   writes into the senior mailbox with the peer SA.
 *   'destination' — seniors tenant. Only exempts transitioning members from the
 *                   Level I gate so the destination account exists to receive
 *                   mail. No migration work happens here.
 *   ''            — off (pacific).
 *
 * Deliberately single-owner: the source tenant polls the peer directory for the
 * destination account rather than the two tenants signaling each other, so there
 * is no shared state to get out of sync.
 */
const TRANSITION_CONFIG = {
  ROLE: PROFILE_.TRANSITION_ROLE,

  /**
   * CAPWATCH types that mean "left the cadet program but is still a member".
   *
   * PATRON belongs here even though patrons get no Workspace account. A cadet
   * who ages out while National is still processing their fingerprint cards
   * lands on PATRON and converts to SENIOR weeks later; if PATRON did not hold
   * the mailbox open, that mail would be deleted before the conversion landed.
   * Holding it costs a suspended seat, which still counts against the 2000-user
   * cap — see HOLD_DAYS.
   */
  TRANSITION_TYPES: ['SENIOR', 'FIFTY YEAR', 'INDEFINITE', 'CADET SPONSOR', 'PATRON'],

  /**
   * Types that can actually receive migrated mail (i.e. get a senior account).
   * A PATRON row waits at PENDING until they convert to one of these or the
   * hold expires.
   */
  DESTINATION_TYPES: ['SENIOR', 'FIFTY YEAR', 'INDEFINITE', 'CADET SPONSOR'],

  /**
   * Days from detection before the cadet account is deleted.
   *
   * This clock is authoritative and is measured from the Transitions sheet's
   * DetectedDate. It deliberately does NOT reuse
   * LICENSE_CONFIG.DAYS_BEFORE_DELETE_INELIGIBLE, which times a different event:
   * a member lapsing. A transitioning cadet has not lapsed — they are converting
   * — so that clock says nothing useful about when their mail is safe to destroy.
   * Only this module knows whether the mail is across yet, hence the separate
   * timer and the getHeldTransitionCapids() hold in deleteIneligibleSuspendedUsers().
   *
   * 90 days covers National being slow on fingerprint processing.
   */
  HOLD_DAYS: 90,

  /**
   * Days after a VERIFIED migration before the cadet account is deleted and the
   * forwarding group takes its address.
   *
   * HOLD_DAYS exists to wait out National's fingerprint processing — i.e. for
   * members who may yet convert. Once someone has converted and their mail is
   * across, that purpose is served and the remaining wait is pure cost: a Google
   * Group cannot share an address with a User, so the forward cannot exist until
   * the account is gone. Waiting the full 90 days would leave the old address
   * dead that whole time.
   *
   * The 14-day buffer is there to catch a migration that reported success but
   * was not, while the source mailbox still exists to re-run from. Deletion is
   * permanent on this edition.
   */
  POST_MIGRATION_DELETE_DAYS: 14,

  /**
   * Months the old cadet address forwards to the new senior address after the
   * cadet account is deleted. Implemented as a Group at the old address, which
   * costs no license seat (unlike keeping the account alive).
   */
  FORWARD_GROUP_MONTHS: 12,

  /** Subject of the transition-complete notification. */
  EMAIL_SUBJECT: 'Your CAP email has moved to your senior account',

  /**
   * When true, refuse to delete a cadet account whose migration has not
   * confirmed success, and alert instead. An account lingering against the cap
   * is recoverable; a mailbox deleted after a silent migration failure is not.
   */
  REQUIRE_MIGRATION_BEFORE_DELETE: true,

  /** Tab in the automation spreadsheet holding transition state. */
  SHEET_NAME: 'Transitions',

  /** Migration lifecycle states (MigrationStatus column). */
  STATUS: {
    PENDING: 'PENDING',           // detected; destination account does not exist yet
    IN_PROGRESS: 'IN_PROGRESS',   // migration started, resumable via LastCursor
    COMPLETE: 'COMPLETE',         // mail moved and verified; safe to delete
    FAILED: 'FAILED',             // migration errored; blocks deletion, needs a human
    NOT_APPLICABLE: 'NOT_APPLICABLE' // hold expired with no destination (e.g. stayed PATRON)
  }
};

// ============================================================================
// LICENSE MANAGEMENT
// ============================================================================

/**
 * License Management Configuration
 * Controls the lifecycle management of Google Workspace accounts
 */
const LICENSE_CONFIG = {
  /**
   * Number of days a user must be suspended before being archived.
   * NOTE: Archiving requires Archived User licenses which are not provisioned
   * on this Workspace for Nonprofits edition (confirmed via 412 errors).
   * archiveLongSuspendedUsers() is effectively a no-op on this domain.
   */
  DAYS_BEFORE_ARCHIVE: 365,

  /**
   * Number of days a user must be archived before being deleted.
   * Used by deleteLongArchivedUsers() (currently commented out in
   * manageLicenseLifecycle() since archiving is unavailable).
   */
  DAYS_BEFORE_DELETE: 1825, // 5 years

  /**
   * Number of days a suspended, ineligible user is kept before deletion.
   * Ineligible means: suspended AND not an eligible active CAPWATCH member
   * (not in CONFIG.MEMBER_TYPES.ACTIVE with ACTIVE status, and not a Manual
   * Member). This timer controls how long before the account is permanently
   * deleted to free a seat against the 2000-user domain cap.
   *
   * Measured from the member's CAPWATCH Expiration date (Member.txt column 16),
   * i.e. from when they stopped being a member — NOT from lastLoginTime, which
   * measures something else entirely (a member who simply stopped reading their
   * mail is not a member who lapsed). Workspace exposes no suspension
   * timestamp, and suspension follows expiry on the next sync, so expiry is the
   * closest authoritative date to the event this grace period is about.
   *
   * If a member becomes eligible again before this deadline (e.g. a PATRON
   * upgrades to SENIOR), the account is automatically unsuspended by
   * reactivateRenewedMembers() and the deletion never fires.
   */
  DAYS_BEFORE_DELETE_INELIGIBLE: 30,

  /**
   * Minimum plausible data row count for CAPWATCH Member.txt.
   *
   * deleteIneligibleSuspendedUsers() treats "no CAPWATCH record" as proof that
   * a member lapsed long ago, because the extract only retains roughly three
   * months of expired members. That inference inverts dangerously if the file
   * is short: a truncated or partially downloaded Member.txt makes thousands of
   * current members look like they have no record, and the reaper would delete
   * their mailboxes. parseFile() also has a fallback parser that can quietly
   * return a partial row set.
   *
   * So the deletion path refuses to run below this floor. The seniors extract
   * carries ~5,000 rows; this is set well under that to catch gross truncation
   * without tripping on normal membership churn.
   */
  MIN_MEMBER_ROWS: 1000,

  /**
   * Email addresses to receive license management reports
   * These recipients will get monthly reports of:
   * - Users reactivated
   * - Users archived
   * - Users deleted
   * - Any errors encountered
   */
  NOTIFICATION_EMAILS: [
    DIRECTOR_RECRUITING_EMAIL,  // Primary contact
    AUTOMATION_SENDER_EMAIL,     // Backup/monitoring
    ITSUPPORT_EMAIL   // IT Notification
  ],

};

// ============================================================================
// SQUADRON GROUPS CONFIGURATION
// ============================================================================

/**
 * Configuration for Squadron-Level Groups
 * Controls automatic creation and management of squadron groups
 */
const SQUADRON_GROUP_CONFIG = {
  /**
   * Access Group Configuration (ag.mixxx@miwg.cap.gov)
   * Internal access groups for shared drives and resources
   */
  ACCESS_GROUP: {
    /**
     * Description template for access groups
     */
    DESCRIPTION_TEMPLATE: 'Internal access group for {squadron}. CAWG accounts only. Used for shared drive permissions and internal resource access.',
    
    /**
     * Whether to auto-create access groups for all squadrons (per TENANT_PROFILE:
     * on for seniors, off for cadets).
     */
    AUTO_CREATE: PROFILE_.SQUADRON_ACCESS_GROUP_AUTO_CREATE,

    /**
     * Whether to include access groups in Global Address List
     * Set to false to keep internal access groups less visible
     */
    INCLUDE_IN_GAL: false
  },
  
  /**
   * Public Contact Group Configuration (mixxx@miwg.cap.gov)
   * External-facing email addresses for public inquiries
   */
  PUBLIC_CONTACT: {
    /**
     * Duty position codes that should be included in public contact groups
     * Includes commanders, deputy commanders, PAO, and recruiting officers
     * Both primary positions and assistants are included
     */
    DUTY_POSITIONS: [
      'Commander',
      'Deputy Commander',
      'Public Affairs Officer',
      'Deputy Commander for Seniors',
      'Deputy Commander for Cadets',
      'Recruiting Officer'
    ],
    
    /**
     * Wing-level recruiting mailbox to include in all public contact groups
     * This allows the Wing Director of Recruiting & Retention to monitor
     * squadron public contact inquiries for recruiting opportunities
     */
    RECRUITING_MAILBOX: '<recruiting email DL here>',
    
    /**
     * Description template for public contact groups
     */
    DESCRIPTION_TEMPLATE: 'Public contact email for {squadron}. For external inquiries, website contact forms, and business cards.',
    
    /**
     * Whether to auto-create public contact groups for all squadrons (per
     * TENANT_PROFILE: on for seniors, off for cadets).
     */
    AUTO_CREATE: PROFILE_.SQUADRON_PUBLIC_CONTACT_AUTO_CREATE
  },
  
  /**
   * Distribution List Configuration
   * Communication lists using preferred email addresses
   */
  DISTRIBUTION_LIST: {
    /**
     * Distribution list types to create per squadron — selected by TENANT_PROFILE.
     * Seniors: all-hands / cadets / seniors / parents. Cadets: cadets / parents only.
     * (Member-type lists use 'INDEFINITE', not 'LIFE' — see TENANT_PROFILES_.)
     *
     * NOTE: this descriptor list is not currently consumed by SquadronGroups.gs;
     * which lists get created is governed by PROFILE_.SQUADRON_DISTRIBUTION_TOGGLES
     * (see getSquadronDistributionToggles_). Keep the two in sync until they are
     * consolidated.
     */
    TYPES: PROFILE_.SQUADRON_DISTRIBUTION_TYPES,

    /**
     * Whether to auto-create distribution lists for all squadrons
     */
    AUTO_CREATE: true,
    
    /**
     * Whether to include distribution lists in Global Address List
     * Makes groups discoverable in autocomplete when composing emails
     */
    INCLUDE_IN_GAL: true
  },
  
  /**
   * Maximum execution time in milliseconds before stopping
   * Google Apps Script has a 6-minute execution limit
   * Set to 5.5 minutes (330 seconds) to allow graceful shutdown
   */
  MAX_EXECUTION_TIME_MS: 1750000,
  
  /**
   * Maximum number of groups to process in a single execution
   * Safety limit to prevent runaway processing and API quota issues
   * 
   * With 7 groups per squadron, this allows processing of ~71 squadrons
   * Adjust based on your number of squadrons and script timeout limits
   */
  MAX_GROUPS_PER_RUN: 500,
  
  /**
   * Delay between processing each squadron (in milliseconds)
   * Helps avoid API rate limits
   * Default: 200ms between squadrons
   */
  SQUADRON_PROCESSING_DELAY_MS: 200,
  
  /**
   * Whether to enable collaborative inbox for all squadron groups
   * When enabled, groups can be used as shared inboxes with conversation history
   * Recommended: true for all squadron groups
   */
  ENABLE_COLLABORATIVE_INBOX: true,
  
  /**
   * Whether to include squadron groups in Global Address List
   * Makes groups discoverable in autocomplete when composing emails
   * Recommended: true
   */
  INCLUDE_IN_GAL: true,
  
  /**
   * Default message moderation level for squadron groups
   * Options:
   * - MODERATE_NONE: No moderation (recommended for most groups)
   * - MODERATE_ALL_MESSAGES: All messages require approval
   * - MODERATE_NON_MEMBERS: Only non-member messages require approval
   * - MODERATE_NEW_MEMBERS: Messages from new members require approval
   */
  DEFAULT_MODERATION_LEVEL: 'MODERATE_NONE',
  
  /**
   * Organization units to exclude from squadron group creation
   * These org IDs will be skipped when processing squadrons
   * Typically includes holding units, test units, or administrative units
   */
  EXCLUDED_ORGIDS: [
    '744',   // MI-000 (Holding unit)
    '1920'   // MI-999 (Transition unit)
  ]
};