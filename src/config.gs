/***********************************************
 * File: config.gs
 * Description: Centralized configuration and constants for CAPWATCH automation scripts.
 * Provides organization-specific parameters, email domains, folder IDs, and time zone mapping.
 * Author: Noel Luneau
 * Date: November 14, 2025
 ***********************************************/

/**
 * Builds the Service Account JSON at runtime.
 *
 * Secrets MUST NOT be committed to source.
 * Store these in Script Properties:
 *   - SA_IMPERSONATION_EMAIL  (service account client_email)
 *   - SA_PRIVATE_KEY          (private key PEM; may contain literal \n sequences)
 *
 * Optional:
 *   - SA_PRIVATE_KEY_ID
 */
function getServiceAccountJson_() {
  const props = PropertiesService.getScriptProperties();
  const clientEmail = String(props.getProperty('SA_IMPERSONATION_EMAIL') || '').trim();
  let privateKey = String(props.getProperty('SA_PRIVATE_KEY') || '');
  const privateKeyId = String(props.getProperty('SA_PRIVATE_KEY_ID') || '').trim();

  privateKey = privateKey.replace(/\\n/g, '\n');

  if (!clientEmail || !privateKey) {
    return JSON.stringify({});
  }

  const sa = {
    type: 'service_account',
    project_id: 'pcr-capwatch',
    private_key_id: privateKeyId || undefined,
    private_key: privateKey,
    client_email: clientEmail,
    client_id: '117582666328715304137',
    auth_uri: 'https://accounts.google.com/o/oauth2/auth',
    token_uri: 'https://oauth2.googleapis.com/token',
    auth_provider_x509_cert_url: 'https://www.googleapis.com/oauth2/v1/certs',
    client_x509_cert_url: clientEmail
      ? `https://www.googleapis.com/robot/v1/metadata/x509/${encodeURIComponent(clientEmail)}`
      : undefined,
    universe_domain: 'googleapis.com'
  };

  Object.keys(sa).forEach(k => sa[k] === undefined && delete sa[k]);
  return JSON.stringify(sa, null, 2);
}

const CONFIG = {
/**
 * Configuration Constants
 *
 * Centralized configuration for CAPWATCH automation system.
 * Update these values to match your organization's settings.
 */

/**
 * CAPWATCH organization ID for data download
 * This should be your Wing ORGID
 * MI Wing = 223
 */
CAPWATCH_ORGID: '188',

/**
 * Wing abbreviation
 * Used for building squadron identifiers
 */
REGION: "",

/**
* If unit is a Region, set three letter abbreviation here
*/
WING: "CA",

/**
 * Email domain for CAP accounts
 * All members get username@miwg.cap.gov
 */
EMAIL_DOMAIN: "@cawgcap.org",

/**
 * Google Workspace domain
 * Used for API calls
 */
DOMAIN: "cawgcap.org",

// ============================================================================
// ACCOUNTS AND GROUPS
// ============================================================================

/**
* Member type definitions
* Determines which member types are processed in different scenarios
*/
MEMBER_TYPES: {
  /** All active member types */
  /** ACTIVE: ['CADET', 'SENIOR', 'FIFTY YEAR', 'INDEFINITE', ''],  */
  ACTIVE: ['', 'SENIOR', 'FIFTY YEAR', 'INDEFINITE', 'CADET SPONSOR', ''],
  /** Only Aerospace Education Members */
  AEM_ONLY: ['AEM']
},

/**
 * Cadet-lite mode:
 * When enabled:
 *   - Excludes low grades from Workspace account creation
 *   - CADET, C/Amn, C/A1C, C/SrA will be ignored
 */
CADET_LITE: false,   // turn OFF by setting false

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
   * Organization IDs that should have users suspended, and (via
   * shouldProcessMember() in UpdateMembers.gs) are never eligible for
   * accounts or group membership regardless of member type.
   * 1297 = CALIF WING HQ SQ (CA-000, the holding unit for members not
   * assigned to a squadron); 368 = CALIFORNIA LEGISLATIVE SQ (CA-999).
   */
  EXCLUDED_ORG_IDS: ['1297', '368'],

  /**
   * Special organization configurations
   */
  SPECIAL_ORGS: {
    /**
     * Artificial org ID for Aerospace Education Members
     * Uses MIWG (223) as template but with separate identity
     */
    AEM_UNIT: ''
  },

  /**
   * When true, senior members without a completed Level I achievement in
   * MbrAchievements will not receive a new Workspace account. Existing
   * accounts are unaffected — only new provisioning is gated.
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
  // SERVICE ACCOUNT JSON
  // ------------------------------------------------------
  SERVICE_ACCOUNT_JSON: getServiceAccountJson_(),

  // ------------------------------------------------------
  // FOLDER + SHEET IDS
  // ------------------------------------------------------

  /**
   * Google Drive folder ID where CAPWATCH data files are stored
   * Downloaded files (Member.txt, Organization.txt, etc.) go here
   */
  CAPWATCH_DATA_FOLDER_ID: '10T0wBubqzUzHa_7nx__eNfuzhTpFRDs3',

  /**
   * Google Drive folder ID for automation files
   * Contains configuration spreadsheets and logs
   */
  AUTOMATION_FOLDER_ID: '1lLUs0RsTQXNgRnt_fURsw8B3E8DpsgE2',

  /**
   * Google Sheets ID for automation configuration
   * Contains 'Groups', 'User Additions', 'Error Emails' sheets
   */
  AUTOMATION_SPREADSHEET_ID: '1UqCc6aRMEYw-Y_bTcTDKXuaYLsQ6bQzkdoVG7rRsV9Q',

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

/** Google Sheets ID for retention tracking log */
const RETENTION_LOG_SPREADSHEET_ID = '1ouL6YHtTfpJs32YQ2NyfYxjHSDg39RydHMamGHXM7yA';

/** Email address for retention Google Group */
const RETENTION_EMAIL = 'recruiting@cawgcap.org';

/** Email address for Director of Recruiting and Retention */
const DIRECTOR_RECRUITING_EMAIL = 'adam.staley@cawgcap.org';

/** Email alias to use as sender for automated emails */
const AUTOMATION_SENDER_EMAIL = 'automation@cawgcap.org';

/** Display name for automated email sender */
const SENDER_NAME = 'CAWG Information Technology';

/** Test email address for development/testing */
const TEST_EMAIL = 'it@cawgcap.org';

/** IT support mailbox for notifications */
const ITSUPPORT_EMAIL = 'it@cawgcap.org'

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
   * Member). Accounts are suspended immediately on determination of
   * ineligibility; this timer controls how long before the account is
   * permanently deleted to free a seat against the 2000-user domain cap.
   *
   * If a member becomes eligible again before this deadline (e.g. a PATRON
   * upgrades to SENIOR), the account is automatically unsuspended by
   * reactivateRenewedMembers() and the deletion never fires.
   */
  DAYS_BEFORE_DELETE_INELIGIBLE: 30,

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
     * Whether to auto-create access groups for all squadrons
     */
    AUTO_CREATE: true,
    
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
     * Whether to auto-create public contact groups for all squadrons
     */
    AUTO_CREATE: true
  },
  
  /**
   * Distribution List Configuration
   * Communication lists using preferred email addresses
   */
  DISTRIBUTION_LIST: {
    /**
     * Distribution list types to create
     * Each squadron gets all of these automatically
     */
    TYPES: [
      {
        suffix: 'allhands',
        name: 'All Hands',
        description: 'All members (cadets and seniors)',
        includeTypes: ['CADET', 'SENIOR', 'FIFTY YEAR', 'INDEFINITE', 'CADET SPONSOR']
      },
      {
        suffix: 'cadets',
        name: 'Cadets',
        description: 'Cadet members only',
        includeTypes: ['CADET']
      },
      {
        suffix: 'seniors',
        name: 'Seniors',
        description: 'Senior members only',
        includeTypes: ['SENIOR', 'FIFTY YEAR', 'INDEFINITE', 'CADET SPONSOR']
      },
      {
        suffix: 'parents',
        name: 'Parents & Guardians',
        description: 'Parent and guardian contacts for cadet members',
        isParentList: true
      }
    ],
    
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