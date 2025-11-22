  /**
   * CAPWATCH organization ID for data download
   * This should be your Wing ORGID
   * MI Wing = 223
   */
  CAPWATCH_ORGID: '163',

  /**
   * Wing abbreviation
   * Used for building squadron identifiers
   */

  REGION: "",

  /**
  * If unit is a Region, set three letter abbreviation here
  */

  WING: "HI",

  /**
   * Email domain for CAP accounts
   * All members get username@miwg.cap.gov
   */
  EMAIL_DOMAIN: "@hiwgcap.org",

  /**
   * Google Workspace domain
   * Used for API calls
   */
  DOMAIN: "hiwgcap.org",

  /**
   * Google Drive folder ID where CAPWATCH data files are stored
   * Downloaded files (Member.txt, Organization.txt, etc.) go here
   */
  CAPWATCH_DATA_FOLDER_ID: '1nA1msHn1_0c9xEHTN-JOzH5joAvf6J43',

  /**
   * Google Drive folder ID for automation files
   * Contains configuration spreadsheets and logs
   */
  AUTOMATION_FOLDER_ID: '1552GOTspzZFv2NPiHONQd9LyAi0lGrWg',

  /**
   * Google Sheets ID for automation configuration
   * Contains 'Groups', 'User Additions', 'Error Emails' sheets
   */
  AUTOMATION_SPREADSHEET_ID: '1b-LyVmoBDf19YlHTvylbPzNV5VsPDrCl93sI_vbyUho'
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
const RETENTION_LOG_SPREADSHEET_ID = '<id for the spreadsheet here>';

/** Email address for retention Google Group */
const RETENTION_EMAIL = 'it@hiwgcap.org';

/** Email address for Director of Recruiting and Retention */
const DIRECTOR_RECRUITING_EMAIL = 'it@hiwgcap.org';

/** Email alias to use as sender for automated emails */
const AUTOMATION_SENDER_EMAIL = 'automation@hiwgcap.org';

/** Display name for automated email sender */
const SENDER_NAME = 'HIWG Information Technology';

/** Test email address for development/testing */
const TEST_EMAIL = 'it@hiwgcap.org';

/** IT support mailbox for notifications */
const ITSUPPORT_EMAIL = 'it@hiwgcap.org'

