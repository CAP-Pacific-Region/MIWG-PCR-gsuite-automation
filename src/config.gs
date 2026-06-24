

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
    DESCRIPTION_TEMPLATE: 'Internal access group for {squadron}. MIWG accounts only. Used for shared drive permissions and internal resource access.',
    
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
        includeTypes: ['CADET', 'SENIOR', 'FIFTY YEAR', 'LIFE']
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
        includeTypes: ['SENIOR', 'FIFTY YEAR', 'LIFE']
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
  MAX_EXECUTION_TIME_MS: 330000,
  
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