/**
 * Retention Email Automation Module
 *
 * Version: 1.4.0
 * Date: 2026-07-25
 * Changes: Two guards that had to exist before this could be put on a trigger.
 *   ALREADY-SENT: the Log sheet has always been written and never read, so
 *   nothing knew what a previous run had done — an execution that died partway
 *   through the expiring batch would restart from the top of the list on the next
 *   firing, and a manual re-run after a fix re-mailed everyone it had already
 *   reached. Sends are now filtered against (email type, CAPID) for the current
 *   calendar month. It FAILS OPEN: an unreadable log leaves the run behaving as
 *   it did before, but says so in the execution log and in a banner on the
 *   summary email, so a low send count is never ambiguous.
 *   TENANT: sendRetentionEmails() is now gated on PROFILE_.RUN_RETENTION_EMAILS
 *   (config.gs 1.12.0) — true on seniors, false on cadets and region. This module
 *   hardcodes 'CADET'/'SENIOR' instead of reading MEMBER_TYPES.ACTIVE, and both
 *   wing tenants pull the same wing-wide extract, so it addresses the entire wing
 *   from wherever it runs; arming both tenants mailed every member twice rather
 *   than splitting the work.
 *   The summary email and testRetentionEmail() both now report skipped counts.
 *   See PCR_CHANGELOG.md.
 *   1.3.0: Dropped `bcc: RETENTION_EMAIL` from all three member-facing sends. The
 *   retention group now receives the run summary only, not a copy of every
 *   message — at wing scale that BCC was a few hundred messages a month into one
 *   mailbox, and it duplicated a record the Log sheet already keeps per send
 *   (timestamp, type, CAPID, name, address, commander). RETENTION_EMAIL still
 *   addresses the summary. Note the previous value, recruiting@cawgcap.org, did
 *   not exist as a group, so the summary had nowhere to land and
 *   sendRetentionSummaryEmail()'s catch would have swallowed the failure — see
 *   config-tenants/ for the replacement role group. See PCR_CHANGELOG.md.
 *   1.2.0: Commander CC now goes to the commander's CAP account rather than the
 *   personal address on their CAPWATCH record — real Workspace account first,
 *   then the derived first.last@<command domain>, with CAPWATCH primary kept only
 *   as a last resort. This reuses rcResolveRecipientEmail_() from
 *   notifications/RecoveryEmailNotify.gs so the two modules cannot disagree about
 *   how a commander is reached. getCommanderInfo() is now backed by
 *   retentionCommanderIndex_(), built once per execution: the previous version
 *   re-walked Commanders.txt and rebuilt the entire CAPWATCH email map on every
 *   call, i.e. once per cadet email sent. See PCR_CHANGELOG.md.
 *   1.1.0: Genericized the three member-facing templates, which still carried a
 *   hard-coded 'CALIFORNIA WING' masthead/footer and a hard-coded role holder in
 *   the signature — the last thing blocking another wing from adopting this
 *   module by Script Property alone. New placeholders {{wingName}}, {{orgLabel}}
 *   and {{signature}} are filled by retentionRenderTemplate_(), which replaces
 *   the substitution chains that were duplicated across all seven render sites.
 *   The signature name comes from the new optional Script Property
 *   TENANT_DIRECTOR_RECRUITING_NAME (blank signs with the office title alone).
 *   Also removed the dead feedback-survey block from ExpiringEmail.html, which
 *   shipped with an unfilled 'LINK TO FORM HERE' placeholder in the button href,
 *   the fallback href and the visible link text — it would have mailed every
 *   expiring member a broken link. The feedback ask now routes to replyTo, which
 *   is already the Director of Recruiting. Dropped the commented-out MIWG
 *   'Phoenix Senior Flight' block and the CSS orphaned by both removals, and
 *   closed the unbalanced <p> tags in all three signature blocks.
 *   See PCR_CHANGELOG.md.
 *   1.0.1: Retention-report footer uses the programmable CONFIG.ORG_LABEL instead
 *   of literal 'CAWG', so it reads correctly for any wing.
 *   1.0.0: Reconciled with live tenant code (INDEFINITE senior type; profile-driven
 *   config references).
 *
 * Automates sending retention-focused emails to CAP members based on lifecycle events:
 * - Cadets turning 18 (transition to senior member opportunities)
 * - Cadets turning 21 (aging out of cadet program)
 * - Members with expiring memberships (renewal reminders)
 * 
 * Email Features:
 * - Personalized with member rank and name
 * - CC to squadron commander for cadet emails
 * - Reply-to set to the Director of Recruiting role group
 * - Logged to retention tracking spreadsheet (the per-send record; the retention
 *   group receives the run summary, not a copy of every message)
 * 
 * RECOMMENDED SCHEDULE: Run monthly on the 1st at 10:00 AM
 * This allows time for CAPWATCH data to be updated after month-end processing.
 * 
 * Setup Instructions:
 * 1. Create email templates: Turning18Email.html, Turning21Email.html, ExpiringEmail.html
 * 2. Set RETENTION_LOG_SPREADSHEET_ID in config.gs
 * 3. Verify TENANT_RETENTION_EMAIL and TENANT_DIRECTOR_RECRUITING_EMAIL in
 *    Script Properties (NOT config.gs — a push overwrites that file)
 * 4. Run testAllRetentionEmails() to verify templates and configuration
 * 5. Set up time-driven trigger for sendRetentionEmails()
 * 
 * Authors: Luke Bunge, luke.bunge@miwg.cap.gov
 */


/**
 * Main function to send all retention emails
 * 
 * Process:
 * 1. Clears cache for fresh CAPWATCH data
 * 2. Retrieves members in each category
 * 3. Sends personalized emails to each member
 * 4. Tracks send statistics and errors
 * 5. Logs summary to console and spreadsheet
 * 
 * This function should be scheduled to run monthly via time-driven trigger.
 * Recommended: 1st of month at 10:00 AM
 * 
 * @returns {Object} Summary of email operations with sent counts and errors
 */
function sendRetentionEmails() {
  if (!PROFILE_.RUN_RETENTION_EMAILS) {
    Logger.info('Retention emails disabled for this tenant profile', {
      profile: TENANT_PROFILE
    });
    return { skipped: true };
  }

  clearCache(); // Ensure fresh CAPWATCH data
  _commanderIndex = null; // derived from that data — must not outlive it
  const start = new Date();
  Logger.info('Starting retention email process');

  // Initialize summary tracking
  const summary = {
    sent: { turning18: 0, turning21: 0, expiring: 0 },
    failed: { turning18: [], turning21: [], expiring: [] },
    skipped: { turning18: 0, turning21: 0, expiring: 0 },
    startTime: start.toISOString()
  };

  try {
    // Get members for each category
    const turning18 = getMembersTurning18();
    const turning21 = getMembersTurning21();
    const expiring = getExpiringMembers();

    Logger.info('Member categories retrieved', {
      turning18Count: turning18.length,
      turning21Count: turning21.length,
      expiringCount: expiring.length,
      totalToProcess: turning18.length + turning21.length + expiring.length
    });

    // Drop anyone already mailed this period. Without this, a run that dies
    // partway (or a second firing, or a manual re-run after a fix) re-mails
    // everyone it already reached.
    const alreadySent = retentionAlreadySentThisPeriod_(start);
    const due18 = retentionFilterUnsent_('TURNING_18', turning18, alreadySent);
    const due21 = retentionFilterUnsent_('TURNING_21', turning21, alreadySent);
    const dueExp = retentionFilterUnsent_('EXPIRING', expiring, alreadySent);

    summary.dedupeAvailable = alreadySent.usable;
    summary.skipped.turning18 = turning18.length - due18.length;
    summary.skipped.turning21 = turning21.length - due21.length;
    summary.skipped.expiring = expiring.length - dueExp.length;

    Logger.info('Already-sent filter applied', {
      period: alreadySent.period,
      dedupeAvailable: alreadySent.usable,
      skipped: summary.skipped,
      toSend: { turning18: due18.length, turning21: due21.length, expiring: dueExp.length }
    });

    // Send emails for each category with progress tracking
    summary.sent.turning18 = sendTurning18Emails(due18, summary.failed.turning18);
    summary.sent.turning21 = sendTurning21Emails(due21, summary.failed.turning21);
    summary.sent.expiring = sendExpiringEmails(dueExp, summary.failed.expiring);

  } catch (err) {
    Logger.error('Retention email process failed', err);
    throw err;
  }

  summary.endTime = new Date().toISOString();
  summary.duration = new Date() - start;
  summary.totalSent = summary.sent.turning18 + summary.sent.turning21 + summary.sent.expiring;
  summary.totalFailed = summary.failed.turning18.length +
                        summary.failed.turning21.length +
                        summary.failed.expiring.length;
  summary.totalSkipped = summary.skipped.turning18 +
                         summary.skipped.turning21 +
                         summary.skipped.expiring;

  Logger.info('Retention email process completed', {
    duration: summary.duration + 'ms',
    totalSent: summary.totalSent,
    totalFailed: summary.totalFailed,
    totalSkipped: summary.totalSkipped,
    dedupeAvailable: summary.dedupeAvailable,
    breakdown: {
      turning18: { sent: summary.sent.turning18, failed: summary.failed.turning18.length, skipped: summary.skipped.turning18 },
      turning21: { sent: summary.sent.turning21, failed: summary.failed.turning21.length, skipped: summary.skipped.turning21 },
      expiring: { sent: summary.sent.expiring, failed: summary.failed.expiring.length, skipped: summary.skipped.expiring }
    }
  });
  
  // Send summary report to retention team
  sendRetentionSummaryEmail(summary);
  
  return summary;
}

// ============================================================================
// MEMBER RETRIEVAL FUNCTIONS
// ============================================================================

/**
 * Retrieves cadets turning 18 this month
 * 
 * Filters for ACTIVE CADET members whose birth month matches current month
 * and who will turn 18 this year. Requires valid PRIMARY EMAIL contact.
 * 
 * @returns {Array<Object>} Array of member objects with properties:
 *   - capid: Member's CAP ID
 *   - firstName: Member's first name
 *   - lastName: Member's last name
 *   - email: Member's primary email (sanitized)
 *   - orgid: Organization ID
 *   - rank: Member's rank
 *   - expiration: Membership expiration date
 */
function getMembersTurning18() {
  Logger.info('Retrieving members turning 18');
  
  const members = [];
  const memberData = parseFile('Member');
  const emailMap = createEmailMap();
  const currentDate = new Date();
  const currentMonth = currentDate.getMonth() + 1; // 1-12
  const currentYear = currentDate.getFullYear();
  const targetAge = RETENTION_CONFIG.AGE_THRESHOLDS.TRANSITION_TO_SENIOR;
  
  for (let i = 0; i < memberData.length; i++) {
    // Filter for active cadets with DOB
    if (memberData[i][24] !== 'ACTIVE' || 
        memberData[i][21] !== 'CADET' || 
        !memberData[i][7]) {
      continue;
    }
    
    // Parse DOB (format: M/DD/YYYY)
    const dobParts = memberData[i][7].split('/');
    if (dobParts.length !== 3) {
      Logger.warn('Invalid DOB format', {
        capsn: memberData[i][0],
        dob: memberData[i][7]
      });
      continue;
    }
    
    const birthMonth = parseInt(dobParts[0]);
    const birthYear = parseInt(dobParts[2]);
    
    // Validate parsed values
    if (isNaN(birthMonth) || isNaN(birthYear) || 
        birthMonth < 1 || birthMonth > 12) {
      Logger.warn('Invalid DOB values', {
        capsn: memberData[i][0],
        birthMonth: birthMonth,
        birthYear: birthYear
      });
      continue;
    }
    
    // Check if turning 18 this month
    if (birthMonth === currentMonth && (currentYear - birthYear) === targetAge) {
      const capid = memberData[i][0];
      const email = emailMap[capid];
      
      if (!email) {
        Logger.warn('No valid email for member turning 18', {
          capsn: capid,
          name: memberData[i][3] + ' ' + memberData[i][2]
        });
        continue;
      }
      
      members.push({
        capid: capid,
        firstName: memberData[i][3],
        lastName: memberData[i][2],
        email: email,
        orgid: memberData[i][11],
        rank: memberData[i][14],
        expiration: memberData[i][16]
      });
    }
  }
  
  Logger.info('Members turning 18 retrieved', { count: members.length });
  return members;
}

/**
 * Retrieves cadets turning 21 this month (aging out of cadet program)
 * 
 * Filters for ACTIVE CADET members whose birth month matches current month
 * and who will turn 21 this year. Requires valid PRIMARY EMAIL contact.
 * 
 * @returns {Array<Object>} Array of member objects with properties:
 *   - capid: Member's CAP ID
 *   - firstName: Member's first name
 *   - lastName: Member's last name
 *   - email: Member's primary email (sanitized)
 *   - orgid: Organization ID
 *   - rank: Member's rank
 *   - expiration: Membership expiration date
 */
function getMembersTurning21() {
  Logger.info('Retrieving members turning 21');
  
  const members = [];
  const memberData = parseFile('Member');
  const emailMap = createEmailMap();
  const currentDate = new Date();
  const currentMonth = currentDate.getMonth() + 1; // 1-12
  const currentYear = currentDate.getFullYear();
  const targetAge = RETENTION_CONFIG.AGE_THRESHOLDS.CADET_AGE_OUT;
  
  for (let i = 0; i < memberData.length; i++) {
    // Filter for active cadets with DOB
    if (memberData[i][24] !== 'ACTIVE' || 
        memberData[i][21] !== 'CADET' || 
        !memberData[i][7]) {
      continue;
    }
    
    // Parse DOB (format: M/DD/YYYY)
    const dobParts = memberData[i][7].split('/');
    if (dobParts.length !== 3) {
      Logger.warn('Invalid DOB format', {
        capsn: memberData[i][0],
        dob: memberData[i][7]
      });
      continue;
    }
    
    const birthMonth = parseInt(dobParts[0]);
    const birthYear = parseInt(dobParts[2]);
    
    // Validate parsed values
    if (isNaN(birthMonth) || isNaN(birthYear) || 
        birthMonth < 1 || birthMonth > 12) {
      Logger.warn('Invalid DOB values', {
        capsn: memberData[i][0],
        birthMonth: birthMonth,
        birthYear: birthYear
      });
      continue;
    }
    
    // Check if turning 21 this month
    if (birthMonth === currentMonth && (currentYear - birthYear) === targetAge) {
      const capid = memberData[i][0];
      const email = emailMap[capid];
      
      if (!email) {
        Logger.warn('No valid email for member turning 21', {
          capsn: capid,
          name: memberData[i][3] + ' ' + memberData[i][2]
        });
        continue;
      }
      
      members.push({
        capid: capid,
        firstName: memberData[i][3],
        lastName: memberData[i][2],
        email: email,
        orgid: memberData[i][11],
        rank: memberData[i][14],
        expiration: memberData[i][16]
      });
    }
  }
  
  Logger.info('Members turning 21 retrieved', { count: members.length });
  return members;
}

/**
 * Retrieves members expiring this month who haven't renewed
 * 
 * Filters for ACTIVE CADET and SENIOR members whose expiration date
 * falls in the current month. Requires valid PRIMARY EMAIL contact.
 * 
 * @returns {Array<Object>} Array of member objects with properties:
 *   - capid: Member's CAP ID
 *   - firstName: Member's first name
 *   - lastName: Member's last name
 *   - email: Member's primary email (sanitized)
 *   - orgid: Organization ID
 *   - rank: Member's rank
 *   - expiration: Expiration date (MM/DD/YYYY)
 *   - type: Member type (CADET or SENIOR)
 */
function getExpiringMembers() {
  Logger.info('Retrieving expiring members');
  
  const members = [];
  const memberData = parseFile('Member');
  const emailMap = createEmailMap();
  const currentDate = new Date();
  const currentMonth = currentDate.getMonth() + 1; // 1-12
  const currentYear = currentDate.getFullYear();
  
  for (let i = 0; i < memberData.length; i++) {
    // Filter for active cadets/seniors with expiration date
    if (memberData[i][24] !== 'ACTIVE' || 
        (memberData[i][21] !== 'CADET' && memberData[i][21] !== 'SENIOR') ||
        !memberData[i][16]) {
      continue;
    }
    
    // Parse expiration date (format: MM/DD/YYYY)
    const expParts = memberData[i][16].split('/');
    if (expParts.length !== 3) {
      Logger.warn('Invalid expiration date format', {
        capsn: memberData[i][0],
        expiration: memberData[i][16]
      });
      continue;
    }
    
    const expMonth = parseInt(expParts[0]);
    const expYear = parseInt(expParts[2]);
    
    // Validate parsed values
    if (isNaN(expMonth) || isNaN(expYear) || 
        expMonth < 1 || expMonth > 12) {
      Logger.warn('Invalid expiration date values', {
        capsn: memberData[i][0],
        expMonth: expMonth,
        expYear: expYear
      });
      continue;
    }
    
    // Check if expiring this month
    if (expMonth === currentMonth && expYear === currentYear) {
      const capid = memberData[i][0];
      const email = emailMap[capid];
      
      if (!email) {
        Logger.warn('No valid email for expiring member', {
          capsn: capid,
          name: memberData[i][3] + ' ' + memberData[i][2],
          type: memberData[i][21]
        });
        continue;
      }
      
      members.push({
        capid: capid,
        firstName: memberData[i][3],
        lastName: memberData[i][2],
        email: email,
        orgid: memberData[i][11],
        rank: memberData[i][14],
        expiration: memberData[i][16],
        type: memberData[i][21]
      });
    }
  }
  
  Logger.info('Expiring members retrieved', { count: members.length });
  return members;
}

/**
 * Creates email lookup map from CAPWATCH contact data
 * 
 * Extracts PRIMARY EMAIL contacts and sanitizes email addresses.
 * Used by member retrieval functions to map CAPID to email.
 * 
 * @returns {Object} Map of CAPID to sanitized primary email address
 */
function createEmailMap() {
  const contactData = parseFile('MbrContact');
  const emailMap = {};
  
  for (let i = 0; i < contactData.length; i++) {
    if (contactData[i][1] === 'EMAIL' && contactData[i][2] === 'PRIMARY') {
      const sanitized = sanitizeEmail(contactData[i][3]);
      if (sanitized) {
        emailMap[contactData[i][0]] = sanitized;
      }
    }
  }
  
  Logger.info('Email map created', { 
    totalEmails: Object.keys(emailMap).length 
  });
  return emailMap;
}

/** ORGID -> commander record, built once per execution. See below. */
let _commanderIndex = null;

/**
 * Builds ORGID -> commander record for every unit, resolving each commander's
 * CC address.
 *
 * ADDRESS ORDER — org account first, CAPWATCH last:
 *   1. their real Workspace account, read from this tenant's directory
 *   2. the derived CAP account first.last@<command domain>
 *   3. their CAPWATCH PRIMARY
 *
 * This is `rcResolveRecipientEmail_()` from notifications/RecoveryEmailNotify.gs,
 * reused rather than reimplemented so the two modules cannot drift on what
 * "reach the commander" means. (The dependency already runs the other way too —
 * that module calls createEmailMap() from this one.) Step 1 covers the cases
 * derivation cannot see: a `.2` duplicate, a manual creation, a rename. Step 2
 * covers a commander whose account exists on a domain this tenant cannot read —
 * on the CADETS tenant command staff are seniors, so COMMAND_EMAIL_DOMAIN points
 * at the senior domain, the directory read yields nothing usable, and the
 * derived senior address is the right answer. CAPWATCH primary is last because
 * it is a personal address: it reaches the person, but not at the CAP account
 * this mail belongs in.
 *
 * Derived addresses are never verified, so a commander whose account does not
 * follow the default naming will have their CC bounce. That costs the commander
 * their copy; the member's own send is unaffected, since Gmail accepts the
 * message and bounces per-recipient.
 *
 * Built once and cached because the send loop asks per member: the previous
 * per-call implementation re-walked Commanders.txt and rebuilt the whole
 * CAPWATCH email map for every single cadet email sent.
 *
 * @returns {Object} Map of ORGID to commander record
 */
function retentionCommanderIndex_() {
  if (_commanderIndex) return _commanderIndex;

  const emailMap = createEmailMap();

  // One directory listing per run. rcBuildCommandDirectoryMap_ decides whether
  // this tenant can actually see command staff and returns {} when it cannot,
  // so the guard stays in one place rather than being restated here.
  let directoryMap = {};
  try {
    directoryMap = rcBuildCommandDirectoryMap_(getActiveUsers());
  } catch (e) {
    Logger.warn('Directory unreadable — falling back to derived/CAPWATCH commander addresses', {
      errorMessage: e.message
    });
  }

  // Commanders.txt: ORGID=0, CAPID=4, NameLast=8, NameFirst=9, Rank=12.
  // Nationwide, so this is only ever read for ORGIDs of this tenant's members.
  const index = {};
  parseFile('Commanders').forEach(function (row) {
    const orgid = String(row[0] || '').trim();
    if (!orgid || index[orgid]) return; // first row per org wins, as before

    const capid = String(row[4] || '').trim();
    const info = {
      firstName: String(row[9] || '').trim(),
      lastName: String(row[8] || '').trim(),
      rank: String(row[12] || '').trim()
    };

    index[orgid] = {
      capid: capid,
      firstName: info.firstName,
      lastName: info.lastName,
      rank: info.rank,
      email: rcResolveRecipientEmail_(info, capid, emailMap, directoryMap)
    };
  });

  Logger.info('Commander index built', {
    organizations: Object.keys(index).length,
    fromDirectory: Object.keys(directoryMap).length
  });

  _commanderIndex = index;
  return index;
}

/**
 * Retrieves commander information for a given organization
 *
 * Looks up the unit commander for the specified organization ID and returns
 * their details, including the address to CC — see retentionCommanderIndex_()
 * for how that address is chosen.
 *
 * @param {string} orgid - Organization ID to look up commander for
 * @returns {Object|null} Commander object with properties or null if not found:
 *   - capid: Commander's CAP ID
 *   - firstName: Commander's first name
 *   - lastName: Commander's last name
 *   - rank: Commander's rank
 *   - email: Commander's CAP account, or CAPWATCH primary, or null
 */
function getCommanderInfo(orgid) {
  const commander = retentionCommanderIndex_()[String(orgid || '').trim()];

  if (!commander) {
    Logger.warn('Commander not found for organization', { orgid: orgid });
    return null;
  }

  return commander;
}

// ============================================================================
// TEMPLATE RENDERING
// ============================================================================

/**
 * Builds the closing signature block for member-facing retention email.
 *
 * The role holder is an individual, so their name is not version-controlled —
 * it comes from the TENANT_DIRECTOR_RECRUITING_NAME Script Property. Blank is a
 * valid state (a tenant that has not named one, or does not want a personal name
 * on automated mail) and signs with the office title alone.
 *
 * @returns {string} HTML fragment for the {{signature}} placeholder
 */
function retentionSignatureHtml_() {
  const wingLine = CONFIG.WING_NAME + ' Civil Air Patrol';
  const name = String(DIRECTOR_RECRUITING_NAME || '').trim();

  return name
    ? '<strong>' + name + '</strong><br>Director of Recruiting<br>' + wingLine
    : '<strong>Director of Recruiting</strong><br>' + wingLine;
}

/**
 * Renders a retention email template with member and wing values.
 *
 * Every member-facing template shares the wing labels and signature, so all
 * substitution happens here rather than being repeated at each send site.
 *
 * NOTE the template name must be the full slash-prefixed Apps Script filename
 * minus the folder, which this adds — a file at
 * src/recruiting-and-retention/ExpiringEmail.html deploys as
 * 'recruiting-and-retention/ExpiringEmail' and HtmlService needs that exact
 * name. Pass 'ExpiringEmail'.
 *
 * Substitution uses a replacer function, not a string, so a value containing
 * '$&' or "$'" cannot corrupt the output. Unrecognized placeholders are left
 * in place so a typo shows up in the rendered mail instead of silently
 * blanking.
 *
 * @param {string} templateName - Template filename without folder or extension
 * @param {Object} member - Member object (rank, lastName, expiration)
 * @returns {string} Rendered HTML
 */
function retentionRenderTemplate_(templateName, member) {
  const m = member || {};
  const fields = {
    rank: m.rank || '',
    lastName: m.lastName || '',
    expiration: m.expiration || '',
    wingName: CONFIG.WING_NAME,
    orgLabel: CONFIG.ORG_LABEL,
    signature: retentionSignatureHtml_()
  };

  return HtmlService
    .createHtmlOutputFromFile('recruiting-and-retention/' + templateName)
    .getContent()
    .replace(/{{(\w+)}}/g, function (match, key) {
      return Object.prototype.hasOwnProperty.call(fields, key) ? fields[key] : match;
    });
}

// ============================================================================
// EMAIL SENDING FUNCTIONS
// ============================================================================

/**
 * Sends turning 18 emails to all members in list
 * 
 * @param {Array<Object>} members - Array of member objects
 * @param {Array<Object>} failedList - Array to track failed sends
 * @returns {number} Count of successfully sent emails
 */
function sendTurning18Emails(members, failedList) {
  Logger.info('Starting turning 18 email batch', { count: members.length });
  
  let sent = 0;
  for (let i = 0; i < members.length; i++) {
    const success = sendTurning18Email(members[i]);
    
    if (success) {
      sent++;
    } else {
      failedList.push(members[i]);
    }
    
    // Progress logging
    if ((i + 1) % RETENTION_CONFIG.PROGRESS_LOG_INTERVAL === 0) {
      Logger.info('Turning 18 email progress', {
        sent: i + 1,
        total: members.length,
        percentComplete: Math.round(((i + 1) / members.length) * 100)
      });
    }
    
    // Rate limiting
    if (i < members.length - 1) {
      Utilities.sleep(RETENTION_CONFIG.EMAIL_DELAY_MS);
    }
  }
  
  Logger.info('Turning 18 email batch completed', {
    sent: sent,
    failed: failedList.length,
    total: members.length
  });
  
  return sent;
}

/**
 * Sends turning 21 emails to all members in list
 * 
 * @param {Array<Object>} members - Array of member objects
 * @param {Array<Object>} failedList - Array to track failed sends
 * @returns {number} Count of successfully sent emails
 */
function sendTurning21Emails(members, failedList) {
  Logger.info('Starting turning 21 email batch', { count: members.length });
  
  let sent = 0;
  for (let i = 0; i < members.length; i++) {
    const success = sendTurning21Email(members[i]);
    
    if (success) {
      sent++;
    } else {
      failedList.push(members[i]);
    }
    
    // Progress logging
    if ((i + 1) % RETENTION_CONFIG.PROGRESS_LOG_INTERVAL === 0) {
      Logger.info('Turning 21 email progress', {
        sent: i + 1,
        total: members.length,
        percentComplete: Math.round(((i + 1) / members.length) * 100)
      });
    }
    
    // Rate limiting
    if (i < members.length - 1) {
      Utilities.sleep(RETENTION_CONFIG.EMAIL_DELAY_MS);
    }
  }
  
  Logger.info('Turning 21 email batch completed', {
    sent: sent,
    failed: failedList.length,
    total: members.length
  });
  
  return sent;
}

/**
 * Sends expiring membership emails to all members in list
 * 
 * @param {Array<Object>} members - Array of member objects
 * @param {Array<Object>} failedList - Array to track failed sends
 * @returns {number} Count of successfully sent emails
 */
function sendExpiringEmails(members, failedList) {
  Logger.info('Starting expiring member email batch', { count: members.length });
  
  let sent = 0;
  for (let i = 0; i < members.length; i++) {
    const success = sendExpiringEmail(members[i]);
    
    if (success) {
      sent++;
    } else {
      failedList.push(members[i]);
    }
    
    // Progress logging
    if ((i + 1) % RETENTION_CONFIG.PROGRESS_LOG_INTERVAL === 0) {
      Logger.info('Expiring email progress', {
        sent: i + 1,
        total: members.length,
        percentComplete: Math.round(((i + 1) / members.length) * 100)
      });
    }
    
    // Rate limiting
    if (i < members.length - 1) {
      Utilities.sleep(RETENTION_CONFIG.EMAIL_DELAY_MS);
    }
  }
  
  Logger.info('Expiring email batch completed', {
    sent: sent,
    failed: failedList.length,
    total: members.length
  });
  
  return sent;
}

/**
 * Sends email to cadet turning 18
 * 
 * Email highlights transition opportunities for new senior members.
 * CC'd to squadron commander for awareness and follow-up.
 * 
 * @param {Object} member - Member object with email, rank, lastName, orgid
 * @returns {boolean} True if email sent successfully, false otherwise
 */
function sendTurning18Email(member) {
  try {
    const commander = getCommanderInfo(member.orgid);
    const htmlBody = retentionRenderTemplate_('Turning18Email', member);

    executeWithRetry(() =>
      GmailApp.sendEmail(
        member.email,
        RETENTION_CONFIG.SUBJECTS.TURNING_18,
        htmlBody,
        {
          htmlBody: htmlBody,
          cc: commander && commander.email ? commander.email : '',
          replyTo: DIRECTOR_RECRUITING_EMAIL,
          from: AUTOMATION_SENDER_EMAIL,
          name: SENDER_NAME
        }
      )
    );
    
    // Log successful send
    logEmailSent('TURNING_18', member, commander);
    
    Logger.info('Turning 18 email sent', {
      email: member.email,
      capsn: member.capid,
      commanderCc: commander && commander.email ? commander.email : 'none'
    });
    
    return true;
    
  } catch (e) {
    Logger.error('Failed to send turning 18 email', {
      email: member.email,
      capsn: member.capid,
      name: member.rank + ' ' + member.firstName + ' ' + member.lastName,
      errorMessage: e.message,
      errorCode: e.details?.code
    });
    return false;
  }
}

/**
 * Sends email to cadet turning 21 (aging out)
 * 
 * Email explains cadet program age-out and transition to senior membership.
 * CC'd to squadron commander for awareness and transition support.
 * 
 * @param {Object} member - Member object with email, rank, lastName, orgid
 * @returns {boolean} True if email sent successfully, false otherwise
 */
function sendTurning21Email(member) {
  try {
    const commander = getCommanderInfo(member.orgid);
    const htmlBody = retentionRenderTemplate_('Turning21Email', member);

    executeWithRetry(() =>
      GmailApp.sendEmail(
        member.email,
        RETENTION_CONFIG.SUBJECTS.TURNING_21,
        htmlBody,
        {
          htmlBody: htmlBody,
          cc: commander && commander.email ? commander.email : '',
          replyTo: DIRECTOR_RECRUITING_EMAIL,
          from: AUTOMATION_SENDER_EMAIL,
          name: SENDER_NAME
        }
      )
    );
    
    // Log successful send
    logEmailSent('TURNING_21', member, commander);
    
    Logger.info('Turning 21 email sent', {
      email: member.email,
      capsn: member.capid,
      commanderCc: commander && commander.email ? commander.email : 'none'
    });
    
    return true;
    
  } catch (e) {
    Logger.error('Failed to send turning 21 email', {
      email: member.email,
      capsn: member.capid,
      name: member.rank + ' ' + member.firstName + ' ' + member.lastName,
      errorMessage: e.message,
      errorCode: e.details?.code
    });
    return false;
  }
}

/**
 * Sends email to expiring member
 * 
 * Email reminds member of upcoming expiration and renewal process.
 * For cadets: CC'd to squadron commander
 * For seniors: No commander CC
 * 
 * @param {Object} member - Member object with email, rank, lastName, expiration, type, orgid
 * @returns {boolean} True if email sent successfully, false otherwise
 */
function sendExpiringEmail(member) {
  try {
    let commander = null;
    const htmlBody = retentionRenderTemplate_('ExpiringEmail', member);

    const options = {
      htmlBody: htmlBody,
      replyTo: DIRECTOR_RECRUITING_EMAIL,
      from: AUTOMATION_SENDER_EMAIL,
      name: SENDER_NAME
    };
    
    // Add commander CC for cadets only
    if (member.type === 'CADET') {
      commander = getCommanderInfo(member.orgid);
      if (commander && commander.email) {
        options.cc = commander.email;
      }
    }
    
    executeWithRetry(() =>
      GmailApp.sendEmail(
        member.email,
        RETENTION_CONFIG.SUBJECTS.EXPIRING,
        htmlBody,
        options
      )
    );
    
    // Log successful send
    logEmailSent('EXPIRING', member, commander);
    
    Logger.info('Expiring email sent', {
      email: member.email,
      capsn: member.capid,
      type: member.type,
      expiration: member.expiration,
      commanderCc: commander && commander.email ? commander.email : 'none'
    });
    
    return true;
    
  } catch (e) {
    Logger.error('Failed to send expiring email', {
      email: member.email,
      capsn: member.capid,
      name: member.rank + ' ' + member.firstName + ' ' + member.lastName,
      type: member.type,
      expiration: member.expiration,
      errorMessage: e.message,
      errorCode: e.details?.code
    });
    return false;
  }
}

// ============================================================================
// ALREADY-SENT GUARD
// ============================================================================

/**
 * Period key for the already-sent guard: 'yyyy-MM'.
 *
 * Calendar month is the correct grain because it is exactly what member
 * selection keys on — turning 18/21 match birth MONTH against the current month,
 * and expiring matches expiration month and year. Two runs in the same month see
 * the same population by definition, so "already mailed this month" and "already
 * mailed for this occurrence" are the same statement.
 *
 * @param {Date} when - Run start
 * @returns {string} 'yyyy-MM'
 */
function retentionPeriodKey_(when) {
  const d = when || new Date();
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
}

/**
 * Reads the Log sheet and returns who has already been mailed this period.
 *
 * The Log sheet has always been written and never read. That is what made a
 * re-run dangerous: nothing in the module knew what the previous run had done,
 * so an execution that hit the 30-minute limit partway through the expiring
 * batch would, on the next firing, start again from the top of the list.
 *
 * FAILS OPEN, LOUDLY. If the sheet is missing or unreadable this returns
 * usable=false and an empty set, so the run proceeds exactly as it did before
 * this guard existed rather than silently mailing nobody. The summary email
 * carries the same flag, so an operator can see that the run had no duplicate
 * protection instead of assuming it did.
 *
 * Only SUCCESSFUL sends reach the Log (logEmailSent is called after the send
 * returns), so a member whose send failed is correctly retried next run. The
 * converse gap is real but narrow: if the send succeeds and the Log write then
 * fails, that member is re-mailed on a re-run.
 *
 * @param {Date} when - Run start, for the period key
 * @returns {Object} { usable: boolean, period: string, keys: Object }
 */
function retentionAlreadySentThisPeriod_(when) {
  const period = retentionPeriodKey_(when);
  const result = { usable: false, period: period, keys: {} };

  if (!RETENTION_LOG_SPREADSHEET_ID) {
    Logger.warn('No retention log configured — duplicate protection is OFF for this run', {
      period: period
    });
    return result;
  }

  try {
    const sheet = SpreadsheetApp.openById(RETENTION_LOG_SPREADSHEET_ID).getSheetByName('Log');

    // No sheet yet means nothing has ever been sent. That is a usable answer,
    // not a failure — the first run legitimately has an empty history.
    if (!sheet || sheet.getLastRow() < 2) {
      result.usable = true;
      return result;
    }

    // Columns A-C only: Timestamp, Email Type, CAPID.
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    let matched = 0;

    rows.forEach(function (row) {
      const stamp = row[0];
      if (!(stamp instanceof Date)) return; // blank or text row — ignore
      if (retentionPeriodKey_(stamp) !== period) return;

      const type = String(row[1] || '').trim();
      const capid = String(row[2] || '').trim();
      if (!type || !capid) return;

      result.keys[type + '|' + capid] = true;
      matched++;
    });

    result.usable = true;
    Logger.info('Retention log read for duplicate protection', {
      period: period,
      rowsScanned: rows.length,
      alreadySentThisPeriod: matched
    });

  } catch (e) {
    Logger.warn('Retention log unreadable — duplicate protection is OFF for this run', {
      period: period,
      errorMessage: e.message
    });
  }

  return result;
}

/**
 * Drops members already mailed for this email type this period.
 *
 * @param {string} emailType - TURNING_18 | TURNING_21 | EXPIRING
 * @param {Array<Object>} members - Candidates from the retrieval functions
 * @param {Object} alreadySent - retentionAlreadySentThisPeriod_() output
 * @returns {Array<Object>} Members still due
 */
function retentionFilterUnsent_(emailType, members, alreadySent) {
  if (!alreadySent || !alreadySent.usable) return members;

  return members.filter(function (m) {
    return !alreadySent.keys[emailType + '|' + String(m.capid || '').trim()];
  });
}

// ============================================================================
// LOGGING AND REPORTING
// ============================================================================

/**
 * Logs email sent to retention tracking spreadsheet
 * 
 * Records:
 * - Timestamp of send
 * - Email type (TURNING_18, TURNING_21, EXPIRING)
 * - Member details (CAPID, name, email)
 * - Commander details if applicable
 * 
 * Creates 'Log' sheet automatically if it doesn't exist.
 * 
 * @param {string} emailType - Type of email sent (TURNING_18, TURNING_21, EXPIRING)
 * @param {Object} member - Member object
 * @param {Object|null} commander - Commander object or null
 * @returns {void}
 */
function logEmailSent(emailType, member, commander) {
  try {
    const spreadsheet = SpreadsheetApp.openById(RETENTION_LOG_SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName('Log');
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = spreadsheet.insertSheet('Log');
      sheet.appendRow([
        'Timestamp', 
        'Email Type', 
        'CAPID', 
        'Name', 
        'Email', 
        'Commander CAPID', 
        'Commander Name', 
        'Commander Email'
      ]);
      
      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, 8);
      headerRange.setFontWeight('bold')
                 .setBackground('#4285f4')
                 .setFontColor('#ffffff');
    }
    
    const commanderName = commander ? 
      (commander.rank + ' ' + commander.firstName + ' ' + commander.lastName) : '';
    const commanderEmail = commander ? commander.email : '';
    const commanderCapid = commander ? commander.capid : '';
    
    sheet.appendRow([
      new Date(),
      emailType,
      member.capid,
      member.rank + ' ' + member.firstName + ' ' + member.lastName,
      member.email,
      commanderCapid,
      commanderName,
      commanderEmail
    ]);
    
    Logger.info('Email logged to spreadsheet', {
      emailType: emailType,
      capsn: member.capid
    });
    
  } catch (e) {
    Logger.error('Failed to log email to spreadsheet', {
      emailType: emailType,
      capsn: member.capid,
      errorMessage: e.message
    });
  }
}

/**
 * Sends summary email to retention team with processing results
 * 
 * @param {Object} summary - Summary object with sent counts and errors
 * @returns {void}
 */
function sendRetentionSummaryEmail(summary) {
  try {
    const subject = `Retention Email Summary - ${new Date().toLocaleDateString()}`;
    
    let htmlBody = `
      <html>
        <head>
          <style>
            body { font-family: Arial, sans-serif; }
            h2 { color: #1a73e8; }
            table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
            th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
            th { background-color: #1a73e8; color: white; }
            .summary { background-color: #f0f0f0; padding: 15px; margin-bottom: 20px; border-radius: 5px; }
            .success { background-color: #d4edda; padding: 10px; border-left: 4px solid #28a745; margin-bottom: 15px; }
            .warning { background-color: #fff3cd; padding: 10px; border-left: 4px solid #ffc107; margin-bottom: 15px; }
          </style>
        </head>
        <body>
          <h1>Retention Email Summary</h1>
          <div class="summary">
            <h3>Summary</h3>
            <p><strong>Run Date:</strong> ${new Date(summary.startTime).toLocaleString()}</p>
            <p><strong>Duration:</strong> ${Math.round(summary.duration / 1000)} seconds</p>
            <p><strong>Total Sent:</strong> ${summary.totalSent}</p>
            <p><strong>Total Failed:</strong> ${summary.totalFailed}</p>
            <p><strong>Already sent this month (skipped):</strong> ${summary.totalSkipped}</p>
          </div>
    `;

    // An operator reading a low "sent" count needs to know whether the run was
    // deduping against the log or flying blind, so say so rather than let the
    // number be ambiguous.
    if (summary.dedupeAvailable === false) {
      htmlBody += `
        <div class="warning">
          <h2>⚠ Duplicate protection was OFF for this run</h2>
          <p>The retention log could not be read, so this run could not tell who had
          already been mailed this month. Anyone reached by an earlier run this month
          will have received a second copy. Check the execution log for the reason
          before re-running.</p>
        </div>
      `;
    }

    // Breakdown by category
    htmlBody += `
      <h2>Breakdown by Category</h2>
      <table>
        <tr>
          <th>Category</th>
          <th>Sent</th>
          <th>Failed</th>
          <th>Skipped (already sent)</th>
        </tr>
        <tr>
          <td>Turning 18</td>
          <td>${summary.sent.turning18}</td>
          <td>${summary.failed.turning18.length}</td>
          <td>${summary.skipped.turning18}</td>
        </tr>
        <tr>
          <td>Turning 21</td>
          <td>${summary.sent.turning21}</td>
          <td>${summary.failed.turning21.length}</td>
          <td>${summary.skipped.turning21}</td>
        </tr>
        <tr>
          <td>Expiring</td>
          <td>${summary.sent.expiring}</td>
          <td>${summary.failed.expiring.length}</td>
          <td>${summary.skipped.expiring}</td>
        </tr>
      </table>
    `;
    
    // Failed sends if any
    if (summary.totalFailed > 0) {
      htmlBody += `
        <div class="warning">
          <h2>⚠ Failed Sends (${summary.totalFailed})</h2>
          <p>The following members did not receive emails. Please follow up manually.</p>
        </div>
      `;
      
      // Add each failed category
      if (summary.failed.turning18.length > 0) {
        htmlBody += '<h3>Turning 18 Failures</h3><ul>';
        summary.failed.turning18.forEach(m => {
          htmlBody += `<li>${m.rank} ${m.firstName} ${m.lastName} (${m.capid}) - ${m.email}</li>`;
        });
        htmlBody += '</ul>';
      }
      
      if (summary.failed.turning21.length > 0) {
        htmlBody += '<h3>Turning 21 Failures</h3><ul>';
        summary.failed.turning21.forEach(m => {
          htmlBody += `<li>${m.rank} ${m.firstName} ${m.lastName} (${m.capid}) - ${m.email}</li>`;
        });
        htmlBody += '</ul>';
      }
      
      if (summary.failed.expiring.length > 0) {
        htmlBody += '<h3>Expiring Failures</h3><ul>';
        summary.failed.expiring.forEach(m => {
          htmlBody += `<li>${m.rank} ${m.firstName} ${m.lastName} (${m.capid}) - ${m.email}</li>`;
        });
        htmlBody += '</ul>';
      }
    } else {
      htmlBody += `
        <div class="success">
          <h2>✓ All Emails Sent Successfully</h2>
          <p>All retention emails were delivered without errors.</p>
        </div>
      `;
    }
    
    htmlBody += `
          <hr>
          <p style="font-size: 12px; color: #666;">
            This is an automated report from the ${CONFIG.ORG_LABEL} CAPWATCH Retention Email system.
            For questions or issues, please contact ${ITSUPPORT_EMAIL}.
          </p>
        </body>
      </html>
    `;
    
    // Send to retention team
    GmailApp.sendEmail(
      RETENTION_EMAIL,
      subject,
      'See HTML version',
      {
        htmlBody: htmlBody,
        from: AUTOMATION_SENDER_EMAIL,
        name: SENDER_NAME
      }
    );
    
    Logger.info('Summary email sent', {
      recipient: RETENTION_EMAIL
    });
    
  } catch (e) {
    Logger.error('Failed to send summary email', {
      errorMessage: e.message
    });
  }
}

// ============================================================================
// TEST FUNCTIONS
// ============================================================================

/**
 * Test function to preview retention email system without sending emails
 * 
 * Displays:
 * - Count of members in each category
 * - Sample member data
 * - Sample commander data
 * 
 * @returns {void}
 */
function testRetentionEmail() {
  Logger.info('Starting retention email system test');
  
  // Test getting members
  const turning18 = getMembersTurning18();
  const turning21 = getMembersTurning21();
  const expiring = getExpiringMembers();
  
  console.log('\n=== RETENTION EMAIL SYSTEM TEST ===\n');

  // Show what a real run would actually send, not just who matches. Without the
  // already-sent filter these counts overstate the send by everyone the current
  // month has already reached.
  const alreadySent = retentionAlreadySentThisPeriod_(new Date());
  console.log('Tenant profile: ' + TENANT_PROFILE +
    '   retention enabled: ' + (PROFILE_.RUN_RETENTION_EMAILS ? 'YES' : 'NO — sendRetentionEmails() is a no-op here'));
  console.log('Period: ' + alreadySent.period +
    '   duplicate protection: ' + (alreadySent.usable ? 'on' : 'OFF (log unreadable)'));
  console.log('Would send now: ' +
    'turning18=' + retentionFilterUnsent_('TURNING_18', turning18, alreadySent).length + ', ' +
    'turning21=' + retentionFilterUnsent_('TURNING_21', turning21, alreadySent).length + ', ' +
    'expiring=' + retentionFilterUnsent_('EXPIRING', expiring, alreadySent).length + '\n');

  console.log('Members turning 18: ' + turning18.length);
  if (turning18.length > 0) {
    console.log('Sample:', JSON.stringify(turning18[0], null, 2));
    const commander = getCommanderInfo(turning18[0].orgid);
    console.log('Commander:', JSON.stringify(commander, null, 2));
  }
  
  console.log('\nMembers turning 21: ' + turning21.length);
  if (turning21.length > 0) {
    console.log('Sample:', JSON.stringify(turning21[0], null, 2));
  }
  
  console.log('\nExpiring members: ' + expiring.length);
  if (expiring.length > 0) {
    console.log('Sample:', JSON.stringify(expiring[0], null, 2));
  }
  
  console.log('\n=== TEST COMPLETE ===\n');
  
  Logger.info('Test completed', {
    turning18: turning18.length,
    turning21: turning21.length,
    expiring: expiring.length
  });
}

/**
 * Test function to send a single test email to TEST_EMAIL
 * 
 * Sends a Turning 18 email with sample data to verify:
 * - Email template rendering
 * - Variable substitution
 * - Email delivery settings
 * 
 * @returns {void}
 */
function testSendSingleEmail() {
  Logger.info('Sending single test email', { recipient: TEST_EMAIL });
  
  const htmlBody = retentionRenderTemplate_('Turning18Email', {
    rank: 'C/Amn',
    lastName: 'Test Member'
  });

  GmailApp.sendEmail(
    TEST_EMAIL,
    'TEST - Turning 18 Email Preview',
    htmlBody,
    {
      replyTo: DIRECTOR_RECRUITING_EMAIL,
      from: AUTOMATION_SENDER_EMAIL,
      htmlBody: htmlBody,
      name: SENDER_NAME
    }
  );
  
  Logger.info('Test email sent', { recipient: TEST_EMAIL });
}

/**
 * Comprehensive test function that finds real members from each category
 * and sends test emails to TEST_EMAIL with actual member data
 * 
 * This allows full end-to-end testing of:
 * - Member retrieval
 * - Commander lookup
 * - Template rendering with real data
 * - Email delivery
 * 
 * Test emails are sent to TEST_EMAIL instead of actual member addresses.
 * 
 * @returns {void}
 */
function testAllRetentionEmails() {
  Logger.info('Starting comprehensive retention email test');
  
  console.log('\n=== COMPREHENSIVE RETENTION EMAIL TEST ===\n');
  
  // Get members for each category
  const turning18 = getMembersTurning18();
  const turning21 = getMembersTurning21();
  const expiring = getExpiringMembers();
  
  console.log('Found ' + turning18.length + ' members turning 18');
  console.log('Found ' + turning21.length + ' members turning 21');
  console.log('Found ' + expiring.length + ' members expiring\n');
  
  // Test Turning 18 Email
  if (turning18.length > 0) {
    console.log('--- TESTING TURNING 18 EMAIL ---');
    const testMember = turning18[0];
    console.log('Sample Member: ' + JSON.stringify(testMember, null, 2));
    
    const commander = getCommanderInfo(testMember.orgid);
    console.log('Commander: ' + JSON.stringify(commander, null, 2));
    
    const htmlBody = retentionRenderTemplate_('Turning18Email', testMember);

    GmailApp.sendEmail(
      TEST_EMAIL,
      'TEST - Turning 18 Email Preview - ' + testMember.rank + ' ' + testMember.lastName,
      htmlBody,
      {
        replyTo: DIRECTOR_RECRUITING_EMAIL,
        from: AUTOMATION_SENDER_EMAIL,
        htmlBody: htmlBody,
        name: SENDER_NAME
      }
    );
    
    console.log('✓ Sent Turning 18 test email to: ' + TEST_EMAIL + '\n');
  } else {
    console.log('✗ No members turning 18 found - skipping test\n');
  }
  
  // Test Turning 21 Email
  if (turning21.length > 0) {
    console.log('--- TESTING TURNING 21 EMAIL ---');
    const testMember = turning21[0];
    console.log('Sample Member: ' + JSON.stringify(testMember, null, 2));
    
    const commander = getCommanderInfo(testMember.orgid);
    console.log('Commander: ' + JSON.stringify(commander, null, 2));
    
    const htmlBody = retentionRenderTemplate_('Turning21Email', testMember);

    GmailApp.sendEmail(
      TEST_EMAIL,
      'TEST - Turning 21 Email Preview - ' + testMember.rank + ' ' + testMember.lastName,
      htmlBody,
      {
        replyTo: DIRECTOR_RECRUITING_EMAIL,
        from: AUTOMATION_SENDER_EMAIL,
        htmlBody: htmlBody,
        name: SENDER_NAME
      }
    );
    
    console.log('✓ Sent Turning 21 test email to: ' + TEST_EMAIL + '\n');
  } else {
    console.log('✗ No members turning 21 found - skipping test\n');
  }
  
  // Test Expiring Email
  if (expiring.length > 0) {
    console.log('--- TESTING EXPIRING EMAIL ---');
    const testMember = expiring[0];
    console.log('Sample Member: ' + JSON.stringify(testMember, null, 2));
    
    let commander = null;
    if (testMember.type === 'CADET') {
      commander = getCommanderInfo(testMember.orgid);
      console.log('Commander (Cadet): ' + JSON.stringify(commander, null, 2));
    } else {
      console.log('Member is SENIOR - no commander CC');
    }
    
    const htmlBody = retentionRenderTemplate_('ExpiringEmail', testMember);

    GmailApp.sendEmail(
      TEST_EMAIL,
      'TEST - Expiring Email Preview - ' + testMember.rank + ' ' + testMember.lastName,
      htmlBody,
      {
        replyTo: DIRECTOR_RECRUITING_EMAIL,
        from: AUTOMATION_SENDER_EMAIL,
        htmlBody: htmlBody,
        name: SENDER_NAME
      }
    );
    
    console.log('✓ Sent Expiring test email to: ' + TEST_EMAIL + '\n');
  } else {
    console.log('✗ No expiring members found - skipping test\n');
  }
  
  console.log('=== TEST COMPLETE ===');
  console.log('Check your inbox at: ' + TEST_EMAIL + '\n');
  
  Logger.info('Comprehensive test completed', {
    turning18Available: turning18.length > 0,
    turning21Available: turning21.length > 0,
    expiringAvailable: expiring.length > 0,
    testEmail: TEST_EMAIL
  });
}