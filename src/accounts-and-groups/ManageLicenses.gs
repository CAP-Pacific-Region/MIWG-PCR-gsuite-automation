/**
 * License Lifecycle Management Module
 *
 * Version: 1.1.0
 * Date: 2026-07-15
 * Changes: 1.1.0 — deleteIneligibleSuspendedUsers() repaired. It had never once
 *   run: its field selector asked for employeeId, which is not a Directory User
 *   field, so every call 400'd on its first page and manageLicenseLifecycle()
 *   filed the error into the monthly report. Its grace period also measured days
 *   since last login rather than days since the member lapsed, and Google returns
 *   lastLoginTime as the Unix epoch (not absent) for never-logged-in accounts, so
 *   those read as ~20649 days stale. Grace is now measured from the CAPWATCH
 *   Expiration column; current members are never auto-deleted; and the function
 *   defaults to a dry run. See PCR_CHANGELOG.md.
 *   1.0.0 — AdminDirectory.Users.list standardized to customer:"my_customer"
 *   (was domain — 400 Bad Request on the cadets tenant).
 *
 * Manages Google Workspace license optimization and user account lifecycle:
 * - Auto-reactivates users who renewed their membership
 * - Archives users suspended for 1+ year who are inactive in CAPWATCH
 * - Deletes users archived for 5+ years who are inactive in CAPWATCH
 * - Deletes suspended accounts of members who lapsed 30+ days ago
 * - Maintains license pool availability
 *
 * RECOMMENDED SCHEDULE: Run monthly around the 15th
 * This avoids conflicts with beginning-of-month member sync and provides
 * buffer after renewal processing.
 */

/**
 * Main function to manage license lifecycle
 * This should be scheduled to run monthly (recommend mid-month, around the 15th)
 * 
 * Process:
 * 1. Archives users suspended 1+ year who are inactive in CAPWATCH
 * 2. Deletes users archived 5+ years who are inactive in CAPWATCH
 * 
 * @returns {Object} Summary of actions taken
 */
function manageLicenseLifecycle() {
  const start = new Date();
  Logger.info('Starting license lifecycle management');
  
  // Clear cache to ensure fresh CAPWATCH data
  clearCache();
  
  // Get active members from CAPWATCH
  const activeMembers = getMembers();
  const activeCapsns = new Set(Object.keys(activeMembers));
  
  Logger.info('Active members loaded from CAPWATCH', { 
    count: activeCapsns.size 
  });
  
  // Initialize summary
  const summary = {
    archived: [],
    deleted: [],
    deletedIneligible: [],
    ineligibleResult: null,
    errors: [],
    startTime: start.toISOString()
  };

  try {
    // Step 1: Archive long-suspended users not active in CAPWATCH.
    // NOTE: archiving is unavailable on this Workspace for Nonprofits edition
    // (no Archived User licenses provisioned). This step is effectively a
    // no-op until/unless Archived User licenses are added to the domain.
    summary.archived = archiveLongSuspendedUsers(activeCapsns);

    // Step 2: Delete long-archived users not active in CAPWATCH.
    // Kept commented out since archiving is unavailable (step above no-ops).
    // summary.deleted = deleteLongArchivedUsers(activeCapsns);

    // Step 3: Report suspended ineligible users that are past the grace period.
    // Ineligible = suspended in Workspace AND not an eligible active CAPWATCH
    // member (lapsed, no record, etc.). Accounts become eligible again before
    // deletion if the member's CAPWATCH type or status changes —
    // reactivateRenewedMembers() handles that path and prevents deletion.
    //
    // Exception: members mid-transition to the senior tenant are skipped here.
    // CadetTransition.gs holds their mailbox open on its own 90-day clock and
    // deletes it once their mail has been migrated across.
    //
    // STILL A DRY RUN. The grace period now measures the right thing (days since
    // CAPWATCH expiry, not days since last login), so this list is a safe delete
    // set — but nobody has watched it run live yet, and deletion is permanent on
    // this edition. Flip to false to arm the monthly reaper.
    summary.ineligibleResult = deleteIneligibleSuspendedUsers(true);
    summary.deletedIneligible = summary.ineligibleResult.deleted;

  } catch (err) {
    Logger.error('License lifecycle management failed', err);
    summary.errors.push({
      message: err.message,
      timestamp: new Date().toISOString()
    });
  }
  
  summary.endTime = new Date().toISOString();
  summary.duration = new Date() - start;
  
  Logger.info('License lifecycle management completed', summary);
  
  // Send notification email with summary
  sendLicenseManagementReport(summary);
  
  return summary;
}

/**
 * Reactivates users who are suspended in Google but active in CAPWATCH
 * This handles cases where users renewed but account wasn't automatically unsuspended
 * 
 * @param {Set<string>} activeCapsns - Set of active CAPSNs from CAPWATCH
 * @returns {Array<Object>} Array of reactivated users
 */
function reactivateSuspendedActiveUsers(activeCapsns) {
  Logger.info('Starting reactivation of suspended active users');
  
  const reactivated = [];
  let nextPageToken = '';
  
  do {
    const page = AdminDirectory.Users.list({
      customer: "my_customer",
      maxResults: 500,
      query: 'isSuspended=true',
      fields: 'users(primaryEmail,name,orgUnitPath,suspended,customSchemas),nextPageToken',
      pageToken: nextPageToken
    });
    
    nextPageToken = page.nextPageToken;
    
    if (page.users) {
      for (const user of page.users) {
        // Extract CAPSN from email (format: 123456@miwg.cap.gov)
        const capsn = user.primaryEmail.split('@')[0];
        
        // Skip if not numeric (admin accounts, etc.)
        if (!/^\d+$/.test(capsn)) {
          continue;
        }
        
        // Check if user is active in CAPWATCH
        if (activeCapsns.has(capsn)) {
          try {
            // Unsuspend the user
            const unsuspendResult = executeWithRetry(() => 
              AdminDirectory.Users.update(
                { suspended: false },
                user.primaryEmail
              )
            );
            
            reactivated.push({
              email: user.primaryEmail,
              capsn: capsn,
              name: `${user.name.givenName} ${user.name.familyName}`,
              orgUnitPath: user.orgUnitPath || '/',
              reactivatedAt: new Date().toISOString()
            });
            
            Logger.info('User reactivated', {
              email: user.primaryEmail,
              capsn: capsn,
              name: `${user.name.givenName} ${user.name.familyName}`
            });
            
            // Small delay to avoid rate limits
            Utilities.sleep(100);
            
          } catch (err) {
            Logger.error('Failed to reactivate user', {
              email: user.primaryEmail,
              capsn: capsn,
              errorMessage: err.message,
              errorCode: err.details?.code
            });
          }
        }
      }
    }
  } while (nextPageToken);
  
  Logger.info('Reactivation completed', { count: reactivated.length });
  return reactivated;
}

/**
 * Archives users who have been suspended for 1+ year and are not active in CAPWATCH
 * Sets archived flag - Google handles license change automatically
 * Users remain in their current OU until deletion
 * 
 * @param {Set<string>} activeCapsns - Set of active CAPSNs from CAPWATCH
 * @returns {Array<Object>} Array of archived users
 */
function archiveLongSuspendedUsers(activeCapsns) {
  Logger.info('Starting archival of long-suspended users');
  
  const archived = [];
  const oneYearAgo = new Date();
  oneYearAgo.setDate(oneYearAgo.getDate() - LICENSE_CONFIG.DAYS_BEFORE_ARCHIVE);
  
  let nextPageToken = '';
  let processedCount = 0;
  
  do {
    const page = AdminDirectory.Users.list({
      customer: "my_customer",
      maxResults: 500,
      query: 'isSuspended=true isArchived=false',
      fields: 'users(primaryEmail,name,orgUnitPath,creationTime,lastLoginTime,customSchemas),nextPageToken',
      pageToken: nextPageToken
    });
    
    nextPageToken = page.nextPageToken;
    
    if (page.users) {
      for (const user of page.users) {
        // Safety limit
        if (processedCount >= LICENSE_CONFIG.MAX_BATCH_SIZE) {
          Logger.warn('Reached max batch size for archival', {
            maxSize: LICENSE_CONFIG.MAX_BATCH_SIZE
          });
          break;
        }
        
        // Extract CAPSN from email
        const capsn = user.primaryEmail.split('@')[0];
        
        // Skip if not numeric (admin accounts, etc.)
        if (!/^\d+$/.test(capsn)) {
          continue;
        }
        
        // Skip if active in CAPWATCH
        if (activeCapsns.has(capsn)) {
          continue;
        }
        
        // Determine suspension date (use lastLoginTime or creationTime as proxy)
        // Note: Google doesn't expose exact suspension date, so we use lastLoginTime
        const lastActivityDate = user.lastLoginTime ? 
          new Date(user.lastLoginTime) : 
          new Date(user.creationTime);
        
        // Check if suspended long enough
        if (lastActivityDate < oneYearAgo) {
          try {
            // Archive the user (Google handles license change automatically)
            const archiveResult = executeWithRetry(() => 
              AdminDirectory.Users.update(
                { archived: true },
                user.primaryEmail
              )
            );
            
            archived.push({
              email: user.primaryEmail,
              capsn: capsn,
              name: `${user.name.givenName} ${user.name.familyName}`,
              orgUnitPath: user.orgUnitPath || '/',
              lastActivity: lastActivityDate.toISOString(),
              archivedAt: new Date().toISOString(),
              daysSinceActivity: Math.floor((new Date() - lastActivityDate) / (1000 * 60 * 60 * 24))
            });
            
            Logger.info('User archived', {
              email: user.primaryEmail,
              capsn: capsn,
              name: `${user.name.givenName} ${user.name.familyName}`,
              orgUnitPath: user.orgUnitPath,
              daysSinceActivity: Math.floor((new Date() - lastActivityDate) / (1000 * 60 * 60 * 24))
            });
            
            processedCount++;
            
            // Small delay to avoid rate limits
            Utilities.sleep(100);
            
          } catch (err) {
            Logger.error('Failed to archive user', {
              email: user.primaryEmail,
              capsn: capsn,
              errorMessage: err.message,
              errorCode: err.details?.code
            });
          }
        }
      }
      
      if (processedCount >= LICENSE_CONFIG.MAX_BATCH_SIZE) {
        break;
      }
    }
  } while (nextPageToken);
  
  Logger.info('Archival completed', { count: archived.length });
  return archived;
}

/**
 * Deletes users who have been archived for 5+ years and are not active in CAPWATCH
 * This is the final stage of account lifecycle management
 * 
 * @param {Set<string>} activeCapsns - Set of active CAPSNs from CAPWATCH
 * @returns {Array<Object>} Array of deleted users
 */
function deleteLongArchivedUsers(activeCapsns) {
  Logger.info('Starting deletion of long-archived users');
  
  const deleted = [];
  const fiveYearsAgo = new Date();
  fiveYearsAgo.setDate(fiveYearsAgo.getDate() - LICENSE_CONFIG.DAYS_BEFORE_DELETE);
  
  let nextPageToken = '';
  let processedCount = 0;
  
  do {
    const page = AdminDirectory.Users.list({
      customer: "my_customer",
      maxResults: 500,
      query: 'isArchived=true',
      fields: 'users(primaryEmail,name,orgUnitPath,creationTime,lastLoginTime,customSchemas),nextPageToken',
      pageToken: nextPageToken
    });
    
    nextPageToken = page.nextPageToken;
    
    if (page.users) {
      for (const user of page.users) {
        // Safety limit
        if (processedCount >= LICENSE_CONFIG.MAX_BATCH_SIZE) {
          Logger.warn('Reached max batch size for deletion', {
            maxSize: LICENSE_CONFIG.MAX_BATCH_SIZE
          });
          break;
        }
        
        // Extract CAPSN from email
        const capsn = user.primaryEmail.split('@')[0];
        
        // Skip if not numeric (admin accounts, etc.)
        if (!/^\d+$/.test(capsn)) {
          continue;
        }
        
        // Skip if active in CAPWATCH (someone rejoined!)
        if (activeCapsns.has(capsn)) {
          Logger.warn('Archived user is now active in CAPWATCH - manual reactivation needed', {
            email: user.primaryEmail,
            capsn: capsn
          });
          continue;
        }
        
        // Determine archive date (use lastLoginTime as proxy)
        const lastActivityDate = user.lastLoginTime ? 
          new Date(user.lastLoginTime) : 
          new Date(user.creationTime);
        
        // Check if archived long enough
        if (lastActivityDate < fiveYearsAgo) {
          try {
            // Store user info before deletion
            const userName = `${user.name.givenName} ${user.name.familyName}`;
            const orgUnit = user.orgUnitPath || '/';
            
            // Delete the user
            executeWithRetry(() => 
              AdminDirectory.Users.remove(user.primaryEmail)
            );
            
            deleted.push({
              email: user.primaryEmail,
              capsn: capsn,
              name: userName,
              orgUnitPath: orgUnit,
              lastActivity: lastActivityDate.toISOString(),
              deletedAt: new Date().toISOString(),
              daysSinceActivity: Math.floor((new Date() - lastActivityDate) / (1000 * 60 * 60 * 24))
            });
            
            Logger.info('User deleted', {
              email: user.primaryEmail,
              capsn: capsn,
              name: userName,
              orgUnitPath: orgUnit,
              daysSinceActivity: Math.floor((new Date() - lastActivityDate) / (1000 * 60 * 60 * 24))
            });
            
            processedCount++;
            
            // Small delay to avoid rate limits
            Utilities.sleep(100);
            
          } catch (err) {
            Logger.error('Failed to delete user', {
              email: user.primaryEmail,
              capsn: capsn,
              errorMessage: err.message,
              errorCode: err.details?.code
            });
          }
        }
      }
      
      if (processedCount >= LICENSE_CONFIG.MAX_BATCH_SIZE) {
        break;
      }
    }
  } while (nextPageToken);
  
  Logger.info('Deletion completed', { count: deleted.length });
  return deleted;
}

/**
 * Deletes suspended Workspace accounts belonging to members who lapsed at least
 * LICENSE_CONFIG.DAYS_BEFORE_DELETE_INELIGIBLE days ago (default 30), freeing
 * seats against the 2000-user domain cap.
 *
 * Defaults to a dry run. Pass dryRun=false to actually delete. Deletion is
 * PERMANENT here: this edition has no Archived User licenses, so there is no
 * archive and no undo beyond Google's ~20-day recovery window.
 *
 * THE GRACE PERIOD IS MEASURED FROM CAPWATCH EXPIRATION, not from lastLoginTime.
 * The two are unrelated: someone who stopped reading their mail a year ago but
 * renewed last month is a current member, while lastLoginTime says they are the
 * stalest account on the domain. Workspace exposes no suspension timestamp, and
 * suspension follows expiry on the next sync, so the Expiration column is the
 * closest authoritative date to the event this grace period is actually about.
 *
 * lastLoginTime never dates a lapse and so never drives the grace period — but
 * it can DISPROVE one, and is used for exactly that in the no-record branch
 * below. It is otherwise reported as context for a human reader.
 *
 * Decision order per suspended account:
 *
 *   Manual Member          -> never deleted (allow-list)
 *   held for transition    -> never deleted (CadetTransition.gs owns the clock)
 *   CADET                  -> skipped (separate domain, out of scope)
 *   MbrStatus ACTIVE       -> never auto-deleted, even when ineligible by type.
 *                             A PATRON is a current member; their expiry is in
 *                             the future and cannot date their conversion, so
 *                             there is nothing safe to measure. Surfaced for a
 *                             human instead.
 *   lapsed, expiry known   -> deleted once >= grace days past Expiration
 *   lapsed, expiry absent  -> held and surfaced; cannot measure, so no deletion
 *   no CAPWATCH record     -> deleted, but only if the claim survives a check.
 *                             Member.txt retains a rolling window of lapsed
 *                             members (~3 months observed), so falling out of it
 *                             means lapsing before the oldest lapse it still
 *                             carries — far beyond any 30-day grace. Two things
 *                             can falsify that:
 *                               - a short extract, which would make current
 *                                 members look absent (MIN_MEMBER_ROWS guard);
 *                               - the account being alive after that boundary,
 *                                 which is impossible for someone who lapsed
 *                                 before it, since lapsing gets you suspended and
 *                                 a suspended account cannot sign in. A wing
 *                                 transfer looks exactly like this. Held for
 *                                 review rather than deleted.
 *
 * @param {boolean} dryRun - When true (default), no accounts are deleted.
 * @returns {Object} { dryRun, graceDays, memberRows, oldestRetainedLapse,
 *   deleted, candidates, withinGrace, activeIneligible, unknownExpiry,
 *   contradicted }
 */
function deleteIneligibleSuspendedUsers(dryRun = true) {
  Logger.info('Starting deleteIneligibleSuspendedUsers', { dryRun: dryRun });

  const graceDays = LICENSE_CONFIG.DAYS_BEFORE_DELETE_INELIGIBLE;

  // Build eligibility lookup: CAPID -> {status, type, expiration}
  // [16] Expiration, [21] Type, [24] MbrStatus — verified against the Member.txt
  // header by debugCapwatchMemberExpirationColumn().
  const memberData = parseFile('Member');

  // Oldest lapse still present in the extract, derived rather than assumed: the
  // window has been ~3 months, but nothing guarantees CAPWATCH keeps it there.
  // This is the boundary that gives "no CAPWATCH record" its meaning — you only
  // fall out of Member.txt by lapsing before it.
  let oldestRetainedLapse = null;

  // A short Member.txt would make current members look like they have no
  // CAPWATCH record and get them deleted. Refuse rather than guess.
  if (!memberData || memberData.length < LICENSE_CONFIG.MIN_MEMBER_ROWS) {
    throw new Error(
      `deleteIneligibleSuspendedUsers: CAPWATCH Member data looks truncated ` +
      `(${memberData ? memberData.length : 0} rows, expected at least ` +
      `${LICENSE_CONFIG.MIN_MEMBER_ROWS}). Refusing to run: this function treats ` +
      `a missing CAPWATCH record as proof a member lapsed long ago, which would ` +
      `delete current members' accounts if the extract is incomplete.`
    );
  }
  const memberInfoByCapid = {};
  for (let i = 0; i < memberData.length; i++) {
    const capid = String(memberData[i][0] || '').trim();
    if (!capid) continue;
    memberInfoByCapid[capid] = {
      status: memberData[i][24],
      type: memberData[i][21],
      expiration: memberData[i][16]
    };

    if (String(memberData[i][24] || '').trim().toUpperCase() !== 'ACTIVE') {
      const lapsed = parseCapwatchDate_(memberData[i][16]);
      if (lapsed && (!oldestRetainedLapse || lapsed < oldestRetainedLapse)) {
        oldestRetainedLapse = lapsed;
      }
    }
  }
  Logger.info('CAPWATCH extract window', {
    memberRows: memberData.length,
    oldestRetainedLapse: oldestRetainedLapse ?
      oldestRetainedLapse.toISOString().slice(0, 10) : null
  });

  // Manual Members allow-list — never delete these
  const manualCapids = {};
  try {
    const ss = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Manual Members') || ss.getSheetByName('ManualMembers');
    if (sheet) {
      const rows = sheet.getDataRange().getValues();
      if (rows && rows.length > 1) {
        const header = rows[0].map(h => String(h || '').trim());
        const capidIdx = header.indexOf('CAPID');
        if (capidIdx > -1) {
          for (let r = 1; r < rows.length; r++) {
            const capid = String(rows[r][capidIdx] || '').trim();
            if (capid) manualCapids[capid] = 1;
          }
        }
      }
    }
  } catch (e) {
    Logger.warn('Unable to read Manual Members sheet for deletion allow-list', {
      errorMessage: e && e.message ? e.message : String(e)
    });
  }

  // CAPIDs mid-transition to the senior tenant. Their cadet mailbox is still the
  // only copy of their mail until it is migrated. A transitioning cadet's old
  // type has lapsed, so the expiry clock below would happily reap them the first
  // run after their type flips. CadetTransition.gs deletes them instead, on its
  // own authoritative clock, once the mail is safely across.
  const heldForTransition = getHeldTransitionCapids();

  const eligibleTypes = CONFIG.MEMBER_TYPES.ACTIVE;
  const candidates = [];        // past grace — will be deleted
  const withinGrace = [];       // lapsed recently — spared, will age in
  const activeIneligible = [];  // current members, ineligible by type — never auto-deleted
  const unknownExpiry = [];     // lapsed but undateable — held for review
  const contradicted = [];      // no CAPWATCH record, but demonstrably alive since — held
  let nextPageToken = '';

  do {
    const page = AdminDirectory.Users.list({
      customer: "my_customer",
      maxResults: 500,
      query: 'isSuspended=true isAdmin=false',
      fields: 'users(primaryEmail,name,orgUnitPath,suspended,externalIds,creationTime,lastLoginTime),nextPageToken',
      pageToken: nextPageToken
    });

    nextPageToken = page.nextPageToken;
    if (!page.users) continue;

    for (const user of page.users) {
      const ids = user.externalIds || [];
      const capidExt = ids.find(id => id.type === 'organization');
      const capid = capidExt ? String(capidExt.value || '').trim() : '';

      if (!capid) continue;
      if (manualCapids[capid]) continue;
      if (heldForTransition[capid]) continue;

      const info = memberInfoByCapid[capid];
      const type = info ? String(info.type || '').trim().toUpperCase() : '';
      const status = info ? String(info.status || '').trim().toUpperCase() : '';

      // Cadets live on a separate domain — out of scope here entirely.
      if (type === 'CADET') continue;

      // lastLoginTime is reported for human context only; nothing below keys off
      // it. A user who has never signed in comes back with lastLoginTime set to
      // the Unix epoch rather than with the field absent, so test the value, not
      // its presence — otherwise every such account reads as ~20649 days stale.
      const lastLogin = user.lastLoginTime ? new Date(user.lastLoginTime) : null;
      const neverLoggedIn = !lastLogin || lastLogin.getTime() <= 0;
      const lastActivity = neverLoggedIn ? new Date(user.creationTime) : lastLogin;

      const name = user.name && user.name.fullName ? user.name.fullName :
        `${user.name.givenName} ${user.name.familyName}`;
      const expiration = info ? parseCapwatchDate_(info.expiration) : null;
      const daysSinceExpiration = expiration ?
        Math.floor((new Date() - expiration) / (1000 * 60 * 60 * 24)) : null;

      const record = {
        email: user.primaryEmail,
        capsn: capid,
        name: name,
        orgUnitPath: user.orgUnitPath || '/',
        neverLoggedIn: neverLoggedIn,
        lastActivity: lastActivity.toISOString(),
        daysSinceActivity: Math.floor((new Date() - lastActivity) / (1000 * 60 * 60 * 24)),
        expiration: expiration ? expiration.toISOString().slice(0, 10) : null,
        daysSinceExpiration: daysSinceExpiration,
        reason: info ? `${type || '(blank)'} / ${status || '(blank)'}` : 'no CAPWATCH record',
        deleted: false,
        dryRun: dryRun
      };

      // A current member is never auto-deleted, whatever their type. Eligible
      // ones shouldn't be suspended at all (reactivateRenewedMembers() unsuspends
      // them); ineligible-by-type ones are real members whose expiry lies in the
      // future and so cannot date the conversion that made them ineligible.
      // There is nothing safe to measure, so a human decides.
      if (status === 'ACTIVE') {
        if (eligibleTypes.indexOf(type) === -1) activeIneligible.push(record);
        continue;
      }

      if (info) {
        // Lapsed member with an unreadable Expiration — cannot measure the
        // grace period, so hold rather than guess.
        if (!expiration) {
          unknownExpiry.push(record);
          continue;
        }
        if (daysSinceExpiration < graceDays) {
          withinGrace.push(record);
          continue;
        }
        record.basis = `lapsed ${daysSinceExpiration}d ago (expired ${record.expiration})`;
      } else {
        // No CAPWATCH record. The claim is: they lapsed before
        // oldestRetainedLapse, which is the only way to fall out of Member.txt.
        //
        // That claim is checkable. Lapsing gets you suspended on the next sync,
        // and a suspended account cannot sign in — so if this account was still
        // alive after that boundary, it did not lapse before it, and something
        // else removed them from our extract (a wing transfer leaves CAPWATCH
        // silent while the member stays perfectly current elsewhere). Deleting on
        // a falsified premise is how you destroy a live mailbox, so hand it to a
        // human. lastActivity is the account's newest sign of life: last login,
        // or creation for an account never signed into.
        if (oldestRetainedLapse && lastActivity > oldestRetainedLapse) {
          record.reason = 'no CAPWATCH record, but alive after the extract window';
          contradicted.push(record);
          continue;
        }
        record.basis = 'no CAPWATCH record — lapsed beyond the extract window';
      }

      candidates.push(record);
    }
  } while (nextPageToken);

  // Longest-lapsed first; no-record accounts (undateable, but lapsed beyond the
  // extract window) sort ahead of everything with a date.
  candidates.sort((a, b) => {
    if (a.daysSinceExpiration === null && b.daysSinceExpiration === null) return 0;
    if (a.daysSinceExpiration === null) return -1;
    if (b.daysSinceExpiration === null) return 1;
    return b.daysSinceExpiration - a.daysSinceExpiration;
  });
  withinGrace.sort((a, b) => b.daysSinceExpiration - a.daysSinceExpiration);

  console.log(`\n=== ${dryRun ? 'DRY RUN - ' : ''}INELIGIBLE SUSPENDED DELETIONS ===\n`);
  console.log(`Grace period: ${graceDays} days, measured from CAPWATCH expiry.`);
  console.log(`Extract retains lapses back to ${oldestRetainedLapse ? oldestRetainedLapse.toISOString().slice(0, 10) : '(unknown)'}.`);
  console.log('LAPSED = days since membership expired; "no record" = absent from the extract,');
  console.log('         i.e. lapsed before that date. DAYS = days since last login (* = never');
  console.log('         logged in); it never sets the grace period, but a no-record account');
  console.log('         alive after the window is held rather than deleted.\n');

  console.log(`-- ${dryRun ? 'WOULD DELETE' : 'DELETING'} (${candidates.length}) --`);
  if (candidates.length === 0) console.log('    (none)');
  candidates.forEach(c => console.log(
    `${c.name.padEnd(34)} ${c.capsn.padEnd(8)} ${(c.daysSinceExpiration === null ? 'no record' : String(c.daysSinceExpiration) + 'd').padEnd(10)} ${(String(c.daysSinceActivity) + (c.neverLoggedIn ? '*' : '')).padEnd(7)} ${c.email}`
  ));

  console.log(`\n-- SPARED, still within grace (${withinGrace.length}) --`);
  withinGrace.forEach(c => console.log(
    `${c.name.padEnd(34)} ${c.capsn.padEnd(8)} lapsed ${c.daysSinceExpiration}d ago, deletable in ${graceDays - c.daysSinceExpiration}d`
  ));

  console.log(`\n-- SPARED, current member ineligible by type — needs a human (${activeIneligible.length}) --`);
  activeIneligible.forEach(c => console.log(
    `${c.name.padEnd(34)} ${c.capsn.padEnd(8)} ${c.reason.padEnd(24)} ${c.email}`
  ));

  if (contradicted.length > 0) {
    console.log(`\n-- SPARED, no CAPWATCH record but alive since the extract window — needs a human (${contradicted.length}) --`);
    console.log(`   (extract retains lapses back to ${oldestRetainedLapse ? oldestRetainedLapse.toISOString().slice(0, 10) : '?'}; these signed in after that,`);
    console.log('    so they did not lapse — likely transferred out of the wing.)');
    contradicted.forEach(c => console.log(
      `${c.name.padEnd(34)} ${c.capsn.padEnd(8)} last active ${String(c.daysSinceActivity) + 'd ago'} ${c.email}`
    ));
  }

  if (unknownExpiry.length > 0) {
    console.log(`\n-- SPARED, lapsed but expiration unreadable — needs a human (${unknownExpiry.length}) --`);
    unknownExpiry.forEach(c => console.log(
      `${c.name.padEnd(34)} ${c.capsn.padEnd(8)} ${c.reason.padEnd(24)} ${c.email}`
    ));
  }

  const summary = {
    dryRun: dryRun,
    graceDays: graceDays,
    memberRows: memberData.length,
    oldestRetainedLapse: oldestRetainedLapse ?
      oldestRetainedLapse.toISOString().slice(0, 10) : null,
    candidates: candidates,
    withinGrace: withinGrace,
    activeIneligible: activeIneligible,
    unknownExpiry: unknownExpiry,
    contradicted: contradicted,
    deleted: []
  };

  if (dryRun) {
    console.log(`\nDry run — nothing deleted. Call deleteIneligibleSuspendedUsers(false) to delete the ${candidates.length} above.\n`);
    Logger.info('deleteIneligibleSuspendedUsers dry run completed', {
      wouldDelete: candidates.length,
      withinGrace: withinGrace.length,
      activeIneligible: activeIneligible.length,
      unknownExpiry: unknownExpiry.length,
      contradicted: contradicted.length,
      graceDays: graceDays
    });
    return summary;
  }

  candidates.forEach(c => {
    try {
      executeWithRetry(() => AdminDirectory.Users.remove(c.email));
      c.deleted = true;
      c.deletedAt = new Date().toISOString();
      summary.deleted.push(c);
      Logger.info('Ineligible suspended user deleted', {
        email: c.email, capsn: c.capsn, basis: c.basis
      });
    } catch (e) {
      c.errorMessage = e.message;
      Logger.error('Failed to delete ineligible suspended user', {
        email: c.email,
        capsn: c.capsn,
        errorMessage: e.message,
        errorCode: e.details?.code
      });
    }
    Utilities.sleep(100);
  });

  console.log(`\nDeleted ${summary.deleted.length} of ${candidates.length}.\n`);
  Logger.info('deleteIneligibleSuspendedUsers completed', {
    deleted: summary.deleted.length,
    failed: candidates.length - summary.deleted.length,
    withinGrace: withinGrace.length,
    activeIneligible: activeIneligible.length,
    graceDays: graceDays
  });

  return summary;
}

/**
 * Read-only diagnostic for the CAPWATCH Member.txt expiration column.
 *
 * deleteIneligibleSuspendedUsers() reads column [16] as the membership
 * expiration date, a positional index inherited from getExpiringMembers().
 * parseFile() strips the header, so nothing in the codebase actually verifies
 * that index. Before any deletion policy keys off it, this prints:
 *
 *   1. The real header, so [16] can be confirmed by name.
 *   2. The distribution of expiration values by member status, to test whether
 *      the extract only retains recently-expired members.
 *   3. Raw rows for specific CAPIDs, to confirm column alignment end to end.
 *
 * Makes no changes and touches no accounts.
 *
 * @param {Array<string>} sampleCapids - CAPIDs to dump raw rows for
 * @returns {void}
 */
function debugCapwatchMemberExpirationColumn(sampleCapids) {
  const samples = sampleCapids || ['618863', '762848', '632609', '582442'];

  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files = folder.getFilesByName('Member.txt');
  if (!files.hasNext()) {
    console.log('Member.txt not found in the CAPWATCH data folder.');
    return;
  }

  const rows = Utilities.parseCsv(files.next().getBlob().getDataAsString());
  if (!rows || rows.length < 2) {
    console.log('Member.txt is empty or has no data rows.');
    return;
  }

  const header = rows[0];
  const data = rows.slice(1);

  console.log('\n=== 1. Member.txt header (index: name) ===\n');
  header.forEach((h, i) => {
    const mark = (i === 16) ? '   <-- read as expiration' :
      (i === 21) ? '   <-- read as type' :
      (i === 24) ? '   <-- read as status' :
      (i === 0) ? '   <-- read as CAPID' : '';
    console.log(`[${String(i).padStart(2)}] ${String(h).padEnd(28)}${mark}`);
  });
  console.log(`\nTotal columns: ${header.length}   Total data rows: ${data.length}`);

  console.log('\n=== 2. Expiration [16] distribution by status [24] ===\n');
  const byStatus = {};
  for (let i = 0; i < data.length; i++) {
    const status = String(data[i][24] || '(blank)').trim();
    const exp = String(data[i][16] || '(blank)').trim();
    if (!byStatus[status]) byStatus[status] = {};
    byStatus[status][exp] = (byStatus[status][exp] || 0) + 1;
  }
  Object.keys(byStatus).sort().forEach(status => {
    const values = byStatus[status];
    const distinct = Object.keys(values);
    const total = distinct.reduce((sum, v) => sum + values[v], 0);
    console.log(`-- status "${status}": ${total} rows, ${distinct.length} distinct expiration values`);
    distinct
      .sort((a, b) => values[b] - values[a])
      .slice(0, 8)
      .forEach(v => console.log(`     ${String(v).padEnd(14)} x${values[v]}`));
    if (distinct.length > 8) console.log(`     ... and ${distinct.length - 8} more`);
    console.log('');
  });

  console.log('=== 3. Raw rows for sample CAPIDs ===\n');
  samples.forEach(capid => {
    const row = data.find(r => String(r[0] || '').trim() === String(capid));
    if (!row) {
      console.log(`CAPID ${capid}: NOT PRESENT in Member.txt`);
      return;
    }
    console.log(`CAPID ${capid}:`);
    console.log(`     [16] expiration = "${row[16]}"`);
    console.log(`     [21] type       = "${row[21]}"`);
    console.log(`     [24] status     = "${row[24]}"`);
    console.log('');
  });
}

/**
 * Parses a CAPWATCH date ("MM/DD/YYYY") into a Date.
 *
 * @param {string} value - Raw CAPWATCH date cell
 * @returns {Date|null} Parsed date, or null if absent/unparseable
 */
function parseCapwatchDate_(value) {
  const parts = String(value || '').trim().split('/');
  if (parts.length !== 3) return null;

  const month = parseInt(parts[0], 10);
  const day = parseInt(parts[1], 10);
  const year = parseInt(parts[2], 10);
  if (isNaN(month) || isNaN(day) || isNaN(year)) return null;
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;

  const parsed = new Date(year, month - 1, day);
  // Rejects overflow like 02/31 rolling into March.
  if (parsed.getFullYear() !== year || parsed.getMonth() !== month - 1 ||
      parsed.getDate() !== day) {
    return null;
  }
  return parsed;
}

/**
 * Preview what deleteIneligibleSuspendedUsers() would delete, and who it would
 * spare. Makes no changes. Kept for symmetry with previewArchival() /
 * previewDeletion() and for running standalone from the Apps Script editor.
 *
 * @returns {Object} See deleteIneligibleSuspendedUsers()
 */
function previewIneligibleSuspendedDeletion() {
  return deleteIneligibleSuspendedUsers(true);
}

/**
 * Sends email report of license management actions
 *
 * @param {Object} summary - Summary object with reactivated, archived, and deleted arrays
 * @returns {void}
 */
function sendLicenseManagementReport(summary) {
  const subject = `License Management Report - ${new Date().toLocaleDateString()}`;
  
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
          .warning { background-color: #fff3cd; padding: 10px; border-left: 4px solid #ffc107; margin-bottom: 15px; }
          .success { background-color: #d4edda; padding: 10px; border-left: 4px solid #28a745; margin-bottom: 15px; }
        </style>
      </head>
      <body>
        <h1>License Lifecycle Management Report</h1>
        <div class="summary">
          <h3>Summary</h3>
          <p><strong>Run Date:</strong> ${new Date(summary.startTime).toLocaleString()}</p>
          <p><strong>Duration:</strong> ${Math.round(summary.duration / 1000)} seconds</p>
          <p><strong>Users Archived:</strong> ${summary.archived.length}</p>
          <p><strong>Users Deleted (long-archived):</strong> ${summary.deleted.length}</p>
          <p><strong>Users Deleted (ineligible, ${LICENSE_CONFIG.DAYS_BEFORE_DELETE_INELIGIBLE}-day):</strong> ${summary.deletedIneligible ? summary.deletedIneligible.length : 0}</p>
          <p><strong>Errors:</strong> ${summary.errors.length}</p>
        </div>
  `;

  const ineligible = summary.ineligibleResult;
  if (ineligible) {
    const previewOnly = ineligible.dryRun;

    if (ineligible.candidates.length > 0) {
      htmlBody += `
        <div class="${previewOnly ? 'warning' : 'success'}">
          <h2>Ineligible Suspended Accounts — ${previewOnly ? 'Would Be Deleted' : 'Deleted'} (${ineligible.candidates.length})</h2>
          <p>${previewOnly
            ? '<strong>Nothing was deleted.</strong> Automated deletion is not armed yet; this is what the reaper would take on its next live run.'
            : 'These accounts were permanently deleted to free seats against the 2000-user cap.'}
          Grace period is ${ineligible.graceDays} days measured from CAPWATCH membership expiry.
          "No CAPWATCH record" means the member lapsed beyond the extract's ~3-month window.</p>
        </div>
        <table>
          <tr><th>Name</th><th>CAPID</th><th>Lapsed</th><th>Email</th></tr>
      `;
      ineligible.candidates.forEach(c => {
        htmlBody += `
          <tr>
            <td>${c.name}</td>
            <td>${c.capsn}</td>
            <td>${c.daysSinceExpiration === null ? 'no CAPWATCH record' : c.daysSinceExpiration + ' days ago'}</td>
            <td>${c.email}</td>
          </tr>
        `;
      });
      htmlBody += `</table>`;
    }

    if (ineligible.withinGrace.length > 0) {
      htmlBody += `
        <div class="summary">
          <h2>Within Grace — Not Deleted (${ineligible.withinGrace.length})</h2>
          <p>These members lapsed less than ${ineligible.graceDays} days ago. They keep their
          accounts until the grace period runs out, and will be reclaimed automatically if they
          do not renew. Renewing restores eligibility and cancels the deletion.</p>
        </div>
        <table>
          <tr><th>Name</th><th>CAPID</th><th>Lapsed</th><th>Deletable In</th></tr>
      `;
      ineligible.withinGrace.forEach(c => {
        htmlBody += `
          <tr>
            <td>${c.name}</td>
            <td>${c.capsn}</td>
            <td>${c.daysSinceExpiration} days ago</td>
            <td>${ineligible.graceDays - c.daysSinceExpiration} days</td>
          </tr>
        `;
      });
      htmlBody += `</table>`;
    }

    const needsReview = ineligible.activeIneligible
      .concat(ineligible.contradicted || [])
      .concat(ineligible.unknownExpiry);
    if (needsReview.length > 0) {
      htmlBody += `
        <div class="warning">
          <h2>Needs a Human Decision (${needsReview.length})</h2>
          <p>These suspended accounts are never deleted automatically, and they hold a seat until
          someone decides. Add them to the Manual Members sheet to keep them, or delete them by
          hand.</p>
          <p><em>Ineligible by type</em> (e.g. PATRON) are current members whose expiry lies in the
          future, so there is no lapse date to measure a grace period from.
          <em>No CAPWATCH record, but alive after the extract window</em> means the account was
          signed into more recently than the oldest lapse CAPWATCH still retains
          (${ineligible.oldestRetainedLapse || 'unknown'}) — so they did not lapse, and most
          likely transferred out of the wing while remaining a current member elsewhere.</p>
        </div>
        <table>
          <tr><th>Name</th><th>CAPID</th><th>CAPWATCH</th><th>Email</th></tr>
      `;
      needsReview.forEach(c => {
        htmlBody += `
          <tr>
            <td>${c.name}</td>
            <td>${c.capsn}</td>
            <td>${c.reason}</td>
            <td>${c.email}</td>
          </tr>
        `;
      });
      htmlBody += `</table>`;
    }
  }
  
  // Archived users section
  if (summary.archived.length > 0) {
    htmlBody += `
      <div class="warning">
        <h2>⚠ Archived Users (${summary.archived.length})</h2>
        <p>These users have been suspended for over 1 year and are not active in CAPWATCH. They have been moved to archived status to free up standard licenses.</p>
      </div>
      <table>
        <tr>
          <th>Name</th>
          <th>CAPSN</th>
          <th>Email</th>
          <th>Org Unit</th>
          <th>Days Since Activity</th>
          <th>Archived At</th>
        </tr>
    `;
    
    summary.archived.forEach(user => {
      htmlBody += `
        <tr>
          <td>${user.name}</td>
          <td>${user.capsn}</td>
          <td>${user.email}</td>
          <td>${user.orgUnitPath}</td>
          <td>${user.daysSinceActivity}</td>
          <td>${new Date(user.archivedAt).toLocaleString()}</td>
        </tr>
      `;
    });
    
    htmlBody += '</table>';
  }
  
  // Deleted users section
  if (summary.deleted.length > 0) {
    htmlBody += `
      <div class="warning">
        <h2>🗑 Deleted Users (${summary.deleted.length})</h2>
        <p>These users have been archived for over 5 years and are not active in CAPWATCH. Their accounts have been permanently deleted.</p>
      </div>
      <table>
        <tr>
          <th>Name</th>
          <th>CAPSN</th>
          <th>Email</th>
          <th>Org Unit</th>
          <th>Days Since Activity</th>
          <th>Deleted At</th>
        </tr>
    `;
    
    summary.deleted.forEach(user => {
      htmlBody += `
        <tr>
          <td>${user.name}</td>
          <td>${user.capsn}</td>
          <td>${user.email}</td>
          <td>${user.orgUnitPath}</td>
          <td>${user.daysSinceActivity}</td>
          <td>${new Date(user.deletedAt).toLocaleString()}</td>
        </tr>
      `;
    });
    
    htmlBody += '</table>';
  }
  
  // Errors section
  if (summary.errors.length > 0) {
    htmlBody += `
      <div style="background-color: #f8d7da; padding: 10px; border-left: 4px solid #dc3545; margin-bottom: 15px;">
        <h2>❌ Errors (${summary.errors.length})</h2>
        <p>The following errors occurred during processing:</p>
      </div>
      <table>
        <tr>
          <th>Error Message</th>
          <th>Timestamp</th>
        </tr>
    `;
    
    summary.errors.forEach(error => {
      htmlBody += `
        <tr>
          <td>${error.message}</td>
          <td>${new Date(error.timestamp).toLocaleString()}</td>
        </tr>
      `;
    });
    
    htmlBody += '</table>';
  }
  
  // No action needed
  if (summary.archived.length === 0 && summary.deleted.length === 0) {
    htmlBody += `
      <div class="success">
        <h2>✓ No Action Needed</h2>
        <p>All user accounts are in the appropriate lifecycle stage. No archival or deletions were necessary.</p>
      </div>
    `;
  }
  
  htmlBody += `
        <hr>
        <p style="font-size: 12px; color: #666;">
          This is an automated report from the MIWG CAPWATCH Automation system.
          For questions or issues, please contact the IT administrator (${ITSUPPORT_EMAIL}).
        </p>
      </body>
    </html>
  `;
  
  // Send email to notification list
  LICENSE_CONFIG.NOTIFICATION_EMAILS.forEach(email => {
    try {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: htmlBody
      });
      Logger.info('Report email sent', { recipient: email });
    } catch (err) {
      Logger.error('Failed to send report email', {
        recipient: email,
        errorMessage: err.message
      });
    }
  });
}

/**
 * Gets current license usage statistics
 * Useful for monitoring license pool availability
 * 
 * @returns {Object} License statistics
 */
function getLicenseStatistics() {
  Logger.info('Retrieving license statistics');
  
  const stats = {
    standard: { total: 0, assigned: 0, available: 0 },
    archived: { total: 0, assigned: 0, available: 0 },
    users: {
      active: 0,
      suspended: 0,
      archived: 0
    }
  };
  
  // Note: Google doesn't provide direct license pool APIs via Apps Script
  // This would need to be done through Admin Console or API
  // We can count users by status instead
  
  let nextPageToken = '';
  
  do {
    const page = AdminDirectory.Users.list({
      customer: "my_customer",
      maxResults: 500,
      fields: 'users(suspended,archived),nextPageToken',
      pageToken: nextPageToken
    });
    
    nextPageToken = page.nextPageToken;
    
    if (page.users) {
      page.users.forEach(user => {
        if (user.archived) {
          stats.users.archived++;
        } else if (user.suspended) {
          stats.users.suspended++;
        } else {
          stats.users.active++;
        }
      });
    }
  } while (nextPageToken);
  
  Logger.info('License statistics retrieved', stats);
  return stats;
}

// ============================================================================
// TEST FUNCTIONS
// ============================================================================

/**
 * Preview what users would be archived without actually archiving them
 * Shows users who have been suspended 1+ year and are not active in CAPWATCH
 * 
 * @returns {Array<Object>} Array of users who would be archived
 */
function previewArchival() {
  Logger.info('Starting archival preview (no changes will be made)');
  
  clearCache();
  const activeMembers = getMembers();
  const activeCapsns = new Set(Object.keys(activeMembers));
  
  const candidates = [];
  const oneYearAgo = new Date();
  oneYearAgo.setDate(oneYearAgo.getDate() - LICENSE_CONFIG.DAYS_BEFORE_ARCHIVE);
  
  let nextPageToken = '';
  
  do {
    const page = AdminDirectory.Users.list({
      customer: "my_customer",
      maxResults: 500,
      query: 'isSuspended=true isArchived=false',
      fields: 'users(primaryEmail,name,orgUnitPath,creationTime,lastLoginTime,customSchemas),nextPageToken',
      pageToken: nextPageToken
    });
    
    nextPageToken = page.nextPageToken;
    
    if (page.users) {
      for (const user of page.users) {
        const capsn = user.primaryEmail.split('@')[0];
        
        if (!/^\d+$/.test(capsn)) continue;
        if (activeCapsns.has(capsn)) continue;
        
        const lastActivityDate = user.lastLoginTime ? 
          new Date(user.lastLoginTime) : 
          new Date(user.creationTime);
        
        if (lastActivityDate < oneYearAgo) {
          candidates.push({
            email: user.primaryEmail,
            capsn: capsn,
            name: `${user.name.givenName} ${user.name.familyName}`,
            orgUnitPath: user.orgUnitPath || '/',
            lastActivity: lastActivityDate.toISOString(),
            daysSinceActivity: Math.floor((new Date() - lastActivityDate) / (1000 * 60 * 60 * 24))
          });
        }
      }
    }
  } while (nextPageToken);
  
  // Sort by days since activity (oldest first)
  candidates.sort((a, b) => b.daysSinceActivity - a.daysSinceActivity);
  
  Logger.info('Archival preview completed', { 
    count: candidates.length,
    oldestDays: candidates.length > 0 ? candidates[0].daysSinceActivity : 0
  });
  
  // Log details for each candidate
  console.log('\n=== USERS THAT WOULD BE ARCHIVED ===\n');
  console.log(`Total: ${candidates.length} users\n`);
  
  if (candidates.length > 0) {
    console.log('Name'.padEnd(30) + 'CAPSN'.padEnd(10) + 'Org Unit'.padEnd(15) + 'Days Inactive');
    console.log('-'.repeat(80));
    
    candidates.forEach(user => {
      console.log(
        user.name.padEnd(30) +
        user.capsn.padEnd(10) +
        user.orgUnitPath.padEnd(15) +
        user.daysSinceActivity
      );
    });
  } else {
    console.log('No users meet the archival criteria.');
  }
  
  return candidates;
}

/**
 * Preview what users would be deleted without actually deleting them
 * Shows users who have been archived 5+ years and are not active in CAPWATCH
 * 
 * @returns {Array<Object>} Array of users who would be deleted
 */
function previewDeletion() {
  Logger.info('Starting deletion preview (no changes will be made)');
  
  clearCache();
  const activeMembers = getMembers();
  const activeCapsns = new Set(Object.keys(activeMembers));
  
  const candidates = [];
  const fiveYearsAgo = new Date();
  fiveYearsAgo.setDate(fiveYearsAgo.getDate() - LICENSE_CONFIG.DAYS_BEFORE_DELETE);
  
  let nextPageToken = '';
  
  do {
    const page = AdminDirectory.Users.list({
      customer: "my_customer",
      maxResults: 500,
      query: 'isArchived=true',
      fields: 'users(primaryEmail,name,orgUnitPath,creationTime,lastLoginTime,customSchemas),nextPageToken',
      pageToken: nextPageToken
    });
    
    nextPageToken = page.nextPageToken;
    
    if (page.users) {
      for (const user of page.users) {
        const capsn = user.primaryEmail.split('@')[0];
        
        if (!/^\d+$/.test(capsn)) continue;
        
        // Flag if user is active in CAPWATCH (shouldn't be deleted)
        const isActiveInCapwatch = activeCapsns.has(capsn);
        
        const lastActivityDate = user.lastLoginTime ? 
          new Date(user.lastLoginTime) : 
          new Date(user.creationTime);
        
        if (lastActivityDate < fiveYearsAgo) {
          candidates.push({
            email: user.primaryEmail,
            capsn: capsn,
            name: `${user.name.givenName} ${user.name.familyName}`,
            orgUnitPath: user.orgUnitPath || '/',
            lastActivity: lastActivityDate.toISOString(),
            daysSinceActivity: Math.floor((new Date() - lastActivityDate) / (1000 * 60 * 60 * 24)),
            activeInCapwatch: isActiveInCapwatch,
            warning: isActiveInCapwatch ? '⚠️ ACTIVE IN CAPWATCH - WOULD BE SKIPPED' : ''
          });
        }
      }
    }
  } while (nextPageToken);
  
  // Sort by days since activity (oldest first)
  candidates.sort((a, b) => b.daysSinceActivity - a.daysSinceActivity);
  
  // Count how many would actually be deleted vs skipped
  const wouldDelete = candidates.filter(u => !u.activeInCapwatch).length;
  const wouldSkip = candidates.filter(u => u.activeInCapwatch).length;
  
  Logger.info('Deletion preview completed', { 
    totalCandidates: candidates.length,
    wouldDelete: wouldDelete,
    wouldSkip: wouldSkip,
    oldestDays: candidates.length > 0 ? candidates[0].daysSinceActivity : 0
  });
  
  // Log details for each candidate
  console.log('\n=== USERS THAT WOULD BE DELETED ===\n');
  console.log(`Total candidates: ${candidates.length} users`);
  console.log(`Would delete: ${wouldDelete} users`);
  console.log(`Would skip (active in CAPWATCH): ${wouldSkip} users\n`);
  
  if (candidates.length > 0) {
    console.log('Name'.padEnd(30) + 'CAPSN'.padEnd(10) + 'Org Unit'.padEnd(15) + 'Days'.padEnd(8) + 'Status');
    console.log('-'.repeat(90));
    
    candidates.forEach(user => {
      console.log(
        user.name.padEnd(30) +
        user.capsn.padEnd(10) +
        user.orgUnitPath.padEnd(15) +
        user.daysSinceActivity.toString().padEnd(8) +
        (user.activeInCapwatch ? '⚠️ ACTIVE - SKIP' : '✓ Would delete')
      );
    });
  } else {
    console.log('No users meet the deletion criteria.');
  }
  
  return candidates;
}

/**
 * Preview all license lifecycle actions without making any changes
 * Shows what would be archived and deleted
 * 
 * @returns {Object} Summary of what would happen
 */
function previewLicenseLifecycle() {
  console.log('\n' + '='.repeat(80));
  console.log('LICENSE LIFECYCLE PREVIEW - NO CHANGES WILL BE MADE');
  console.log('='.repeat(80) + '\n');
  
  const archived = previewArchival();
  console.log('\n');
  const deleted = previewDeletion();
  console.log('\n');
  const ineligibleSuspended = previewIneligibleSuspendedDeletion();

  console.log('\n' + '='.repeat(80));
  console.log('SUMMARY');
  console.log('='.repeat(80));
  console.log(`Users that would be ARCHIVED: ${archived.length}`);
  console.log(`Users that would be DELETED: ${deleted.filter(u => !u.activeInCapwatch).length}`);
  console.log(`Users that would be SKIPPED (active): ${deleted.filter(u => u.activeInCapwatch).length}`);
  console.log(`Ineligible suspended, past grace (would be DELETED): ${ineligibleSuspended.candidates.length}`);
  console.log(`Ineligible suspended, within grace (spared): ${ineligibleSuspended.withinGrace.length}`);
  console.log(`Current members ineligible by type (need review): ${ineligibleSuspended.activeIneligible.length}`);
  console.log('='.repeat(80) + '\n');

  return {
    archived: archived,
    deleted: deleted,
    ineligibleSuspended: ineligibleSuspended,
    summary: {
      archivedCount: archived.length,
      deletedCount: deleted.filter(u => !u.activeInCapwatch).length,
      skippedCount: deleted.filter(u => u.activeInCapwatch).length,
      ineligibleSuspendedPastGrace: ineligibleSuspended.candidates.length,
      ineligibleSuspendedWithinGrace: ineligibleSuspended.withinGrace.length,
      activeIneligibleNeedingReview: ineligibleSuspended.activeIneligible.length
    }
  };
}

/**
 * Test function to run license management on a small subset
 * Use this to verify functionality before full deployment
 * 
 * @returns {void}
 */
function testLicenseManagement() {
  Logger.info('Starting test license management');
  
  // Temporarily reduce batch size for testing
  const originalBatchSize = LICENSE_CONFIG.MAX_BATCH_SIZE;
  LICENSE_CONFIG.MAX_BATCH_SIZE = 5;
  
  const summary = manageLicenseLifecycle();
  
  Logger.info('Test completed', summary);
  
  // Restore original batch size
  LICENSE_CONFIG.MAX_BATCH_SIZE = originalBatchSize;
}

/**
 * Test function to check license statistics
 *
 * @returns {void}
 */
function testGetLicenseStats() {
  const stats = getLicenseStatistics();
  Logger.info('Current license statistics', stats);
}

/**
 * Preview Workspace accounts for members who are ACTIVE in CAPWATCH but whose
 * member type is not in CONFIG.MEMBER_TYPES.ACTIVE (e.g. PATRON) and are
 * therefore not eligible for a seat. Cadets are always excluded - they are
 * provisioned on a separate Workspace domain and out of scope here.
 *
 * Makes no changes. Run suspendAndArchiveIneligibleMembers() after reviewing
 * this list to actually free up the licenses.
 *
 * @returns {Array<Object>} Candidates with email, capsn, name, type, orgUnitPath
 */
function previewIneligibleMembers() {
  Logger.info('Starting ineligible-member preview (no changes will be made)');

  clearCache();

  // Build CAPID -> {type, status, name} directly from Member.txt, bypassing
  // the type allow-list that getMembers()/shouldProcessMember() applies.
  const memberData = parseFile('Member');
  const memberInfoByCapid = {};
  for (let i = 0; i < memberData.length; i++) {
    const capid = String(memberData[i][0] || '').trim();
    if (!capid) continue;
    memberInfoByCapid[capid] = {
      status: memberData[i][24],
      type: memberData[i][21],
      lastName: memberData[i][2],
      firstName: memberData[i][3]
    };
  }

  const eligibleTypes = CONFIG.MEMBER_TYPES.ACTIVE;
  const activeUsers = getActiveUsers();
  const candidates = [];

  for (let i = 0; i < activeUsers.length; i++) {
    const capid = String(activeUsers[i].capid || '').trim();
    const info = memberInfoByCapid[capid];
    if (!info) continue;

    if (info.status !== 'ACTIVE') continue;
    if (info.type === 'CADET') continue;
    if (eligibleTypes.indexOf(info.type) > -1) continue;

    candidates.push({
      email: activeUsers[i].email,
      capsn: capid,
      name: `${info.firstName} ${info.lastName}`,
      type: info.type || '(blank)'
    });
  }

  candidates.sort((a, b) => a.type.localeCompare(b.type) || a.name.localeCompare(b.name));

  const byType = {};
  candidates.forEach(c => { byType[c.type] = (byType[c.type] || 0) + 1; });

  Logger.info('Ineligible-member preview completed', {
    count: candidates.length,
    byType: byType
  });

  console.log('\n=== ACTIVE WORKSPACE ACCOUNTS WITH INELIGIBLE CAPWATCH TYPE ===\n');
  console.log(`Total: ${candidates.length} accounts\n`);
  if (candidates.length > 0) {
    console.log('Name'.padEnd(30) + 'CAPSN'.padEnd(10) + 'Type'.padEnd(16) + 'Email');
    console.log('-'.repeat(90));
    candidates.forEach(c => {
      console.log(c.name.padEnd(30) + c.capsn.padEnd(10) + c.type.padEnd(16) + c.email);
    });
  } else {
    console.log('No ineligible-type accounts found.');
  }

  return candidates;
}

/**
 * Suspends Workspace accounts for members identified by
 * previewIneligibleMembers(). Run the preview first and review the list -
 * this directly affects live accounts and is not limited to one batch size
 * like manageLicenseLifecycle().
 *
 * Note: on this Workspace for Nonprofits edition, the 2000-user domain cap
 * counts ALL user accounts regardless of suspended status (confirmed via the
 * Admin Console Users export), and Archived User licenses are not
 * provisioned (confirmed via 412 errors from AdminDirectory.Users.update).
 * Suspending alone does NOT free a seat against that cap - only deleting the
 * account does. Use auditWorkspaceUsersForRemoval() / deleteIneligibleMembers()
 * once you've confirmed deletion is the right call.
 *
 * @returns {Array<Object>} Accounts that were suspended
 */
function suspendIneligibleMembers() {
  Logger.info('Starting suspension of ineligible members');

  const candidates = previewIneligibleMembers();
  const results = [];

  candidates.forEach(c => {
    try {
      executeWithRetry(() =>
        AdminDirectory.Users.update({ suspended: true }, c.email)
      );
      results.push(Object.assign({}, c, { success: true }));
      Logger.info('Ineligible member suspended', c);
    } catch (e) {
      results.push(Object.assign({}, c, { success: false, errorMessage: e.message }));
      Logger.error('Failed to suspend ineligible member', {
        email: c.email,
        capsn: c.capsn,
        errorMessage: e.message,
        errorCode: e.details?.code
      });
    }
    Utilities.sleep(100);
  });

  const succeeded = results.filter(r => r.success).length;
  const failed = results.length - succeeded;

  Logger.info('Suspension of ineligible members completed', {
    total: results.length,
    succeeded: succeeded,
    failed: failed
  });

  return results;
}

/**
 * Audits every non-admin Workspace user (active AND suspended) against
 * current CAPWATCH data, using CAPID as the join key, to find every account
 * that could be removed to free seats against the 2000-user domain cap.
 *
 * Wider than previewIneligibleMembers() (which only looks at ACTIVE
 * Workspace users with an ineligible CAPWATCH type). This also catches:
 * - ineligibleType: ACTIVE in CAPWATCH but type not in CONFIG.MEMBER_TYPES.ACTIVE
 * - lapsed: has a CAPID match but CAPWATCH status is not ACTIVE
 * - noCapwatchRecord: no CAPID match in Member.txt at all (and not a Manual
 *   Member) - e.g. former members fully removed from CAPWATCH
 *
 * Manual Members (PCR/NHQ/etc. tracked via the Manual Members sheet) and
 * cadets are always excluded - cadets are provisioned on a separate
 * Workspace domain and out of scope here.
 *
 * Makes no changes.
 *
 * @returns {Object} { ineligibleType, lapsed, noCapwatchRecord, totalCandidates }
 */
function auditWorkspaceUsersForRemoval() {
  Logger.info('Starting full Workspace-vs-CAPWATCH removal audit (no changes will be made)');

  clearCache();

  // CAPID -> {status, type, name} from Member.txt, unfiltered by type.
  const memberData = parseFile('Member');
  const memberInfoByCapid = {};
  for (let i = 0; i < memberData.length; i++) {
    const capid = String(memberData[i][0] || '').trim();
    if (!capid) continue;
    memberInfoByCapid[capid] = {
      status: memberData[i][24],
      type: memberData[i][21],
      lastName: memberData[i][2],
      firstName: memberData[i][3]
    };
  }

  // Manual Members allow-list (PCR/NHQ/etc.) - never flag these.
  const manualCapids = {};
  try {
    const ss = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Manual Members') || ss.getSheetByName('ManualMembers');
    if (sheet) {
      const rows = sheet.getDataRange().getValues();
      if (rows && rows.length > 1) {
        const header = rows[0].map(h => String(h || '').trim());
        const capidIdx = header.indexOf('CAPID');
        if (capidIdx > -1) {
          for (let r = 1; r < rows.length; r++) {
            const capid = String(rows[r][capidIdx] || '').trim();
            if (capid) manualCapids[capid] = 1;
          }
        }
      }
    }
  } catch (e) {
    Logger.warn('Unable to read Manual Members sheet for audit allow-list', {
      errorMessage: e && e.message ? e.message : String(e)
    });
  }

  const eligibleTypes = CONFIG.MEMBER_TYPES.ACTIVE;
  const ineligibleType = [];
  const lapsed = [];
  const noCapwatchRecord = [];

  let nextPageToken = '';
  do {
    const page = AdminDirectory.Users.list({
      customer: "my_customer",
      maxResults: 500,
      query: 'isAdmin=false',
      projection: 'full',
      pageToken: nextPageToken
    });

    nextPageToken = page.nextPageToken;
    if (!page.users) continue;

    for (let i = 0; i < page.users.length; i++) {
      const user = page.users[i];
      const ids = user.externalIds || [];
      const capidExt = ids.find(id => id.type === 'organization');
      const capid = capidExt ? String(capidExt.value || '').trim() :
        String(user.employeeId || '').trim();

      if (!capid) continue; // can't map - covered separately by debugListUsersMissingOrganizationExternalId()
      if (manualCapids[capid]) continue;

      const info = memberInfoByCapid[capid];
      const record = {
        email: user.primaryEmail,
        capsn: capid,
        suspended: !!user.suspended,
        orgUnitPath: user.orgUnitPath || '/'
      };

      if (!info) {
        record.name = user.name && user.name.fullName ? user.name.fullName : '';
        noCapwatchRecord.push(record);
        continue;
      }

      record.name = `${info.firstName} ${info.lastName}`;

      if (info.type === 'CADET') continue;

      if (info.status !== 'ACTIVE') {
        record.status = info.status;
        record.type = info.type || '(blank)';
        lapsed.push(record);
        continue;
      }

      if (eligibleTypes.indexOf(info.type) === -1) {
        record.type = info.type || '(blank)';
        ineligibleType.push(record);
      }
    }
  } while (nextPageToken);

  [ineligibleType, lapsed, noCapwatchRecord].forEach(list =>
    list.sort((a, b) => a.name.localeCompare(b.name))
  );

  const summary = {
    ineligibleTypeCount: ineligibleType.length,
    lapsedCount: lapsed.length,
    noCapwatchRecordCount: noCapwatchRecord.length,
    totalCandidates: ineligibleType.length + lapsed.length + noCapwatchRecord.length
  };

  Logger.info('Workspace-vs-CAPWATCH removal audit completed', summary);

  console.log('\n=== REMOVAL AUDIT (no changes made) ===\n');

  console.log(`\n-- Ineligible type, ACTIVE in CAPWATCH (${ineligibleType.length}) --`);
  ineligibleType.forEach(r =>
    console.log(`${r.name.padEnd(28)} ${r.capsn.padEnd(10)} ${r.type.padEnd(14)} ${r.suspended ? 'suspended' : 'ACTIVE WS'.padEnd(9)} ${r.email}`)
  );

  console.log(`\n-- Lapsed in CAPWATCH (status != ACTIVE) (${lapsed.length}) --`);
  lapsed.forEach(r =>
    console.log(`${r.name.padEnd(28)} ${r.capsn.padEnd(10)} ${String(r.status).padEnd(14)} ${r.suspended ? 'suspended' : 'ACTIVE WS'.padEnd(9)} ${r.email}`)
  );

  console.log(`\n-- No CAPWATCH record at all (${noCapwatchRecord.length}) --`);
  noCapwatchRecord.forEach(r =>
    console.log(`${r.name.padEnd(28)} ${r.capsn.padEnd(10)} ${r.suspended ? 'suspended' : 'ACTIVE WS'.padEnd(9)} ${r.email}`)
  );

  console.log(`\nTotal removal candidates: ${summary.totalCandidates}\n`);

  return { ineligibleType, lapsed, noCapwatchRecord, summary };
}

/**
 * Manually-reviewed exceptions for deleteIneligibleWorkspaceUsers(). Emails
 * here are skipped even if auditWorkspaceUsersForRemoval() flags them.
 *
 * (Currently empty - jared.law@cawgcap.org / jared.law.2@cawgcap.org for
 * CAPID 663829 were initially going to keep the higher-storage account, but
 * neither appears in the wing registry at all, so both are eligible for
 * deletion via the normal audit candidate list.)
 */
const DELETE_AUDIT_EXCEPTIONS = [];

/**
 * Deletes Workspace accounts identified by auditWorkspaceUsersForRemoval().
 * Re-runs the audit fresh on every call (so the Manual Members sheet is
 * always re-checked live, not from a stale snapshot) and applies two safety
 * filters on top of the audit:
 *
 * 1. Only ever deletes accounts that are currently SUSPENDED in Workspace.
 *    An account that is still active (not suspended) is left alone even if
 *    it has no CAPWATCH record - that combination means either a manual
 *    exception that hasn't been added to the Manual Members sheet yet, or a
 *    member whose suspension recently lapsed and is eligible again (this is
 *    what excludes Brigitte Furra automatically, since her suspension
 *    expired and she shows as active in Workspace).
 * 2. Skips anything in DELETE_AUDIT_EXCEPTIONS.
 *
 * Defaults to a dry run (dryRun=true): logs exactly what WOULD be deleted
 * without calling AdminDirectory.Users.remove(). Pass dryRun=false to
 * actually delete. Deletion is not reversible beyond Google's ~20-day
 * recovery window for deleted users - confirm a Data Export has completed
 * (Admin Console > Account > Data export) before running with dryRun=false.
 *
 * @param {boolean} dryRun - When true (default), no accounts are deleted.
 * @returns {Array<Object>} Accounts that were (or would be) deleted
 */
function deleteIneligibleWorkspaceUsers(dryRun = true) {
  Logger.info('Starting deleteIneligibleWorkspaceUsers', { dryRun: dryRun });

  const audit = auditWorkspaceUsersForRemoval();
  const allCandidates = [].concat(audit.ineligibleType, audit.lapsed, audit.noCapwatchRecord);

  const toDelete = allCandidates.filter(c =>
    c.suspended === true && DELETE_AUDIT_EXCEPTIONS.indexOf(c.email) === -1
  );
  const skippedActive = allCandidates.filter(c => c.suspended !== true);
  const skippedException = allCandidates.filter(c =>
    c.suspended === true && DELETE_AUDIT_EXCEPTIONS.indexOf(c.email) > -1
  );

  Logger.info('Delete candidate filtering completed', {
    totalAuditCandidates: allCandidates.length,
    toDelete: toDelete.length,
    skippedActive: skippedActive.length,
    skippedException: skippedException.length
  });

  if (skippedActive.length > 0) {
    Logger.info('Skipped (currently active in Workspace, not suspended)', {
      accounts: skippedActive.map(c => ({ name: c.name, email: c.email, capsn: c.capsn }))
    });
  }
  if (skippedException.length > 0) {
    Logger.info('Skipped (manual exception)', {
      accounts: skippedException.map(c => ({ name: c.name, email: c.email, capsn: c.capsn }))
    });
  }

  console.log(`\n=== ${dryRun ? 'DRY RUN - ' : ''}DELETE CANDIDATES (${toDelete.length}) ===\n`);
  toDelete.forEach(c => console.log(`${c.name.padEnd(28)} ${c.capsn.padEnd(10)} ${c.email}`));

  const results = [];

  if (dryRun) {
    Logger.info('Dry run complete - no accounts deleted. Call deleteIneligibleWorkspaceUsers(false) to actually delete.');
    return toDelete.map(c => Object.assign({}, c, { deleted: false, dryRun: true }));
  }

  toDelete.forEach(c => {
    try {
      executeWithRetry(() => AdminDirectory.Users.remove(c.email));
      results.push(Object.assign({}, c, { deleted: true }));
      Logger.info('Account deleted', { email: c.email, capsn: c.capsn, name: c.name });
    } catch (e) {
      results.push(Object.assign({}, c, { deleted: false, errorMessage: e.message }));
      Logger.error('Failed to delete account', {
        email: c.email,
        capsn: c.capsn,
        errorMessage: e.message,
        errorCode: e.details?.code
      });
    }
    Utilities.sleep(100);
  });

  const deleted = results.filter(r => r.deleted).length;
  Logger.info('deleteIneligibleWorkspaceUsers completed', {
    total: results.length,
    deleted: deleted,
    failed: results.length - deleted
  });

  return results;
}

/**
 * Manual function to reactivate a specific archived user
 * Use when someone rejoins after being archived
 * 
 * @param {string} email - User's email address
 * @returns {boolean} True if successful, false otherwise
 */
function manualReactivateArchivedUser(email) {
  Logger.info('Manual reactivation requested', { email: email });
  
  try {
    // Get user details
    const user = AdminDirectory.Users.get(email);
    
    if (!user.archived) {
      Logger.warn('User is not archived', { email: email });
      return false;
    }
    
    // Unarchive and unsuspend (keeps current OU)
    AdminDirectory.Users.update(
      { 
        archived: false,
        suspended: false
      },
      email
    );
    
    Logger.info('User manually reactivated', { email: email });
    
    // Send notification
    MailApp.sendEmail({
      to: LICENSE_CONFIG.NOTIFICATION_EMAILS[0],
      subject: `Manual User Reactivation: ${email}`,
      body: `User ${email} has been manually reactivated from archived status.\n\nReactivated by: ${Session.getActiveUser().getEmail()}\nTimestamp: ${new Date().toISOString()}`
    });
    
    return true;
    
  } catch (err) {
    Logger.error('Failed to manually reactivate user', {
      email: email,
      errorMessage: err.message,
      errorCode: err.details?.code
    });
    return false;
  }
}
