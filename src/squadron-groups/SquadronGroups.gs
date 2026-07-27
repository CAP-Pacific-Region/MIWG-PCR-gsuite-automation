/*******************************************************
 * Squadron-Level Group Management Module
 *
 * Version: 1.7.0
 * Filename: SquadronGroups.gs
 * Saved: 2026-07-27
 * Changes: v1.7.0 — updateAllSquadronGroups() can be resumed, and
 *   updateAllSquadronGroupsBatch() drives it in slices. The run gave up on time
 *   without recording where it got to, so every execution restarted at the top of
 *   the same list and died at the same place: on the CAWG cadet tenant the last 9
 *   of 68 squadrons had never been reached, on any run, for weeks. Their lists kept
 *   stale settings and their Cadet Lite members were never added. A stop is not a
 *   finish, and the summary now says which one happened. No-argument behavior is
 *   unchanged.
 *   v1.6.0 — managed distribution lists are created at ANYONE_CAN_POST
 *   with spamModerationLevel MODERATE, and applyGroupSettings() now enforces
 *   whoCanPostMessage and spamModerationLevel alongside allowExternalMembers.
 *   A senior on the wing tenant could not post to ca.all@cawgcadets.org: the
 *   cadet-side all-hands lists sit at ALL_IN_DOMAIN_CAN_POST, which treats the
 *   other tenant as external. Posting policy was previously left to console/GAM
 *   (see v1.2.9), so nothing reconciled it. Widening the scope is safe now only
 *   because the callers pass ANYONE_CAN_POST — see the comment in the body.
 *   v1.5.0 — command-staff DLs now follow what the unit's type actually
 *   establishes. Every cadet and composite squadron was getting a
 *   ca###.deputy-commander DL that no CAPWATCH duty can fill, because CAP
 *   establishes a plain Deputy Commander only at senior units; cadet and
 *   composite units have Deputy Commander for Cadets / for Seniors instead.
 *   Selection moves into COMMAND_STAFF_SUFFIXES_BY_UNIT_KIND_, an undetermined
 *   unit type now gets the Commander DL only rather than all four, and
 *   cleanupUnnecessaryDistributionLists() can delete the wrong ones already
 *   created. See PCR_CHANGELOG.md.
 *   v1.4.0 — the wing recruiting mailbox added to every public-contact
 *   group is now validated with sanitizeEmail() before use, not just checked for
 *   truthiness. The config it reads held a literal placeholder string, which is
 *   truthy, so updatePublicContactGroup() passed '<recruiting email dl here>' to
 *   AdminDirectory.Members.insert once per unit per run and logged the failure.
 *   A set-but-invalid value now warns and is skipped. The value itself moved to
 *   the TENANT_RECRUITING_MAILBOX Script Property (config.gs 1.11.0); blank
 *   disables the behavior and is the default.
 *   v1.3.1 — managed wing-scope display name returns
 *   CONFIG.WING_ABBREVIATION instead of literal 'CAWG' (inside the existing
 *   WING==='CA' display branch — no behavior change for CA).
 *   SQUADRON_DISTRIBUTION_TOGGLES is now profile-driven (v1.3.0) — the
 *   effective toggles come from PROFILE_.SQUADRON_DISTRIBUTION_TOGGLES
 *   (config.gs, selected by TENANT_PROFILE) via getSquadronDistributionToggles_(),
 *   so the cadet tenant no longer creates senior-only lists. The module-level
 *   const is now a fallback default only.
 *   v1.2.9: applyGroupSettings() enforces ONLY allowExternalMembers (
 *   narrowed from v1.2.8). It was a log-only stub, so making it apply the full
 *   settings block would have flipped the cadet-tenant receive lists
 *   (ca###.cadets@cawgcadets.org, live at ANYONE_CAN_POST) to
 *   ALL_MEMBERS_CAN_POST and re-broken cadet delivery. Enforcing only
 *   allowExternalMembers via the AdminGroupsSettings advanced service fixes the
 *   reported bug (nesting the external cadet group into the wing .all lists)
 *   without touching live posting policy.
 *   v1.2.7: Reconciled with live tenant code; AdminDirectory.Users.list
 *   standardized to customer:"my_customer". See PCR_CHANGELOG.md.
 *
 * Manages squadron-specific Google Groups for unit collaboration and communication:
 * - Public Contact (mixxx@miwg.cap.gov) - External-facing email for public inquiries
 * - Distribution Lists:
 *   - mixxx.allhands@miwg.cap.gov - All unit members
 *   - mixxx.cadets@miwg.cap.gov - Cadet members only
 *   - mixxx.seniors@miwg.cap.gov - Senior members only
 *   - mixxx.parents@miwg.cap.gov - Parent/guardian contacts
 *
 * All groups are configured as collaborative inboxes with conversation history enabled.
 */

/**
 * Default squadron-group toggles (fallback only).
 *
 * The AUTHORITATIVE per-tenant values come from
 * PROFILE_.SQUADRON_DISTRIBUTION_TOGGLES in config.gs, selected by the
 * TENANT_PROFILE Script Property — so a shared-code `clasp push` can't make the
 * cadets tenant create senior-only lists (e.g. `.seniors`). This object is used
 * only when a profile doesn't define the toggles. Always read the effective
 * toggles via getSquadronDistributionToggles_() (never this const directly), so
 * PROFILE_ — defined in another file — is resolved lazily at call time.
 */
const SQUADRON_DISTRIBUTION_TOGGLES_DEFAULT_ = {
  PUBLIC_CONTACT: false,
  ALLHANDS: true,
  CADETS: true,
  SENIORS: true,
  PARENTS: true,
  COMMANDER: true,
  DEPUTY_COMMANDER: true,
  DEPUTY_COMMANDER_CADETS: true,
  DEPUTY_COMMANDER_SENIORS: true
};

/**
 * Returns the effective squadron-group toggles for this tenant: the profile's
 * SQUADRON_DISTRIBUTION_TOGGLES overlaid on the defaults (so a profile may
 * specify only the toggles it wants to change). Resolved lazily because PROFILE_
 * lives in config.gs and cross-file top-level const ordering is not guaranteed.
 *
 * @returns {Object} toggle map keyed like SQUADRON_DISTRIBUTION_TOGGLES_DEFAULT_
 */
function getSquadronDistributionToggles_() {
  try {
    if (typeof PROFILE_ !== 'undefined' && PROFILE_ && PROFILE_.SQUADRON_DISTRIBUTION_TOGGLES) {
      return Object.assign({}, SQUADRON_DISTRIBUTION_TOGGLES_DEFAULT_, PROFILE_.SQUADRON_DISTRIBUTION_TOGGLES);
    }
  } catch (e) {
    // fall through to defaults
  }
  return SQUADRON_DISTRIBUTION_TOGGLES_DEFAULT_;
}

function getSquadronGroupMetadata_(squadron, label) {
  const rawUnitName = ((squadron && squadron.name) || (squadron && squadron.charter) || '').toString().trim();
  const unitName = toSentenceCaseSquadronGroups_(rawUnitName);
  const shortUnitName = abbreviateManagedSquadronGroupOrgDisplayName_(squadron) || unitName;
  const groupLabel = (label || '').toString().trim();
  return {
    name: shortUnitName && groupLabel ? `${shortUnitName} - ${groupLabel}` : (shortUnitName || groupLabel || ''),
    description: unitName && groupLabel ? `${unitName} - ${groupLabel}` : (unitName || groupLabel || '')
  };
}

function toSentenceCaseSquadronGroups_(s) {
  const str = (s || '').toString().trim();
  if (!str) return '';

  const preserve = new Set([
    'CAP', 'USAF', 'FAA', 'DOT', 'TSA', 'ICAO', 'EASA', 'HQ'
  ]);

  function isWingAcronymSquadronGroups_(tok) {
    return /^[A-Z]{2,4}WG$/.test(tok) || tok === 'PCR';
  }

  function titleTokenSquadronGroups_(tok) {
    if (!tok) return tok;
    if (/\d/.test(tok)) return tok;

    const m = tok.match(/^(.+?)([.,;:)]?)$/);
    const core = m ? m[1] : tok;
    const punct = m ? m[2] : '';

    const upper = core.toUpperCase();

    if (preserve.has(upper) || isWingAcronymSquadronGroups_(upper)) return upper + punct;

    if (core === upper) {
      return (upper.charAt(0) + upper.slice(1).toLowerCase()) + punct;
    }

    return (core.charAt(0).toUpperCase() + core.slice(1).toLowerCase()) + punct;
  }

  return str
    .split(/\s+/)
    .map(titleTokenSquadronGroups_)
    .join(' ');
}

function isCAWGTenantSquadronGroups_() {
  return String((CONFIG && CONFIG.WING) || '').trim().toUpperCase() === 'CA';
}

function stripLeadingHonorificSquadronGroups_(value) {
  return String(value || '')
    .replace(/^(lt\.?\s*col|col|maj|capt|1st\s*lt|2nd\s*lt|lt)\s+/i, '')
    .trim();
}

function abbreviateManagedGroupCommonShortNameSquadronGroups_(value) {
  const normalized = String(value || '').trim();
  if (!normalized) return '';

  if (/^los angeles county$/i.test(normalized)) return 'LA County';
  if (/^san francisco bay$/i.test(normalized)) return 'SF Bay';

  return normalized;
}

function abbreviateManagedSquadronGroupOrgDisplayName_(org) {
  if (!org || !org.name) return '';

  const fullName = toSentenceCaseSquadronGroups_(String(org.name || '').trim());
  if (!fullName) return '';

  if (!isCAWGTenantSquadronGroups_()) return fullName;

  const scope = String(org.scope || '').trim().toUpperCase();
  const unit = String(org.unit || '').trim().replace(/^0+/, '');

  if (scope === 'WING') {
    return CONFIG.WING_ABBREVIATION;
  }

  if (scope === 'GROUP') {
    const match = fullName.match(/^(.*)\bGroup\s+(\d+)\b$/i);
    if (match) {
      const shortName = abbreviateManagedGroupCommonShortNameSquadronGroups_(String(match[1] || '').trim());
      const number = String(match[2] || '').trim();
      return shortName ? `Grp ${number} ${shortName}` : `Grp ${number}`;
    }
    return unit ? `Grp ${unit} ${fullName}` : fullName;
  }

  const leadingNumberUnitMatch = fullName.match(/^(\d+(?:st|nd|rd|th))\s+(.*?)(?:\s+(?:Challenger\s+)?)?(?:Cadet|Composite)?\s*Sq(?:dn|uadron)?\b$/i);
  if (leadingNumberUnitMatch) {
    const number = String(leadingNumberUnitMatch[1] || '').trim().toUpperCase();
    const shortName = abbreviateManagedGroupCommonShortNameSquadronGroups_(
      stripLeadingHonorificSquadronGroups_(String(leadingNumberUnitMatch[2] || '').trim())
    );
    return shortName ? `Sqdn ${number} ${shortName}` : `Sqdn ${number}`;
  }

  const unitMatch = fullName.match(/^(.*?)(?:\s+(?:Challenger\s+)?)?(?:Cadet|Composite)?\s*Sq(?:dn|uadron)?\s+(\d+)\b$/i)
    || fullName.match(/^(.*?)\s+Squadron\s+(\d+)\b$/i);
  if (unitMatch) {
    const shortName = abbreviateManagedGroupCommonShortNameSquadronGroups_(
      stripLeadingHonorificSquadronGroups_(String(unitMatch[1] || '').trim())
    );
    const number = String(unitMatch[2] || '').trim();
    return shortName ? `Sqdn ${number} ${shortName}` : `Sqdn ${number}`;
  }

  return unit ? `Sqdn ${unit} ${fullName}` : fullName;
}

/**
 * Creates all squadron groups without adding members
 * Use this FIRST TIME ONLY to quickly create all group structures
 * Then run updateAllSquadronGroups() to populate membership
 *
 * This approach is much faster and more reliable for initial setup because:
 * - Groups are created one at a time
 * - No membership management (avoids API conflicts)
 * - Can be safely re-run (skips existing groups)
 * - Then regular updates handle membership smoothly
 *
 * @returns {Object} Summary of groups created
 */
function createAllSquadronGroups() {
  const start = new Date();
  const maxExecutionTime = SQUADRON_GROUP_CONFIG.MAX_EXECUTION_TIME_MS || 400000;

  Logger.info('Starting squadron groups creation (groups only, no members)', {
    maxExecutionTime: maxExecutionTime + 'ms'
  });

  clearCache();

  const summary = {
    created: [],
    alreadyExisted: [],
    errors: [],
    timedOut: false,
    processedSquadrons: 0,
    totalSquadrons: 0,
    startTime: start.toISOString()
  };

  try {
    // Get squadron data
    const squadrons = getSquadrons();
    const unitSquadrons = Object.values(squadrons).filter(sq => sq.scope === 'UNIT');
    summary.totalSquadrons = unitSquadrons.length;

    Logger.info('Creating groups for squadrons', {
      totalSquadrons: unitSquadrons.length,
      groupsPerSquadron: 5,
      totalGroupsToCreate: unitSquadrons.length * 5
    });

    // Process each squadron
    for (const squadron of unitSquadrons) {
      // Check execution time before processing each squadron
      const elapsed = new Date() - start;
      if (elapsed > maxExecutionTime) {
        Logger.warn('Execution time limit approaching - stopping gracefully', {
          elapsed: elapsed + 'ms',
          maxExecutionTime: maxExecutionTime + 'ms',
          processedSquadrons: summary.processedSquadrons,
          remainingSquadrons: unitSquadrons.length - summary.processedSquadrons
        });
        summary.timedOut = true;
        break;
      }

      try {
        const result = createSquadronGroupsOnly(squadron);

        summary.created.push(...result.created);
        summary.alreadyExisted.push(...result.alreadyExisted);
        summary.errors.push(...result.errors);
        summary.processedSquadrons++;

      } catch (err) {
        Logger.error('Failed to create squadron groups', {
          squadron: squadron.charter,
          orgid: squadron.orgid,
          errorMessage: err.message
        });
        summary.errors.push({
          squadron: squadron.charter,
          error: err.message,
          timestamp: new Date().toISOString()
        });
        summary.processedSquadrons++;
      }

      // Small delay to avoid rate limits
      Utilities.sleep(100);
    }

  } catch (err) {
    Logger.error('Squadron groups creation failed', err);
    summary.errors.push({
      message: err.message,
      timestamp: new Date().toISOString()
    });
  }

  summary.endTime = new Date().toISOString();
  summary.duration = new Date() - start;

  Logger.info('Squadron groups creation completed', {
    duration: summary.duration + 'ms',
    created: summary.created.length,
    alreadyExisted: summary.alreadyExisted.length,
    errors: summary.errors.length,
    processedSquadrons: summary.processedSquadrons,
    totalSquadrons: summary.totalSquadrons,
    timedOut: summary.timedOut
  });

  // Display summary
  console.log('\n' + '='.repeat(80));
  console.log('SQUADRON GROUPS CREATION SUMMARY');
  console.log('='.repeat(80));
  console.log(`\nGroups Created: ${summary.created.length}`);
  console.log(`Already Existed: ${summary.alreadyExisted.length}`);
  console.log(`Errors: ${summary.errors.length}`);
  console.log(`Squadrons Processed: ${summary.processedSquadrons}/${summary.totalSquadrons}`);
  console.log(`Duration: ${Math.round(summary.duration / 1000)}s`);

  if (summary.timedOut) {
    console.log('\n⚠ Execution timed out - some squadrons not processed');
    console.log('Run this function again to create remaining groups');
  } else {
    console.log('\n✓ All squadron groups created!');
    console.log('\nNext step: Run updateAllSquadronGroups() to populate membership');
  }

  if (summary.errors.length > 0) {
    console.log('\n⚠ Errors encountered:');
    summary.errors.slice(0, 5).forEach(err => {
      console.log(`  - ${err.squadron || 'Unknown'}: ${err.error || err.message}`);
    });
    if (summary.errors.length > 5) {
      console.log(`  ... and ${summary.errors.length - 5} more`);
    }
  }

  console.log('='.repeat(80) + '\n');

  return summary;
}

/**
 * Creates all groups for a single squadron (no membership)
 * Helper function for createAllSquadronGroups()
 * Respects squadron type - only creates appropriate distribution lists
 *
 * @param {Object} squadron - Squadron object
 * @returns {Object} Result with created, alreadyExisted, and errors arrays
 */
function createSquadronGroupsOnly(squadron) {
  const result = {
    created: [],
    alreadyExisted: [],
    errors: []
  };

  // Skip if squadron doesn't have proper unit number
  if (!squadron.unit || squadron.unit === 0) {
    Logger.warn('Squadron missing unit number - skipping', {
      squadron: squadron.name,
      orgid: squadron.orgid
    });
    return result;
  }

  const unitPrefix = `${squadron.wing.toLowerCase()}${String(squadron.unit).padStart(3, '0')}`;

  const groupsToCreate = [];

  if (isSquadronGroupTypeEnabled_('public-contact')) {
    const metadata = getSquadronGroupMetadata_(squadron, 'Public Contact');
    groupsToCreate.push({
      email: `${unitPrefix}${CONFIG.EMAIL_DOMAIN}`,
      name: metadata.name,
      description: metadata.description,
      type: 'public-contact'
    });
  }

  // Add distribution lists based on squadron type
  if (shouldCreateDistributionLists(squadron)) {
    // Note: We can't pass squadronMembers here since this is group creation only
    // So FLIGHT squadrons will get all 4 lists by default (safe approach)
    // They'll be corrected during membership updates
    const distLists = getDistributionListsForSquadron(squadron, []);

    for (const distList of distLists) {
      const metadata = getSquadronGroupMetadata_(squadron, distList.name);
      groupsToCreate.push({
        email: `${unitPrefix}.${distList.suffix}${CONFIG.EMAIL_DOMAIN}`,
        name: metadata.name,
        description: metadata.description,
        type: `distribution-${distList.suffix}`
      });
    }
  } else {
    Logger.info('Skipping distribution lists for squadron', {
      squadron: squadron.charter,
      type: squadron.type,
      unit: squadron.unit,
      scope: squadron.scope
    });
  }

  if (groupsToCreate.length === 0) {
    Logger.info('No squadron groups enabled for creation', {
      squadron: squadron.charter,
      unitPrefix: unitPrefix
    });
    return result;
  }

  // Create each group
  for (const groupConfig of groupsToCreate) {
    try {
      // Check if group exists
      let groupExists = false;
      try {
        AdminDirectory.Groups.get(groupConfig.email);
        groupExists = true;
      } catch (err) {
        if (err.details?.code !== ERROR_CODES.NOT_FOUND) {
          throw err;
        }
      }

      if (groupExists) {
        result.alreadyExisted.push({
          groupEmail: groupConfig.email,
          type: groupConfig.type,
          squadron: squadron.charter
        });
        Logger.info('Group already exists - skipping', {
          groupEmail: groupConfig.email,
          squadron: squadron.charter
        });
      } else {
        // Create the group
        executeWithRetry(() =>
          AdminDirectory.Groups.insert({
            email: groupConfig.email,
            name: groupConfig.name,
            description: groupConfig.description
          })
        );

        result.created.push({
          groupEmail: groupConfig.email,
          type: groupConfig.type,
          squadron: squadron.charter
        });

        Logger.info('Group created', {
          groupEmail: groupConfig.email,
          squadron: squadron.charter,
          type: groupConfig.type
        });

        // Small delay after creation
        Utilities.sleep(50);
      }

    } catch (err) {
      Logger.error('Failed to create group', {
        groupEmail: groupConfig.email,
        squadron: squadron.charter,
        type: groupConfig.type,
        errorMessage: err.message,
        errorCode: err.details?.code
      });

      result.errors.push({
        groupEmail: groupConfig.email,
        squadron: squadron.charter,
        type: groupConfig.type,
        error: err.message
      });
    }
  }

  return result;
}

/**
 * Main function to create and update all squadron groups
 * Should be scheduled to run daily after member sync
 *
 * Includes execution time protection to prevent timeout.
 *
 * **Stopping is not the same as finishing.** The loop below gives up when it runs
 * out of time, and for a long while it gave up without recording where it got to,
 * so every run restarted at the top of the same list and died at the same place.
 * The units past that point were never reached — not once, on any run. On the CAWG
 * cadet tenant that was the last 9 of 68 squadrons, starved for weeks: their lists
 * kept whatever settings they had and their members, including the Cadet Lite ones
 * added by personal email, were never reconciled. Nothing failed loudly enough to
 * notice, because a run that stops early still reports success for what it did do.
 *
 * So a paused run now hands back its position, and updateAllSquadronGroupsBatch()
 * parks it. No-argument behavior is unchanged: call it bare and it still runs from
 * the beginning until its own time limit.
 *
 * @param {{deadlineMs?: number, resume?: Object}} [options] - deadlineMs stops the
 *   run at a wall-clock instant; resume restarts from a parked position.
 * @returns {Object} Summary of actions taken, including the resume position
 */
function updateAllSquadronGroups(options) {
  const opts = options || {};
  const resume = opts.resume || null;
  const start = new Date();
  const maxExecutionTime = SQUADRON_GROUP_CONFIG.MAX_EXECUTION_TIME_MS || 400000;

  // Two independent stops: the module's own budget, and a deadline handed down by
  // the batch driver. Whichever comes first wins, so a caller can always ask for a
  // shorter slice than the config allows but never a longer one.
  const ownDeadline = start.getTime() + maxExecutionTime;
  const callerDeadline = Number(opts.deadlineMs || 0);
  const effectiveDeadline = callerDeadline > 0 ? Math.min(ownDeadline, callerDeadline) : ownDeadline;

  Logger.info('Starting squadron groups update', {
    maxExecutionTime: maxExecutionTime + 'ms',
    timeoutProtection: 'enabled',
    resuming: resume ? `${resume.squadronIndex}` : 'no'
  });

  // Clear cache to ensure fresh CAPWATCH data
  clearCache();

  const summary = {
    created: [],
    updated: [],
    errors: [],
    timedOut: false,
    complete: false,
    squadronIndex: 0,
    charterAtIndex: '',
    processedSquadrons: (resume && Number(resume.processedSquadrons)) || 0,
    totalSquadrons: 0,
    startTime: start.toISOString()
  };

  try {
    // Get squadron and member data
    const squadrons = getSquadrons();
    const members = getMembers();
    const distributionContext = buildSquadronDistributionContext_();

    // Filter to only UNIT scope squadrons (not GROUP or WING)
    const unitSquadrons = Object.values(squadrons).filter(sq => sq.scope === 'UNIT');
    summary.totalSquadrons = unitSquadrons.length;

    let squadronIndex = resolveSquadronResumePosition_(unitSquadrons, resume);

    Logger.info('Processing squadron groups', {
      totalSquadrons: unitSquadrons.length,
      startingAt: squadronIndex,
      maxExecutionTime: maxExecutionTime + 'ms'
    });

    // Process each squadron with timeout protection
    for (; squadronIndex < unitSquadrons.length; squadronIndex++) {
      const squadron = unitSquadrons[squadronIndex];

      // Check execution time before processing each squadron
      if (Date.now() >= effectiveDeadline) {
        Logger.warn('Execution time limit approaching - stopping gracefully', {
          elapsed: (new Date() - start) + 'ms',
          maxExecutionTime: maxExecutionTime + 'ms',
          pausedAt: `${squadronIndex}/${unitSquadrons.length}`,
          processedSquadrons: summary.processedSquadrons,
          remainingSquadrons: unitSquadrons.length - squadronIndex
        });
        summary.timedOut = true;
        break;
      }

      try {
        const result = updateSquadronGroups(squadron, members, squadrons, distributionContext);

        if (result.created && result.created.length > 0) {
          summary.created.push(...result.created);
        }
        if (result.updated && result.updated.length > 0) {
          summary.updated.push(...result.updated);
        }
        if (result.errors && result.errors.length > 0) {
          summary.errors.push(...result.errors);
        }

        summary.processedSquadrons++;

      } catch (err) {
        Logger.error('Failed to update squadron groups', {
          squadron: squadron.charter,
          orgid: squadron.orgid,
          errorMessage: err.message
        });
        summary.errors.push({
          squadron: squadron.charter,
          error: err.message,
          timestamp: new Date().toISOString()
        });
        summary.processedSquadrons++;
      }

      // Small delay to avoid rate limits
      Utilities.sleep(200);
    }

    summary.squadronIndex = squadronIndex;
    summary.complete = squadronIndex >= unitSquadrons.length;
    // Parked alongside the index so a resume can tell whether the list still means
    // what it meant last time — an index is only a position in an order, and the
    // order comes from CAPWATCH, which changes when units charter or fold.
    summary.charterAtIndex = summary.complete
      ? ''
      : String((unitSquadrons[squadronIndex] && unitSquadrons[squadronIndex].charter) || '');

  } catch (err) {
    Logger.error('Squadron groups update failed', err);
    summary.errors.push({
      message: err.message,
      timestamp: new Date().toISOString()
    });
  }

  summary.endTime = new Date().toISOString();
  summary.duration = new Date() - start;

  Logger.info('Squadron groups update completed', {
    duration: summary.duration + 'ms',
    created: summary.created.length,
    updated: summary.updated.length,
    errors: summary.errors.length,
    processedSquadrons: summary.processedSquadrons,
    totalSquadrons: summary.totalSquadrons,
    position: `${summary.squadronIndex}/${summary.totalSquadrons}`,
    complete: summary.complete,
    timedOut: summary.timedOut
  });

  // Note: Removed automatic email reporting
  // Results are logged and can be reviewed in execution logs

  return summary;
}

/**
 * Decides where a resumed run should pick up.
 *
 * The parked position is an index into the UNIT-scope squadron list, which is
 * rebuilt from CAPWATCH on every execution. That is fine while the roster is
 * stable and wrong the moment it is not: a unit chartering or folding shifts
 * every index after it, and resuming on the old number would silently skip or
 * repeat squadrons. So the charter that was sitting at the index is parked too,
 * and disagreement is treated as "the list moved" — start over rather than
 * resume into the wrong place. Re-processing squadrons is cheap and idempotent;
 * skipping them is the bug this whole mechanism exists to fix.
 *
 * @param {Array<Object>} unitSquadrons - Current UNIT-scope squadrons, in order
 * @param {Object|null} resume - Parked position, or null for a fresh run
 * @returns {number} Index to start from
 */
function resolveSquadronResumePosition_(unitSquadrons, resume) {
  if (!resume) return 0;

  const index = Number(resume.squadronIndex);
  if (!isFinite(index) || index <= 0) return 0;

  if (index >= unitSquadrons.length) {
    Logger.warn('Parked squadron position is past the end of the list; starting over', {
      parkedIndex: index,
      totalSquadrons: unitSquadrons.length
    });
    return 0;
  }

  const expected = String(resume.charterAtIndex || '');
  const actual = String((unitSquadrons[index] && unitSquadrons[index].charter) || '');

  if (expected && expected !== actual) {
    Logger.warn('Squadron list changed since the run was parked; starting over', {
      parkedIndex: index,
      expectedCharter: expected,
      actualCharter: actual
    });
    return 0;
  }

  return index;
}

/**
 * Where a paused run's position lives.
 *
 * SQUADRON_BATCH_INDEX is NOT new — updateSquadronGroupsBatch() has parked its
 * position there all along, in fixed slices of 10 squadrons. Reusing the same key
 * is deliberate: two entry points walking one list must share one pointer, or a
 * daily trigger on one and a manual run of the other would each advance a private
 * cursor and leave units unvisited by both. Which is the bug, wearing a hat.
 *
 * The companions are new. The legacy mechanism parks a bare integer, with no way to
 * tell whether the list still means what it meant — see
 * resolveSquadronResumePosition_() for why that matters.
 */
const SQUADRON_BATCH_INDEX_PROP_ = 'SQUADRON_BATCH_INDEX';
const SQUADRON_BATCH_CHARTER_PROP_ = 'SQUADRON_BATCH_CHARTER';
const SQUADRON_BATCH_SAVED_AT_PROP_ = 'SQUADRON_BATCH_SAVED_AT';
const SQUADRON_BATCH_TOTAL_PROP_ = 'SQUADRON_BATCH_TOTAL';
const SQUADRON_BATCH_PROCESSED_PROP_ = 'SQUADRON_BATCH_PROCESSED';
const SQUADRON_BATCH_STARTED_AT_PROP_ = 'SQUADRON_BATCH_STARTED_AT';

/** Every key this module parks, so save and clear cannot drift out of step. */
const SQUADRON_BATCH_PROPS_ = [
  SQUADRON_BATCH_INDEX_PROP_,
  SQUADRON_BATCH_CHARTER_PROP_,
  SQUADRON_BATCH_SAVED_AT_PROP_,
  SQUADRON_BATCH_TOTAL_PROP_,
  SQUADRON_BATCH_PROCESSED_PROP_,
  SQUADRON_BATCH_STARTED_AT_PROP_
];

/**
 * Wall-clock budget for one slice, in minutes.
 *
 * These tenants allow a 30-minute execution. SQUADRON_GROUP_CONFIG's own budget is
 * 29.2 minutes, which leaves 48 seconds of headroom — and the elapsed check happens
 * only BETWEEN squadrons, so a single slow unit can carry the run past the hard cap
 * and have it killed outright. A killed execution parks nothing, which is the state
 * this mechanism exists to escape. 25 minutes leaves a real margin.
 *
 * A tenant on the 6-minute tier should pass 5 explicitly.
 */
const SQUADRON_GROUPS_BATCH_DEFAULT_BUDGET_MIN_ = 25;
const SQUADRON_GROUPS_BATCH_STALE_HOURS_ = 12;

/**
 * updateAllSquadronGroups() in slices, for when one pass cannot finish inside the
 * Apps Script execution limit.
 *
 * Point the daily trigger at this instead of updateAllSquadronGroups() and the tail
 * of the squadron list stops being permanently starved: each run picks up where the
 * last one stopped, so every unit is reached within a few days at worst, rather than
 * never.
 *
 * Related, and NOT a second mechanism: updateSquadronGroupsBatch(batchSize) slices
 * the same list by COUNT — 10 squadrons per call by default, so a 68-unit wing takes
 * a week of daily runs to come round. This one slices by TIME, using whatever budget
 * it is given, which on these tenants is most of the wing per run. They share the one
 * SQUADRON_BATCH_INDEX position deliberately, so mixing them cannot strand a unit
 * between two private cursors; pick whichever pace suits and let the other alone.
 *
 * Unlike updateEmailGroupsBatch(), this parks only the POSITION, not the computed
 * data. Squadron rosters are rebuilt from CAPWATCH each slice — cheap next to the
 * API calls, and it means a resumed slice acts on today's data rather than
 * replaying a snapshot taken before the pause.
 *
 *   updateAllSquadronGroupsBatch()          // 25 minutes
 *   updateAllSquadronGroupsBatch(5)         // a shorter slice, e.g. on the 6-minute tier
 *   checkSquadronGroupsBatchStatus()        // how far along, without touching anything
 *   resetSquadronGroupsBatchProgress()      // discard the parked run and start fresh
 *
 * @param {number} [budgetMinutes=25] - Wall-clock budget for THIS execution
 * @returns {{complete: boolean, squadronIndex: number, totalSquadrons: number}}
 */
function updateAllSquadronGroupsBatch(budgetMinutes) {
  const budgetMs = Math.max(1, Number(budgetMinutes || SQUADRON_GROUPS_BATCH_DEFAULT_BUDGET_MIN_)) * 60 * 1000;
  const deadlineMs = Date.now() + budgetMs;

  const saved = loadSquadronGroupsBatchState_();
  let resume = null;

  if (saved) {
    resume = {
      squadronIndex: saved.squadronIndex,
      charterAtIndex: saved.charterAtIndex,
      processedSquadrons: saved.processedSquadrons
    };

    // A legacy position carries no total, so say so rather than print "18/0" and
    // leave someone reading the log to wonder which number is wrong.
    Logger.info('Resuming parked squadron groups run', {
      startedAt: saved.startedAt,
      resumingAt: `${saved.squadronIndex}/${saved.totalSquadrons || '?'}`,
      charter: saved.charterAtIndex || '(none parked)',
      processedSoFar: saved.processedSquadrons
    });
  }

  const summary = updateAllSquadronGroups({ deadlineMs: deadlineMs, resume: resume });

  if (summary.complete) {
    clearSquadronGroupsBatchState_();
    Logger.info('Squadron groups batch finished', {
      squadrons: summary.totalSquadrons,
      created: summary.created.length,
      updated: summary.updated.length,
      errors: summary.errors.length
    });
    console.log(`✅ Complete — ${summary.totalSquadrons} squadrons, ` +
      `${summary.created.length} created / ${summary.updated.length} updated, ` +
      `${summary.errors.length} errors.`);
  } else {
    saveSquadronGroupsBatchState_({
      startedAt: (saved && saved.startedAt) || new Date().toISOString(),
      savedAt: new Date().toISOString(),
      squadronIndex: summary.squadronIndex,
      charterAtIndex: summary.charterAtIndex,
      totalSquadrons: summary.totalSquadrons,
      processedSquadrons: summary.processedSquadrons
    });
    console.log(`⏸ Paused at squadron ${summary.squadronIndex}/${summary.totalSquadrons} ` +
      `(${summary.charterAtIndex}). Run updateAllSquadronGroupsBatch() again to continue.`);
  }

  return {
    complete: summary.complete,
    squadronIndex: summary.squadronIndex,
    totalSquadrons: summary.totalSquadrons
  };
}

/**
 * Read-only: how far the parked squadron run got.
 * @returns {void}
 */
function checkSquadronGroupsBatchStatus() {
  const saved = loadSquadronGroupsBatchState_();
  if (!saved) {
    console.log('No parked squadron groups run. The next updateAllSquadronGroupsBatch() starts fresh.');
    return;
  }
  console.log(`Parked run started ${saved.startedAt}, last saved ${saved.savedAt}`);
  console.log(`  position: squadron ${saved.squadronIndex}/${saved.totalSquadrons || '?'} ` +
    `(${saved.charterAtIndex || 'no charter parked — legacy position'})`);
  console.log(`  processed so far: ${saved.processedSquadrons}`);
}

/**
 * Discards a parked run so the next batch starts from the first squadron. Changes
 * nothing in Workspace — groups already reconciled stay reconciled.
 * @returns {void}
 */
function resetSquadronGroupsBatchProgress() {
  clearSquadronGroupsBatchState_();
  console.log('Parked squadron groups run discarded. The next updateAllSquadronGroupsBatch() starts fresh.');
}

/**
 * Reads the shared batch position.
 *
 * A position parked by the legacy updateSquadronGroupsBatch() has an index and no
 * companions; that is a valid state, not a corrupt one, and resolving it is left to
 * resolveSquadronResumePosition_() — which trusts a bare index because an unverified
 * position is still better than restarting a list that was probably fine.
 *
 * @returns {Object|null} Parked state, or null when there is none or it has gone stale
 */
function loadSquadronGroupsBatchState_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty(SQUADRON_BATCH_INDEX_PROP_);
    if (raw === null || raw === '') return null;

    const squadronIndex = parseInt(raw, 10);
    if (!isFinite(squadronIndex) || squadronIndex <= 0) return null;

    const savedAt = props.getProperty(SQUADRON_BATCH_SAVED_AT_PROP_) || '';

    // A run parked long enough ago has lost its claim on the roster it was walking.
    // Starting over costs a re-walk of squadrons already done, which is idempotent;
    // resuming into a stale order risks skipping units, which is not. A legacy
    // position carries no timestamp, so it is taken at face value rather than
    // discarded — it came from the same list and nothing suggests it is wrong.
    if (savedAt) {
      const ageHours = (Date.now() - new Date(savedAt).getTime()) / 3600000;
      if (ageHours > SQUADRON_GROUPS_BATCH_STALE_HOURS_) {
        Logger.warn('Parked squadron groups run is stale; starting over instead of resuming', {
          savedAt: savedAt,
          ageHours: Math.round(ageHours)
        });
        clearSquadronGroupsBatchState_();
        return null;
      }
    }

    // Counts are carried so progress reads cumulatively across slices. A run that
    // reports only the current slice looks like it is starting over every time,
    // which is precisely the appearance the starved tail hid behind. Legacy state
    // has neither, and 0 is the honest answer there rather than a guess.
    return {
      squadronIndex: squadronIndex,
      charterAtIndex: props.getProperty(SQUADRON_BATCH_CHARTER_PROP_) || '',
      savedAt: savedAt,
      startedAt: props.getProperty(SQUADRON_BATCH_STARTED_AT_PROP_) || savedAt,
      totalSquadrons: parseInt(props.getProperty(SQUADRON_BATCH_TOTAL_PROP_) || '0', 10) || 0,
      processedSquadrons: parseInt(props.getProperty(SQUADRON_BATCH_PROCESSED_PROP_) || '0', 10) || 0
    };
  } catch (e) {
    Logger.warn('Could not read parked squadron groups run; starting fresh', { errorMessage: e.message });
    return null;
  }
}

/**
 * @param {Object} state
 * @returns {void}
 */
function saveSquadronGroupsBatchState_(state) {
  try {
    PropertiesService.getScriptProperties().setProperties({
      [SQUADRON_BATCH_INDEX_PROP_]: String(state.squadronIndex),
      [SQUADRON_BATCH_CHARTER_PROP_]: String(state.charterAtIndex || ''),
      [SQUADRON_BATCH_SAVED_AT_PROP_]: String(state.savedAt || new Date().toISOString()),
      [SQUADRON_BATCH_TOTAL_PROP_]: String(state.totalSquadrons || 0),
      [SQUADRON_BATCH_PROCESSED_PROP_]: String(state.processedSquadrons || 0),
      [SQUADRON_BATCH_STARTED_AT_PROP_]: String(state.startedAt || new Date().toISOString())
    });

    Logger.info('Parked squadron groups run saved', {
      position: `${state.squadronIndex}/${state.totalSquadrons}`,
      charter: state.charterAtIndex,
      processedSquadrons: state.processedSquadrons
    });
  } catch (e) {
    Logger.error('Failed to park the squadron groups run - the next call will start over', {
      errorMessage: e.message
    });
  }
}

/**
 * @returns {void}
 */
function clearSquadronGroupsBatchState_() {
  try {
    const props = PropertiesService.getScriptProperties();
    SQUADRON_BATCH_PROPS_.forEach(key => props.deleteProperty(key));
  } catch (e) {
    Logger.warn('Could not clear the parked squadron groups run', { errorMessage: e.message });
  }
}

/**
 * Updates all groups for a single squadron
 *
 * @param {Object} squadron - Squadron object with unit information
 * @param {Object} members - All members indexed by CAPID
 * @param {Object} squadrons - All squadrons indexed by orgid
 * @returns {Object} Result with created, updated, and error arrays
 */
function updateSquadronGroups(squadron, members, squadrons, distributionContext) {
  const result = {
    created: [],
    updated: [],
    errors: []
  };

  // Skip if squadron doesn't have proper unit number
  if (!squadron.unit || squadron.unit === 0) {
    Logger.warn('Squadron missing unit number - skipping', {
      squadron: squadron.name,
      orgid: squadron.orgid
    });
    return result;
  }

  const unitPrefix = `${squadron.wing.toLowerCase()}${String(squadron.unit).padStart(3, '0')}`;

  Logger.info('Updating groups for squadron', {
    charter: squadron.charter,
    unitPrefix: unitPrefix
  });

  // Get squadron members
  const squadronMembers = getSquadronMembers(squadron.orgid, members);

  // 1. Create/Update Public Contact Group (mixxx@miwg.cap.gov)
  if (isSquadronGroupTypeEnabled_('public-contact')) {
    const publicGroupResult = updatePublicContactGroup(unitPrefix, squadron, squadronMembers);
    if (publicGroupResult.created) result.created.push(publicGroupResult);
    if (publicGroupResult.updated) result.updated.push(publicGroupResult);
    if (publicGroupResult.error) result.errors.push(publicGroupResult);
  }

  // 2. Create/Update Distribution Lists
  const distListsResult = updateDistributionLists(unitPrefix, squadron, squadronMembers, members, distributionContext);
  if (distListsResult.created) result.created.push(...distListsResult.created);
  if (distListsResult.updated) result.updated.push(...distListsResult.updated);
  if (distListsResult.errors) result.errors.push(...distListsResult.errors);

  return result;
}

/**
 * Creates or updates the Public Contact Group for a squadron
 * Format: mixxx@miwg.cap.gov
 * Members: Commander, Deputy Commanders, PAO, Recruiting Officer, + Wing Recruiting Mailbox
 *
 * @param {string} unitPrefix - Unit prefix (e.g., "mi100")
 * @param {Object} squadron - Squadron object
 * @param {Array} squadronMembers - Array of member objects in the squadron
 * @returns {Object} Result object
 */
function updatePublicContactGroup(unitPrefix, squadron, squadronMembers) {
  const groupEmail = `${unitPrefix}${CONFIG.EMAIL_DOMAIN}`;
  const metadata = getSquadronGroupMetadata_(squadron, 'Public Contact');
  const groupName = metadata.name;
  const description = metadata.description;

  try {
    // Get or create the group
    const group = getOrCreateGroup(groupEmail, groupName, description, {
      whoCanJoin: 'INVITED_CAN_JOIN',
      whoCanViewMembership: 'ALL_MANAGERS_CAN_VIEW',
      whoCanViewGroup: 'ANYONE_CAN_VIEW',
      whoCanPostMessage: 'ANYONE_CAN_POST',
      allowExternalMembers: 'true',
      whoCanContactOwner: 'ANYONE_CAN_CONTACT',
      messageModerationLevel: 'MODERATE_NONE',
      spamModerationLevel: 'MODERATE',
      enableCollaborativeInbox: 'true',
      includeInGlobalAddressList: SQUADRON_GROUP_CONFIG.DISTRIBUTION_LIST.INCLUDE_IN_GAL ? 'true' : 'false',
      replyTo: 'REPLY_TO_SENDER',
      sendMessageDenyNotification: 'true'
    });

    // Build member list - specific roles plus recruiting mailbox
    const desiredMembers = getPublicContactMembers(squadron, squadronMembers);

    // Add wing recruiting mailbox to all public contact groups. Blank disables
    // it. Validate rather than trust: this config held a literal placeholder for
    // a long time, and a truthiness check was enough to send that string to
    // AdminDirectory.Members.insert once per unit per run.
    const recruitingMailbox = sanitizeEmail(
      SQUADRON_GROUP_CONFIG.PUBLIC_CONTACT.RECRUITING_MAILBOX
    );
    if (recruitingMailbox) {
      desiredMembers[recruitingMailbox] = {
        email: recruitingMailbox,
        role: 'MEMBER',
        reason: 'Wing Recruiting Mailbox'
      };
    } else if (SQUADRON_GROUP_CONFIG.PUBLIC_CONTACT.RECRUITING_MAILBOX) {
      Logger.warn('TENANT_RECRUITING_MAILBOX is set but is not a valid address — skipping', {
        value: SQUADRON_GROUP_CONFIG.PUBLIC_CONTACT.RECRUITING_MAILBOX,
        groupEmail: groupEmail
      });
    }

    // Update membership
    const membershipResult = updateGroupMembership(groupEmail, desiredMembers);

    Logger.info('Public contact group updated', {
      groupEmail: groupEmail,
      squadron: squadron.charter,
      members: Object.keys(desiredMembers).length,
      added: membershipResult.added,
      removed: membershipResult.removed
    });

    return {
      groupEmail: groupEmail,
      groupName: groupName,
      type: 'public-contact',
      squadron: squadron.charter,
      created: group.created,
      updated: !group.created,
      memberCount: Object.keys(desiredMembers).length,
      changes: membershipResult
    };

  } catch (err) {
    Logger.error('Failed to update public contact group', {
      groupEmail: groupEmail,
      squadron: squadron.charter,
      errorMessage: err.message
    });
    return {
      groupEmail: groupEmail,
      squadron: squadron.charter,
      error: err.message
    };
  }
}

/**
 * Creates or updates all distribution lists for a squadron
 * Intelligently creates only relevant lists based on squadron type
 *
 * Squadron Types:
 * - COMPOSITE: Has both cadets and seniors → Create all 4 lists
 * - CADET: Has cadets (and senior staff) → Create all 4 lists
 * - SENIOR: Only seniors → Create only allhands list
 * - GROUP/WING: Administrative → Skip all distribution lists
 * - Special units (000, 999): Skip all distribution lists
 *
 * @param {string} unitPrefix - Unit prefix (e.g., "mi100")
 * @param {Object} squadron - Squadron object
 * @param {Array} squadronMembers - Array of member objects in the squadron
 * @param {Object} allMembers - All members (for parent lookup)
 * @returns {Object} Result with created, updated, and errors arrays
 */
function updateDistributionLists(unitPrefix, squadron, squadronMembers, allMembers, distributionContext) {
  const result = {
    created: [],
    updated: [],
    errors: []
  };

  // Check if squadron should have distribution lists
  if (!shouldCreateDistributionLists(squadron)) {
    Logger.info('Skipping distribution lists for squadron', {
      squadron: squadron.charter,
      orgid: squadron.orgid,
      type: squadron.type || 'Unknown',
      reason: 'Squadron type does not require distribution lists'
    });
    return result;
  }

  // Determine which distribution lists to create based on squadron type
  const distLists = getDistributionListsForSquadron(squadron, squadronMembers);

  Logger.info('Creating distribution lists for squadron', {
    squadron: squadron.charter,
    type: squadron.type || 'Unknown',
    listsToCreate: distLists.map(dl => dl.suffix)
  });

  // Create/update each distribution list
  for (const distList of distLists) {
    try {
      const groupEmail = `${unitPrefix}.${distList.suffix}${CONFIG.EMAIL_DOMAIN}`;
      const metadata = getSquadronGroupMetadata_(squadron, distList.name);
      const groupName = metadata.name;
      const description = metadata.description;

      // Get or create the group
      const group = getOrCreateGroup(groupEmail, groupName, description, {
        whoCanJoin: 'INVITED_CAN_JOIN',
        whoCanViewMembership: 'ALL_MEMBERS_CAN_VIEW',
        whoCanViewGroup: 'ALL_MEMBERS_CAN_VIEW',
        // ANYONE_CAN_POST, not ALL_MEMBERS_CAN_POST: a distribution list has to
        // accept mail from senders who are not members of it — the other tenant's
        // members (a senior on the wing domain writing to ca.all@cawgcadets.org),
        // and the original external sender on cross-tenant fan-out. Google has no
        // "members plus my other domain" value, so the openness is paired with
        // spam moderation below.
        whoCanPostMessage: 'ANYONE_CAN_POST',
        allowExternalMembers: 'true',
        whoCanContactOwner: 'ALL_MEMBERS_CAN_CONTACT',
        messageModerationLevel: 'MODERATE_NONE',
        spamModerationLevel: 'MODERATE',
        enableCollaborativeInbox: 'true',
        includeInGlobalAddressList: SQUADRON_GROUP_CONFIG.DISTRIBUTION_LIST.INCLUDE_IN_GAL ? 'true' : 'false',
        replyTo: 'REPLY_TO_SENDER'
      });

      // Build member list based on type
      const desiredMembers = getDesiredDistributionMembers_(distList, squadron, squadronMembers, distributionContext);

      // Update membership
      const membershipResult = updateGroupMembership(groupEmail, desiredMembers);

      Logger.info('Distribution list updated', {
        groupEmail: groupEmail,
        squadron: squadron.charter,
        type: distList.suffix,
        members: Object.keys(desiredMembers).length,
        added: membershipResult.added,
        removed: membershipResult.removed
      });

      const distResult = {
        groupEmail: groupEmail,
        groupName: groupName,
        type: `distribution-${distList.suffix}`,
        squadron: squadron.charter,
        created: group.created,
        updated: !group.created,
        memberCount: Object.keys(desiredMembers).length,
        changes: membershipResult
      };

      if (group.created) {
        result.created.push(distResult);
      } else {
        result.updated.push(distResult);
      }

    } catch (err) {
      Logger.error('Failed to update distribution list', {
        squadron: squadron.charter,
        suffix: distList.suffix,
        errorMessage: err.message
      });
      result.errors.push({
        squadron: squadron.charter,
        suffix: distList.suffix,
        error: err.message
      });
    }
  }

  return result;
}

let _squadronOrgTypeByIdCache = null;

function getSquadronOrgTypeByIdMap_() {
  if (_squadronOrgTypeByIdCache) {
    return _squadronOrgTypeByIdCache;
  }

  const map = {};
  const orgRows = parseFile('Organization') || [];

  for (let i = 0; i < orgRows.length; i++) {
    const row = orgRows[i] || [];
    const orgid = String(row[0] || '').trim();
    const wing = String(row[2] || '').trim().toUpperCase();
    const type = String(row[6] || '').trim().toUpperCase();

    if (!orgid || wing !== String((CONFIG && CONFIG.WING) || '').trim().toUpperCase()) continue;
    if (type) map[orgid] = type;
  }

  _squadronOrgTypeByIdCache = map;
  return map;
}

function getEffectiveSquadronType_(squadron) {
  const explicitType = String((squadron && squadron.type) || '').trim().toUpperCase();
  if (explicitType) return explicitType;

  const orgid = String((squadron && squadron.orgid) || '').trim();
  if (!orgid) return '';

  return String(getSquadronOrgTypeByIdMap_()[orgid] || '').trim().toUpperCase();
}

/**
 * Determines if a squadron should have distribution lists
 *
 * Squadron Types:
 * - COMPOSITE: Both cadets and seniors → Create all 4 lists
 * - CADET: Has cadets → Create all 4 lists
 * - FLIGHT: Smaller unit (cadets OR seniors) → Create appropriate lists
 * - SENIOR: Only seniors → Create only allhands
 * - GROUP/WING: Administrative → No distribution lists
 *
 * @param {Object} squadron - Squadron object
 * @returns {boolean} True if squadron should have distribution lists
 */
function shouldCreateDistributionLists(squadron) {
  // Skip special units
  if (['000', '999'].includes(String(squadron.unit))) {
    return false;
  }

  // Skip if not a unit-level squadron
  if (squadron.scope !== 'UNIT') {
    return false;
  }

  // Check squadron type (if available)
  const squadronType = getEffectiveSquadronType_(squadron);

  // Valid types that get distribution lists
  const validTypes = ['COMPOSITE', 'CADET', 'SENIOR', 'FLIGHT'];

  // If no type specified, default to creating (for backward compatibility)
  if (!squadronType) {
    Logger.warn('Squadron has no type specified - defaulting to create distribution lists', {
      squadron: squadron.charter,
      orgid: squadron.orgid
    });
    return true;
  }

  return validTypes.includes(squadronType);
}

/**
 * Gets the appropriate distribution lists for a squadron based on its type
 *
 * Squadron Type is from Organization.txt column 10 (Type):
 * - COMPOSITE: Both cadets and seniors → All 4 lists
 * - CADET: Cadet squadron (has senior staff too) → All 4 lists
 * - FLIGHT: Smaller unit - check members to determine
 * - SENIOR: Senior members only → Only allhands
 * - GROUP: Group headquarters (no distribution lists)
 * - WING: Wing headquarters (no distribution lists)
 *
 * @param {Object} squadron - Squadron object with .type property
 * @param {Array} squadronMembers - Optional array of squadron members (used for FLIGHT detection)
 * @returns {Array} Array of distribution list configurations
 */
function getDistributionListsForSquadron(squadron, squadronMembers) {
  const squadronType = getEffectiveSquadronType_(squadron) || 'COMPOSITE';

  // All Hands list - always included for squadrons with distribution lists
  const allHandsList = {
    suffix: 'all',
    name: 'All',
    description: 'All)',
    filter: () => true
  };

  // Cadet-specific lists
  const cadetsList = {
    suffix: 'cadets',
    name: 'Cadets',
    description: 'Cadets',
    filter: (member) => member.type === 'CADET'
  };

  const parentsList = {
    suffix: 'parents',
    name: 'Parents & Guardians',
    description: 'Parents & Guardian',
    filter: null,
    isParentList: true
  };

  // Senior-specific list
  const seniorsList = {
    suffix: 'seniors',
    name: 'Seniors',
    description: 'Seniors',
    filter: (member) => ['SENIOR', 'FIFTY YEAR', 'INDEFINITE'].includes(member.type)
  };

  const commanderList = {
    suffix: 'commander',
    name: 'Commander',
    description: 'Commander',
    dutyPositions: ['Commander']
  };

  const deputyCommanderList = {
    suffix: 'deputy-commander',
    name: 'Deputy Commander',
    description: 'Deputy Commander',
    dutyPositions: ['Deputy Commander']
  };

  const deputyCommanderCadetsList = {
    suffix: 'deputy-commander-cadets',
    name: 'Deputy Commander for Cadets',
    description: 'Deputy Commander for Cadets',
    dutyPositions: ['Deputy Commander for Cadets']
  };

  const deputyCommanderSeniorsList = {
    suffix: 'deputy-commander-seniors',
    name: 'Deputy Commander for Seniors',
    description: 'Deputy Commander for Seniors',
    dutyPositions: ['Deputy Commander for Seniors']
  };

  // Command staff is chosen by what the unit's type actually establishes, NOT by
  // "create everything and let it sit empty" — see getCommandStaffLists_().
  const commandStaffLists = getCommandStaffLists_(squadron, squadronType, squadronMembers, {
    'commander': commanderList,
    'deputy-commander': deputyCommanderList,
    'deputy-commander-cadets': deputyCommanderCadetsList,
    'deputy-commander-seniors': deputyCommanderSeniorsList
  });

  // Determine which lists to create based on squadron type
  switch (squadronType) {
    case 'COMPOSITE':
      // These have both cadets and seniors
      return filterEnabledSquadronDistributionLists_([
        allHandsList,
        cadetsList,
        seniorsList,
        parentsList
      ].concat(commandStaffLists));

    case 'CADET':
      return filterEnabledSquadronDistributionLists_([
        allHandsList,
        cadetsList,
        seniorsList,
        parentsList
      ].concat(commandStaffLists));

    case 'FLIGHT':
      // Flight can be cadet or senior - check membership to determine
      if (squadronMembers && squadronMembers.length > 0) {
        const hasCadets = squadronMembers.some(m => m.type === 'CADET');
        const hasSeniors = squadronMembers.some(m => ['SENIOR', 'FIFTY YEAR', 'INDEFINITE'].includes(m.type));

        if (hasCadets && hasSeniors) {
          // Mixed flight - treat like composite
          Logger.info('Flight has both cadets and seniors - creating all lists', {
            squadron: squadron.charter,
            type: squadronType
          });
          return filterEnabledSquadronDistributionLists_([
            allHandsList,
            cadetsList,
            seniorsList,
            parentsList
          ].concat(commandStaffLists));
        } else if (hasCadets) {
          // Cadet flight - needs cadet lists
          Logger.info('Cadet flight - creating cadet-focused lists', {
            squadron: squadron.charter,
            type: squadronType
          });
          return filterEnabledSquadronDistributionLists_([
            allHandsList,
            cadetsList,
            seniorsList,
            parentsList
          ].concat(commandStaffLists));
        } else if (hasSeniors) {
          // Senior flight - only needs allhands plus senior command staff
          Logger.info('Senior flight - creating only allhands list', {
            squadron: squadron.charter,
            type: squadronType
          });
          return filterEnabledSquadronDistributionLists_([allHandsList].concat(commandStaffLists));
        } else {
          // No members yet - default to all rosters (safe approach). Command staff
          // stays undetermined, so only the Commander DL comes along.
          Logger.warn('Flight has no members - defaulting to all rosters', {
            squadron: squadron.charter,
            type: squadronType
          });
          return filterEnabledSquadronDistributionLists_([
            allHandsList,
            cadetsList,
            seniorsList,
            parentsList
          ].concat(commandStaffLists));
        }
      } else {
        // No member data provided - default to all rosters (safe approach)
        Logger.warn('Flight type but no member data - defaulting to all rosters', {
          squadron: squadron.charter,
          type: squadronType
        });
        return filterEnabledSquadronDistributionLists_([
          allHandsList,
          cadetsList,
          seniorsList,
          parentsList
        ].concat(commandStaffLists));
      }

    case 'SENIOR':
      // Only seniors - just need all hands plus Commander/Deputy Commander
      Logger.info('Senior squadron - creating only allhands list', {
        squadron: squadron.charter,
        type: squadronType
      });
      return filterEnabledSquadronDistributionLists_([allHandsList].concat(commandStaffLists));

    default:
      // Unknown type - default to all rosters for backward compatibility
      Logger.warn('Unknown squadron type - defaulting to all rosters', {
        squadron: squadron.charter,
        type: squadronType
      });
      return filterEnabledSquadronDistributionLists_([
        allHandsList,
        cadetsList,
        seniorsList,
        parentsList
      ].concat(commandStaffLists));
  }
}

/**
 * Command-staff DL suffixes each kind of unit actually establishes.
 *
 * CAP establishes a plain **Deputy Commander only at senior units**. Cadet and
 * composite units have Deputy Commander for Cadets and Deputy Commander for
 * Seniors in its place. A `ca###.deputy-commander` DL at a cadet or composite
 * squadron therefore has no CAPWATCH duty that can ever fill it — it is created,
 * kept in the GAL, and stays empty forever, which is worse than absent: mail sent
 * to it is accepted and delivered to nobody.
 *
 * Commander is the only command-staff list valid at every unit kind, so it is all
 * an undetermined unit gets. Creating all three deputy flavors "to be safe"
 * guarantees at least two of them are wrong.
 *
 * Cadet units get Deputy Commander for Seniors as well as for Cadets: it is a
 * valid billet at both cadet and composite units (confirmed by the wing DA,
 * 2026-07-26). An earlier comment here read the absence of any DCS assignment in
 * the CAPWATCH extract as proof the billet does not exist at cadet units; it is a
 * vacancy, and the DL should exist for the day it is filled.
 *
 * @type {Object<string, string[]>}
 */
const COMMAND_STAFF_SUFFIXES_BY_UNIT_KIND_ = {
  SENIOR: ['commander', 'deputy-commander'],
  CADET_OR_COMPOSITE: ['commander', 'deputy-commander-cadets', 'deputy-commander-seniors'],
  UNDETERMINED: ['commander']
};

/**
 * Classifies a unit for command-staff purposes.
 *
 * FLIGHT is not a command-staff kind of its own: CAPWATCH types a flight as
 * FLIGHT regardless of whether it runs a cadet program, so the flight's own
 * membership decides. A flight with no members at all is undetermined.
 *
 * @param {string} squadronType - Effective CAPWATCH org type
 * @param {Array} [squadronMembers] - Members of this unit (FLIGHT resolution only)
 * @returns {string} Key into COMMAND_STAFF_SUFFIXES_BY_UNIT_KIND_
 */
function classifyUnitCommandStaffKind_(squadronType, squadronMembers) {
  const type = String(squadronType || '').trim().toUpperCase();

  if (type === 'SENIOR') return 'SENIOR';
  if (type === 'CADET' || type === 'COMPOSITE') return 'CADET_OR_COMPOSITE';

  if (type === 'FLIGHT') {
    const members = squadronMembers || [];
    if (members.some(m => m && m.type === 'CADET')) return 'CADET_OR_COMPOSITE';
    if (members.some(m => m && ['SENIOR', 'FIFTY YEAR', 'INDEFINITE'].includes(m.type))) return 'SENIOR';
  }

  return 'UNDETERMINED';
}

/**
 * Returns the command-staff distribution lists a unit should have.
 *
 * @param {Object} squadron - Squadron object (logging only)
 * @param {string} squadronType - Effective CAPWATCH org type
 * @param {Array} [squadronMembers] - Members of this unit (FLIGHT resolution only)
 * @param {Object<string, Object>} listsBySuffix - Available list configs keyed by suffix
 * @returns {Array<Object>} Command-staff list configs, in suffix order
 */
function getCommandStaffLists_(squadron, squadronType, squadronMembers, listsBySuffix) {
  const kind = classifyUnitCommandStaffKind_(squadronType, squadronMembers);
  const suffixes = COMMAND_STAFF_SUFFIXES_BY_UNIT_KIND_[kind] || COMMAND_STAFF_SUFFIXES_BY_UNIT_KIND_.UNDETERMINED;

  if (kind === 'UNDETERMINED') {
    Logger.warn('Unit type does not say which deputy commander it establishes - creating Commander DL only', {
      squadron: squadron && squadron.charter,
      orgid: squadron && squadron.orgid,
      type: squadronType || 'Unknown'
    });
  }

  return suffixes.map(suffix => listsBySuffix[suffix]).filter(Boolean);
}

function filterEnabledSquadronDistributionLists_(lists) {
  return (lists || []).filter(list => isSquadronDistributionListEnabled_(list && list.suffix));
}

function isSquadronDistributionListEnabled_(suffix) {
  const toggles = getSquadronDistributionToggles_();
  switch (String(suffix || '').toLowerCase()) {
    case 'all':
    case 'allhands':
      return !!toggles.ALLHANDS;
    case 'cadets':
      return !!toggles.CADETS;
    case 'seniors':
      return !!toggles.SENIORS;
    case 'parents':
      return !!toggles.PARENTS;
    case 'commander':
      return !!toggles.COMMANDER;
    case 'deputy-commander':
      return !!toggles.DEPUTY_COMMANDER;
    case 'deputy-commander-cadets':
      return !!toggles.DEPUTY_COMMANDER_CADETS;
    case 'deputy-commander-seniors':
      return !!toggles.DEPUTY_COMMANDER_SENIORS;
    default:
      return true;
  }
}

function isSquadronGroupTypeEnabled_(groupType) {
  switch (String(groupType || '').toLowerCase()) {
    case 'public-contact':
      return !!getSquadronDistributionToggles_().PUBLIC_CONTACT;
    default:
      return isSquadronDistributionListEnabled_(groupType);
  }
}

function buildSquadronDistributionContext_() {
  return {
    workspaceEmailByCapid: buildSquadronWorkspaceEmailByCapidMap_(),
    capwatchPrimaryEmailByCapid: buildSquadronCapwatchPrimaryEmailByCapidMap_(),
    excludedCadetsByOrgid: buildCadetLiteExcludedCadetsByOrgid_(),
    manualUserAdditionsByGroupId: getManualUserAdditionsByGroupId_()
  };
}

function buildSquadronWorkspaceEmailByCapidMap_() {
  const map = {};
  let pageToken = '';

  do {
    const page = AdminDirectory.Users.list({
      customer: "my_customer",
      maxResults: 500,
      projection: 'full',
      fields: 'users(primaryEmail,externalIds),nextPageToken',
      pageToken: pageToken || undefined
    });

    (page.users || []).forEach(user => {
      const capidField = (user.externalIds || []).find(id => id && id.type === 'organization');
      const capid = String(capidField && capidField.value || '').trim();
      const email = String(user && user.primaryEmail || '').trim().toLowerCase();
      if (capid && email) {
        map[capid] = email;
      }
    });

    pageToken = page.nextPageToken || '';
  } while (pageToken);

  Logger.info('Squadron distribution Workspace CAPID map built', {
    count: Object.keys(map).length
  });

  return map;
}

function buildSquadronCapwatchPrimaryEmailByCapidMap_() {
  const map = {};
  const contacts = parseFile('MbrContact');

  for (let i = 0; i < contacts.length; i++) {
    const row = contacts[i] || [];
    const capid = String(row[0] || '').trim();
    const type = String(row[1] || '').trim().toUpperCase();
    const priority = String(row[2] || '').trim().toUpperCase();
    const doNotContact = String(row[6] || '').trim().toUpperCase() === 'TRUE';

    if (!capid || doNotContact) continue;
    if (type !== 'EMAIL' || priority !== 'PRIMARY') continue;

    const email = sanitizeEmail(String(row[3] || '').trim());
    if (email && !map[capid]) {
      map[capid] = email.toLowerCase();
    }
  }

  Logger.info('Squadron distribution CAPWATCH primary-email map built', {
    count: Object.keys(map).length
  });

  return map;
}

function buildCadetLiteExcludedCadetsByOrgid_() {
  const byOrgid = {};

  if (
    !CONFIG ||
    CONFIG.CADET_LITE !== true ||
    !Array.isArray(CONFIG.CADET_LITE_EXCLUDED_GRADES) ||
    CONFIG.CADET_LITE_EXCLUDED_GRADES.length === 0
  ) {
    return byOrgid;
  }

  const excludedRanks = {};
  CONFIG.CADET_LITE_EXCLUDED_GRADES.forEach(rank => {
    const normalized = String(rank || '').trim();
    if (normalized) excludedRanks[normalized] = true;
  });

  const memberRows = parseFile('Member');
  for (let i = 0; i < memberRows.length; i++) {
    const row = memberRows[i] || [];
    const capsn = String(row[0] || '').trim();
    const orgid = String(row[11] || '').trim();
    const unit = String(row[13] || '').trim();
    const rank = String(row[14] || '').trim();
    const type = String(row[21] || '').trim().toUpperCase();
    const status = String(row[24] || '').trim().toUpperCase();

    if (!capsn || !orgid) continue;
    if (status !== 'ACTIVE') continue;
    if (type !== 'CADET') continue;
    if (unit === '0' || unit === '000' || unit === '999') continue;
    if (!excludedRanks[rank]) continue;

    if (!byOrgid[orgid]) byOrgid[orgid] = [];
    byOrgid[orgid].push({
      capsn: capsn,
      orgid: orgid,
      rank: rank
    });
  }

  Logger.info('Cadet Lite excluded cadets indexed for squadron distributions', {
    orgs: Object.keys(byOrgid).length,
    cadets: Object.keys(byOrgid).reduce((sum, orgid) => sum + byOrgid[orgid].length, 0)
  });

  return byOrgid;
}

function getCadetLiteExcludedCadetsForOrg_(orgid, distributionContext) {
  const byOrgid = distributionContext && distributionContext.excludedCadetsByOrgid
    ? distributionContext.excludedCadetsByOrgid
    : {};
  return byOrgid[String(orgid || '').trim()] || [];
}

function getDesiredDistributionMembers_(distList, squadron, squadronMembers, distributionContext) {
  const desiredMembers = {};
  const suffix = String(distList && distList.suffix || '').trim().toLowerCase();
  const workspaceEmailByCapid = distributionContext && distributionContext.workspaceEmailByCapid
    ? distributionContext.workspaceEmailByCapid
    : {};

  if (distList && distList.isParentList) {
    const excludedCadets = getCadetLiteExcludedCadetsForOrg_(squadron && squadron.orgid, distributionContext);
    const parentMembers = getParentContacts(squadron && squadron.orgid, squadronMembers, null, excludedCadets);
    mergeManualUserAdditionsIntoDistributionMembers_(parentMembers, squadron, suffix, distributionContext);
    return parentMembers;
  }

  if (distList && Array.isArray(distList.dutyPositions) && distList.dutyPositions.length > 0) {
    const dutyMembers = getDutyPositionMembers_(squadron, squadronMembers, distList.dutyPositions, workspaceEmailByCapid);
    mergeManualUserAdditionsIntoDistributionMembers_(dutyMembers, squadron, suffix, distributionContext);
    return dutyMembers;
  }

  (squadronMembers || [])
    .filter(member => distList && typeof distList.filter === 'function' ? distList.filter(member) : false)
    .forEach(member => {
      const capsn = String(member && member.capsn || '').trim();
      const workspaceEmail = String(workspaceEmailByCapid[capsn] || '').trim().toLowerCase();
      if (!workspaceEmail) return;

      desiredMembers[workspaceEmail] = {
        email: workspaceEmail,
        role: 'MEMBER'
      };
    });

  if (suffix === 'allhands' || suffix === 'all' || suffix === 'cadets') {
    appendCadetLiteExcludedMembers_(desiredMembers, squadron, distributionContext);
  }

  mergeManualUserAdditionsIntoDistributionMembers_(desiredMembers, squadron, suffix, distributionContext);

  return desiredMembers;
}

function getManualUserAdditionsByGroupId_() {
  const out = {};

  try {
    const sheet = SpreadsheetApp
      .openById(CONFIG.AUTOMATION_SPREADSHEET_ID)
      .getSheetByName('User Additions');

    if (!sheet) {
      Logger.warn('User Additions tab not found; skipping squadron distribution manual preserve merge');
      return out;
    }

    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const email = String(rows[i][1] || '').trim().toLowerCase();
      const role = String(rows[i][2] || 'MEMBER').trim().toUpperCase() || 'MEMBER';
      const groupsCell = String(rows[i][3] || '').trim();
      if (!email || !groupsCell) continue;

      const groupTokens = groupsCell.split(',')
        .map(group => String(group || '').trim().toLowerCase())
        .filter(Boolean);

      for (let j = 0; j < groupTokens.length; j++) {
        let groupId = groupTokens[j];
        if (groupId.endsWith(String(CONFIG.EMAIL_DOMAIN || '').toLowerCase())) {
          groupId = groupId.slice(0, -String(CONFIG.EMAIL_DOMAIN || '').length);
        }
        if (!groupId) continue;

        if (!out[groupId]) out[groupId] = {};
        out[groupId][email] = {
          email: email,
          role: role
        };
      }
    }

    Logger.info('Manual User Additions loaded for squadron distributions', {
      groups: Object.keys(out).length
    });
  } catch (err) {
    Logger.warn('Failed to load User Additions for squadron distribution preserve merge', {
      errorMessage: err.message
    });
  }

  return out;
}

function mergeManualUserAdditionsIntoDistributionMembers_(desiredMembers, squadron, suffix, distributionContext) {
  const normalizedSuffix = String(suffix || '').trim().toLowerCase();
  if (!normalizedSuffix) return desiredMembers;

  const manualByGroupId = distributionContext && distributionContext.manualUserAdditionsByGroupId
    ? distributionContext.manualUserAdditionsByGroupId
    : {};
  const groupId = `${String((squadron && squadron.wing) || '').trim().toLowerCase()}${String((squadron && squadron.unit) || '').trim().padStart(3, '0')}.${normalizedSuffix}`;
  const manualMembers = manualByGroupId[groupId] || {};

  for (const email in manualMembers) {
    desiredMembers[email] = {
      email: email,
      role: String((manualMembers[email] && manualMembers[email].role) || 'MEMBER').trim().toUpperCase() || 'MEMBER'
    };
  }

  return desiredMembers;
}

function getDutyPositionMembers_(squadron, squadronMembers, dutyPositions, workspaceEmailByCapid) {
  const members = {};
  const qualifyingPositions = (dutyPositions || []).map(position => String(position || '').trim()).filter(Boolean);

  (squadronMembers || []).forEach(member => {
    if (!member || !Array.isArray(member.dutyPositions) || member.dutyPositions.length === 0) return;

    const capsn = String(member.capsn || '').trim();
    const workspaceEmail = String((workspaceEmailByCapid && workspaceEmailByCapid[capsn]) || '').trim().toLowerCase();
    if (!workspaceEmail) return;

    const matchedPositions = [];
    member.dutyPositions.forEach(dutyPosition => {
      const value = String((dutyPosition && dutyPosition.value) || '').trim();
      const positionMatch = value.match(/^([^(]+)/);
      const charterMatch = value.match(/\(([^)]+)\)$/);
      if (!positionMatch || !charterMatch) return;

      const positionId = String(positionMatch[1] || '').trim();
      const dutyCharter = String(charterMatch[1] || '').trim();
      if (dutyCharter !== String((squadron && squadron.charter) || '').trim()) return;
      if (!qualifyingPositions.includes(positionId)) return;

      matchedPositions.push(positionId);
    });

    if (matchedPositions.length === 0) return;

    members[workspaceEmail] = {
      email: workspaceEmail,
      role: matchedPositions.includes('Commander') ? 'OWNER' : 'MEMBER',
      reason: matchedPositions.join(', '),
      capsn: capsn
    };
  });

  return members;
}

function appendCadetLiteExcludedMembers_(desiredMembers, squadron, distributionContext) {
  if (!CONFIG || CONFIG.CADET_LITE !== true) return desiredMembers;

  const excludedCadets = getCadetLiteExcludedCadetsForOrg_(squadron && squadron.orgid, distributionContext);
  const capwatchPrimaryEmailByCapid = distributionContext && distributionContext.capwatchPrimaryEmailByCapid
    ? distributionContext.capwatchPrimaryEmailByCapid
    : {};

  excludedCadets.forEach(cadet => {
    const capsn = String(cadet && cadet.capsn || '').trim();
    const email = String(capwatchPrimaryEmailByCapid[capsn] || '').trim().toLowerCase();
    if (!email) return;

    desiredMembers[email] = {
      email: email,
      role: 'MEMBER'
    };
  });

  return desiredMembers;
}

/**
 * Gets or creates a Google Group with specified settings
 *
 * @param {string} email - Group email address
 * @param {string} name - Group display name
 * @param {string} description - Group description
 * @param {Object} settings - Group settings to apply
 * @returns {Object} Group object with 'created' flag
 */
function getOrCreateGroup(email, name, description, settings = {}) {
  let group;
  let created = false;

  try {
    // Try to get existing group
    group = executeWithRetry(() => AdminDirectory.Groups.get(email));

    // Update group metadata if needed
    if (group.name !== name || group.description !== description) {
      executeWithRetry(() =>
        AdminDirectory.Groups.update({
          name: name,
          description: description
        }, email)
      );
      Logger.info('Group metadata updated', { email: email });
    }

  } catch (err) {
    if (err.details?.code === ERROR_CODES.NOT_FOUND) {
      // Group doesn't exist - create it
      try {
        group = executeWithRetry(() =>
          AdminDirectory.Groups.insert({
            email: email,
            name: name,
            description: description
          })
        );
        created = true;
        Logger.info('Group created', { email: email, name: name });
      } catch (createErr) {
        Logger.error('Failed to create group', {
          email: email,
          errorMessage: createErr.message,
          errorCode: createErr.details?.code
        });
        throw createErr;
      }
    } else {
      throw err;
    }
  }

  // Apply group settings using Groups Settings API
  try {
    applyGroupSettings(email, settings);
  } catch (settingsErr) {
    Logger.warn('Failed to apply group settings', {
      email: email,
      errorMessage: settingsErr.message
    });
    // Don't fail the entire operation if settings update fails
  }

  group.created = created;
  return group;
}

/**
 * Applies settings to a Google Group using the Groups Settings API.
 *
 * Squadron distribution lists (especially the ".all" lists) must have
 * allowExternalMembers=true so cross-tenant nested groups such as
 * ca###.cadets@cawgcadets.org can be added as members and receive mail.
 * AdminDirectory.Groups.insert does NOT accept these fields, so they have to
 * be patched separately through the AdminGroupsSettings advanced service
 * (enabled in appsscript.json; scope apps.groups.settings).
 *
 * Scope is limited to the three keys that decide whether a message from outside
 * the group reaches it — allowExternalMembers, whoCanPostMessage and
 * spamModerationLevel (see the comment in the body). Everything else the callers
 * pass (visibility, collaborative inbox, reply-to) stays console/GAM territory.
 * Only patched when the live value differs, so it is safe to run on every sync.
 *
 * @param {string} email - Group email address
 * @param {Object} settings - Settings to apply (only the managed keys are enforced)
 * @returns {void}
 */
function applyGroupSettings(email, settings) {
  try {
    const groupEmail = String(email || '').trim().toLowerCase();
    if (!groupEmail) return;

    // Enforce only the keys that govern inbound delivery from outside the group.
    //
    // whoCanPostMessage was deliberately NOT enforced in v1.2.9, because the
    // callers passed ALL_MEMBERS_CAN_POST for every distribution list and applying
    // that would have flipped the cross-tenant cadet receive lists
    // (ca###.cadets@cawgcadets.org, live at ANYONE_CAN_POST) into rejecting the
    // fan-out, which carries the original external sender. The callers now pass
    // ANYONE_CAN_POST for every managed list, so enforcing the key moves every
    // group TOWARD accepting outside mail rather than away from it — the hazard
    // that justified the narrow scope is gone.
    //
    // ANYONE_CAN_POST is genuinely open to the internet; Google has no value
    // meaning "members plus my other tenant". spamModerationLevel is managed
    // alongside it so the openness always arrives with moderation attached and
    // cannot be widened by a caller that forgets it.
    const managedKeys = [
      'allowExternalMembers',
      'whoCanPostMessage',
      'spamModerationLevel'
    ];

    const desired = {};
    managedKeys.forEach(key => {
      if (settings[key] !== undefined && settings[key] !== null && settings[key] !== '') {
        desired[key] = String(settings[key]);
      }
    });

    if (Object.keys(desired).length === 0) return;

    if (typeof DRY_RUN !== 'undefined' && DRY_RUN) {
      Logger.info('💡 [Dry-Run] Would apply group settings', {
        email: groupEmail,
        settings: desired
      });
      return;
    }

    if (typeof AdminGroupsSettings === 'undefined' || !AdminGroupsSettings.Groups || !AdminGroupsSettings.Groups.patch) {
      Logger.warn('AdminGroupsSettings API not available; cannot apply group settings', {
        email: groupEmail,
        settings: desired
      });
      return;
    }

    const existing = executeWithRetry(() => AdminGroupsSettings.Groups.get(groupEmail));
    const patch = {};
    for (const key in desired) {
      const currentValue = (existing && existing[key] != null) ? String(existing[key]) : '';
      if (currentValue !== desired[key]) {
        patch[key] = desired[key];
      }
    }

    if (Object.keys(patch).length === 0) {
      Logger.info('Group settings already correct', {
        email: groupEmail,
        settings: desired
      });
      return;
    }

    executeWithRetry(() => AdminGroupsSettings.Groups.patch(patch, groupEmail));
    Logger.info('Group settings applied', {
      email: groupEmail,
      applied: patch
    });

  } catch (err) {
    Logger.warn('Failed to apply group settings', {
      email: email,
      errorMessage: err.message
    });
  }
}

/**
 * Updates group membership to match desired state
 * Adds missing members and removes members who shouldn't be in the group
 *
 * @param {string} groupEmail - Group email address
 * @param {Object} desiredMembers - Object mapping email to member info
 * @returns {Object} Result with added and removed counts
 */
function updateGroupMembership(groupEmail, desiredMembers) {
  const result = {
    added: 0,
    removed: 0,
    failed: 0
  };

  // Get current members
  const currentMembers = getCurrentGroupMembers(groupEmail);
  const currentEmailSet = new Set(currentMembers.map(m => m.toLowerCase()));
  const desiredEmailSet = new Set(Object.keys(desiredMembers).map(e => e.toLowerCase()));

  // Add missing members
  for (const email in desiredMembers) {
    const normalizedEmail = email.toLowerCase();
    if (!currentEmailSet.has(normalizedEmail)) {
      try {
        executeWithRetry(() =>
          AdminDirectory.Members.insert({
            email: email,
            role: desiredMembers[email].role || 'MEMBER'
          }, groupEmail)
        );
        result.added++;
      } catch (err) {
        if (err.details?.code !== ERROR_CODES.CONFLICT) {
          Logger.error('Failed to add member to squadron group', {
            groupEmail: groupEmail,
            member: email,
            errorMessage: err.message,
            errorCode: err.details?.code
          });
          result.failed++;
        }
      }
    }
  }

  // Remove members who shouldn't be in the group
  for (const currentEmail of currentMembers) {
    const normalizedEmail = currentEmail.toLowerCase();
    if (!desiredEmailSet.has(normalizedEmail)) {
      try {
        executeWithRetry(() =>
          AdminDirectory.Members.remove(groupEmail, currentEmail)
        );
        result.removed++;
      } catch (err) {
        Logger.error('Failed to remove member from squadron group', {
          groupEmail: groupEmail,
          member: currentEmail,
          errorMessage: err.message,
          errorCode: err.details?.code
        });
        result.failed++;
      }
    }
  }

  return result;
}

/**
 * Gets current members of a Google Group
 *
 * @param {string} groupEmail - Group email address
 * @returns {Array<string>} Array of member email addresses
 */
function getCurrentGroupMembers(groupEmail) {
  const members = [];
  let nextPageToken = '';

  try {
    do {
      const page = AdminDirectory.Members.list(groupEmail, {
        maxResults: 200,
        pageToken: nextPageToken
      });

      if (page.members) {
        members.push(...page.members.map(m => m.email.toLowerCase()));
      }

      nextPageToken = page.nextPageToken;
    } while (nextPageToken);

  } catch (err) {
    if (err.details?.code !== ERROR_CODES.NOT_FOUND) {
      Logger.error('Failed to get group members', {
        groupEmail: groupEmail,
        errorMessage: err.message,
        errorCode: err.details?.code
      });
    }
  }

  return members;
}

/**
 * Gets all members for a specific squadron
 *
 * @param {string} orgid - Organization ID
 * @param {Object} allMembers - All members indexed by CAPID
 * @returns {Array<Object>} Array of member objects in the squadron
 */
function getSquadronMembers(orgid, allMembers) {
  return Object.values(allMembers).filter(member => member.orgid === orgid);
}

/**
 * Gets members who should be in the public contact group
 * Includes: Members with qualifying duty positions at this squadron + Unit POC
 * Uses preferred email from CAPWATCH
 *
 * @param {Object} squadron - Squadron object
 * @param {Array} squadronMembers - Array of member objects in the squadron
 * @returns {Object} Object mapping email to member info
 */
function getPublicContactMembers(squadron, squadronMembers) {
  const members = {};

  // Get unit POC from OrgContact file
  const orgContacts = parseFile('OrgContact');
  let unitPOCEmail = null;

  for (const contact of orgContacts) {
    if (contact[0] === squadron.orgid && contact[1] === 'EMAIL') {
      unitPOCEmail = sanitizeEmail(contact[3]);
      break;
    }
  }

  // Add unit POC if found and valid
  if (unitPOCEmail) {
    members[unitPOCEmail] = {
      email: unitPOCEmail,
      role: 'MEMBER',
      reason: 'Unit POC'
    };
  }

  // Get qualifying duty positions from config
  const qualifyingPositions = SQUADRON_GROUP_CONFIG.PUBLIC_CONTACT.DUTY_POSITIONS;

  Logger.info('Looking for public contact members', {
    squadron: squadron.charter,
    orgid: squadron.orgid,
    qualifyingPositions: qualifyingPositions,
    totalSquadronMembers: squadronMembers.length
  });

  // Find members with qualifying duty positions AT THIS SQUADRON
  for (const member of squadronMembers) {
    // Skip if no email
    if (!member.email) {
      continue;
    }

    // Skip if no duty positions
    if (!member.dutyPositions || member.dutyPositions.length === 0) {
      continue;
    }

    // Check each duty position
    let hasQualifyingPosition = false;
    let matchedPositions = [];

    for (const dutyPosition of member.dutyPositions) {
      // Duty position format: "Position (A/P) (CHARTER)"
      // Example: "Commander (P) (GLR-MI-100)"

      // Extract the position ID and charter from the duty position value
      const positionMatch = dutyPosition.value.match(/^([^(]+)/);
      const charterMatch = dutyPosition.value.match(/\(([^)]+)\)$/);

      if (positionMatch && charterMatch) {
        const positionId = positionMatch[1].trim();
        const dutyCharter = charterMatch[1].trim();

        // Check if this position is at THIS squadron
        if (dutyCharter === squadron.charter) {
          // Check if this is a qualifying position
          if (qualifyingPositions.includes(positionId)) {
            hasQualifyingPosition = true;
            matchedPositions.push(positionId);
          }
        }
      }
    }

    // Add member if they have a qualifying position at this squadron
    if (hasQualifyingPosition) {
      const email = member.email.toLowerCase();

      // Commander gets OWNER role, others get MEMBER
      const isCommander = matchedPositions.includes('Commander');

      members[email] = {
        email: email,
        role: isCommander ? 'OWNER' : 'MEMBER',
        reason: matchedPositions.join(', '),
        capsn: member.capsn
      };

      Logger.info('Added public contact member', {
        email: email,
        role: isCommander ? 'OWNER' : 'MEMBER',
        positions: matchedPositions,
        charter: squadron.charter
      });
    }
  }

  Logger.info('Public contact members found', {
    squadron: squadron.charter,
    memberCount: Object.keys(members).length,
    members: Object.keys(members)
  });

  return members;
}

/**
 * Gets parent/guardian contacts for squadron members
 *
 * @param {string} orgid - Organization ID
 * @param {Array} squadronMembers - Array of member objects in the squadron
 * @param {Object} allMembers - All members (for CAPID lookup)
 * @returns {Object} Object mapping email to contact info
 */
function getParentContacts(orgid, squadronMembers, allMembers, excludedCadets) {
  const contacts = {};

  // Get all contacts from MbrContact file
  const allContacts = parseFile('MbrContact');

  // Get CAPIDs for squadron cadets
  const cadetCapsns = squadronMembers
    .filter(m => m.type === 'CADET')
    .map(m => m.capsn);
  const extraCadetCapsns = (excludedCadets || [])
    .map(c => String(c && c.capsn || '').trim())
    .filter(Boolean);
  const targetCadetCapsns = Array.from(new Set(cadetCapsns.concat(extraCadetCapsns)));

  // Find parent/guardian contacts for these cadets
  for (const contact of allContacts) {
    const capsn = contact[0];
    const contactType = contact[1];
    const contactPriority = contact[2];
    const contactValue = contact[3];
    const doNotContact = contact[6];

    // Check if this is a cadet in our squadron
    if (!targetCadetCapsns.includes(capsn)) continue;

    // Check if this is a parent email contact
    if (contactType !== 'CADET PARENT EMAIL') continue;

    // Skip if marked as do not contact
    if (doNotContact === 'True') continue;

    // Sanitize and validate email
    const email = sanitizeEmail(contactValue);
    if (!email) {
      Logger.warn('Invalid parent email - skipping', {
        capsn: capsn,
        rawEmail: contactValue
      });
      continue;
    }

    // Add to contacts
    if (!contacts[email]) {
      contacts[email] = {
        email: email,
        role: 'MEMBER',
        capsns: []
      };
    }

    // Track which cadets this contact is associated with
    contacts[email].capsns.push(capsn);
  }

  Logger.info('Parent contacts retrieved', {
    orgid: orgid,
    cadets: targetCadetCapsns.length,
    parentContacts: Object.keys(contacts).length
  });

  return contacts;
}

// ============================================================================
// CLEANUP FUNCTIONS
// ============================================================================

/**
 * Deletes unnecessary distribution lists based on squadron type
 *
 * This function identifies and removes distribution lists that shouldn't exist:
 * - Senior squadrons: Removes cadets, seniors, parents (keeps only allhands)
 * - Special units (000, 999): Removes all distribution lists
 * - Group/Wing level: Removes all distribution lists
 * - Command staff the unit's type does not establish: a plain .deputy-commander at
 *   a cadet or composite unit, or .deputy-commander-cadets/-seniors at a senior
 *   unit. Flights are left alone (their kind can't be resolved without members).
 *
 * A group deleted here takes its manual "User Additions" members with it. The
 * preview lists every candidate — read it before running live.
 *
 * IMPORTANT: This function actually DELETES groups. Review the preview first!
 *
 * @param {boolean} dryRun - If true, only shows what would be deleted (default: true)
 * @returns {Object} Summary of deletions
 */
function cleanupUnnecessaryDistributionLists(dryRun = true) {
  const start = new Date();

  Logger.info('Starting distribution list cleanup', {
    dryRun: dryRun,
    note: dryRun ? 'DRY RUN - No groups will be deleted' : 'LIVE RUN - Groups will be deleted'
  });

  clearCache();

  const summary = {
    toDelete: [],
    deleted: [],
    errors: [],
    skipped: [],
    startTime: start.toISOString(),
    dryRun: dryRun
  };

  try {
    const squadrons = getSquadrons();
    const unitSquadrons = Object.values(squadrons).filter(sq => sq.scope === 'UNIT');

    Logger.info('Analyzing squadrons for cleanup', {
      totalSquadrons: unitSquadrons.length
    });

    for (const squadron of unitSquadrons) {
      const unitPrefix = `${squadron.wing.toLowerCase()}${String(squadron.unit).padStart(3, '0')}`;
      const squadronType = squadron.type ? squadron.type.toUpperCase() : '';

      // Determine which lists should NOT exist for this squadron
      const listsToDelete = getUnnecessaryDistributionLists(squadron, unitPrefix);

      for (const groupEmail of listsToDelete) {
        try {
          // Check if group exists
          let groupExists = false;
          try {
            AdminDirectory.Groups.get(groupEmail);
            groupExists = true;
          } catch (err) {
            if (err.details?.code === ERROR_CODES.NOT_FOUND) {
              summary.skipped.push({
                groupEmail: groupEmail,
                squadron: squadron.charter,
                reason: 'Does not exist'
              });
              continue;
            }
            throw err;
          }

          if (groupExists) {
            const deleteInfo = {
              groupEmail: groupEmail,
              squadron: squadron.charter,
              squadronType: squadronType,
              reason: getDeleteReason(squadron, groupEmail)
            };

            if (dryRun) {
              // Dry run - just record what would be deleted
              summary.toDelete.push(deleteInfo);
              Logger.info('Would delete group (dry run)', deleteInfo);
            } else {
              // Actually delete the group
              executeWithRetry(() =>
                AdminDirectory.Groups.remove(groupEmail)
              );

              summary.deleted.push(deleteInfo);
              Logger.info('Group deleted', deleteInfo);

              // Small delay after deletion
              Utilities.sleep(200);
            }
          }

        } catch (err) {
          Logger.error('Failed to process group', {
            groupEmail: groupEmail,
            squadron: squadron.charter,
            errorMessage: err.message,
            errorCode: err.details?.code
          });

          summary.errors.push({
            groupEmail: groupEmail,
            squadron: squadron.charter,
            error: err.message
          });
        }
      }
    }

  } catch (err) {
    Logger.error('Cleanup failed', err);
    summary.errors.push({
      message: err.message,
      timestamp: new Date().toISOString()
    });
  }

  summary.endTime = new Date().toISOString();
  summary.duration = new Date() - start;

  // Display summary
  console.log('\n' + '='.repeat(80));
  console.log(dryRun ? 'DRY RUN - PREVIEW OF DELETIONS' : 'DISTRIBUTION LIST CLEANUP SUMMARY');
  console.log('='.repeat(80));

  if (dryRun) {
    console.log(`\n⚠ DRY RUN MODE - No groups were actually deleted`);
    console.log(`Groups that would be deleted: ${summary.toDelete.length}`);

    if (summary.toDelete.length > 0) {
      console.log('\nGroups to delete:');
      console.log('-'.repeat(80));
      summary.toDelete.forEach(item => {
        console.log(`  ${item.groupEmail}`);
        console.log(`    Squadron: ${item.squadron} (${item.squadronType})`);
        console.log(`    Reason: ${item.reason}`);
      });

      console.log('\n' + '='.repeat(80));
      console.log('\nTo actually delete these groups, run:');
      console.log('  cleanupUnnecessaryDistributionLists(false)');
    } else {
      console.log('\n✓ No unnecessary distribution lists found!');
    }
  } else {
    console.log(`\nGroups Deleted: ${summary.deleted.length}`);
    console.log(`Skipped (not found): ${summary.skipped.length}`);
    console.log(`Errors: ${summary.errors.length}`);
    console.log(`Duration: ${Math.round(summary.duration / 1000)}s`);

    if (summary.deleted.length > 0) {
      console.log('\nDeleted groups:');
      console.log('-'.repeat(80));
      summary.deleted.forEach(item => {
        console.log(`  ✓ ${item.groupEmail}`);
        console.log(`    Squadron: ${item.squadron}`);
        console.log(`    Reason: ${item.reason}`);
      });
    }

    if (summary.errors.length > 0) {
      console.log('\nErrors:');
      summary.errors.forEach(err => {
        console.log(`  ✗ ${err.groupEmail}: ${err.error}`);
      });
    }
  }

  console.log('\n' + '='.repeat(80) + '\n');

  Logger.info('Cleanup completed', {
    dryRun: dryRun,
    toDelete: summary.toDelete.length,
    deleted: summary.deleted.length,
    skipped: summary.skipped.length,
    errors: summary.errors.length
  });

  return summary;
}

/**
 * Gets list of unnecessary distribution list emails for a squadron
 *
 * @param {Object} squadron - Squadron object
 * @param {string} unitPrefix - Unit prefix (e.g., "mi100")
 * @returns {Array<string>} Array of group email addresses to delete
 */
function getUnnecessaryDistributionLists(squadron, unitPrefix) {
  const listsToDelete = [];

  // Special units (000, 999) should not have ANY distribution lists
  if (['000', '999'].includes(String(squadron.unit))) {
    listsToDelete.push(
      `${unitPrefix}.allhands${CONFIG.EMAIL_DOMAIN}`,
      `${unitPrefix}.cadets${CONFIG.EMAIL_DOMAIN}`,
      `${unitPrefix}.seniors${CONFIG.EMAIL_DOMAIN}`,
      `${unitPrefix}.parents${CONFIG.EMAIL_DOMAIN}`
    );
    return listsToDelete;
  }

  // Group/Wing level should not have distribution lists
  if (squadron.scope !== 'UNIT') {
    listsToDelete.push(
      `${unitPrefix}.allhands${CONFIG.EMAIL_DOMAIN}`,
      `${unitPrefix}.cadets${CONFIG.EMAIL_DOMAIN}`,
      `${unitPrefix}.seniors${CONFIG.EMAIL_DOMAIN}`,
      `${unitPrefix}.parents${CONFIG.EMAIL_DOMAIN}`
    );
    return listsToDelete;
  }

  // Command-staff DLs for a deputy the unit's type does not establish. These were
  // created in bulk before getCommandStaffLists_() existed (a .deputy-commander at
  // every cadet and composite squadron, none of which CAPWATCH can ever populate).
  // Kind is resolved WITHOUT member data here, so FLIGHT stays undetermined and
  // keeps whatever it has.
  const squadronType = getEffectiveSquadronType_(squadron);
  const commandStaffKind = classifyUnitCommandStaffKind_(squadronType, null);
  if (commandStaffKind !== 'UNDETERMINED') {
    const authorized = COMMAND_STAFF_SUFFIXES_BY_UNIT_KIND_[commandStaffKind] || [];
    Object.keys(COMMAND_STAFF_SUFFIXES_BY_UNIT_KIND_)
      .reduce((all, kind) => all.concat(COMMAND_STAFF_SUFFIXES_BY_UNIT_KIND_[kind]), [])
      .filter((suffix, i, arr) => arr.indexOf(suffix) === i)
      .filter(suffix => authorized.indexOf(suffix) === -1)
      .forEach(suffix => listsToDelete.push(`${unitPrefix}.${suffix}${CONFIG.EMAIL_DOMAIN}`));
  }

  // Senior squadrons and senior flights should only have allhands
  if (squadronType === 'SENIOR') {
    listsToDelete.push(
      `${unitPrefix}.cadets${CONFIG.EMAIL_DOMAIN}`,
      `${unitPrefix}.seniors${CONFIG.EMAIL_DOMAIN}`,
      `${unitPrefix}.parents${CONFIG.EMAIL_DOMAIN}`
    );
    return listsToDelete;
  }

  // For FLIGHT type, we can't determine the roster lists without member data
  // So we don't delete any of them (safer to keep than delete)
  if (squadronType === 'FLIGHT') {
    Logger.info('Flight squadron - cannot determine unnecessary lists without member data', {
      squadron: squadron.charter,
      note: 'Manual review recommended'
    });
    return listsToDelete;
  }

  // Composite and Cadet squadrons should have all roster lists
  // No further deletions needed
  return listsToDelete;
}

/**
 * Gets human-readable reason for why a group should be deleted
 *
 * @param {Object} squadron - Squadron object
 * @param {string} groupEmail - Group email address
 * @returns {string} Reason for deletion
 */
function getDeleteReason(squadron, groupEmail) {
  const squadronType = getEffectiveSquadronType_(squadron) || 'Unknown';

  // Special units
  if (['000', '999'].includes(String(squadron.unit))) {
    return `Special unit (${squadron.unit}) does not need distribution lists`;
  }

  // Group/Wing level
  if (squadron.scope !== 'UNIT') {
    return `${squadron.scope} level squadron does not need distribution lists`;
  }

  // Command staff the unit's type does not establish
  if (groupEmail.includes('.deputy-commander@')) {
    return `${squadronType} unit has Deputy Commander for Cadets/Seniors, not a plain Deputy Commander`;
  }
  if (groupEmail.includes('.deputy-commander-cadets@') || groupEmail.includes('.deputy-commander-seniors@')) {
    return `${squadronType} unit establishes a plain Deputy Commander, not this deputy`;
  }

  // Senior squadron
  if (squadronType === 'SENIOR') {
    if (groupEmail.includes('.cadets@')) {
      return 'Senior squadron has no cadet program';
    }
    if (groupEmail.includes('.seniors@')) {
      return 'Senior squadron only needs allhands list (seniors redundant)';
    }
    if (groupEmail.includes('.parents@')) {
      return 'Senior squadron has no cadets (no parents needed)';
    }
  }

  return 'Unnecessary for squadron type';
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Renames existing legacy distribution lists to the currently configured suffixes.
 * Does not create any new groups.
 *
 * @returns {Object} Summary of renamed, updated, skipped, and errored groups
 */
function renameExistingDL() {
  const squadrons = getSquadrons();
  const configuredLists = getDistributionListsForSquadron({ type: 'COMPOSITE' }, []);
  const summary = {
    renamed: [],
    updated: [],
    skipped: [],
    errors: []
  };

  function getGroupOrNull_(email) {
    try {
      return executeWithRetry(() => AdminDirectory.Groups.get(email));
    } catch (err) {
      if (err.details?.code === ERROR_CODES.NOT_FOUND) {
        return null;
      }
      throw err;
    }
  }

  function getLegacyEmails_(unitPrefix, distList) {
    const suffix = String((distList && distList.suffix) || '').trim().toLowerCase();
    if (!suffix) return [];
    if (suffix === 'all') {
      return [`${unitPrefix}.allhands${CONFIG.EMAIL_DOMAIN}`];
    }
    return [`${unitPrefix}.all-${suffix}${CONFIG.EMAIL_DOMAIN}`];
  }

  function buildMetadata_(squadron, distList) {
    const nameMeta = getSquadronGroupMetadata_(squadron, distList && distList.name);
    const rawUnitName = ((squadron && squadron.name) || (squadron && squadron.charter) || '').toString().trim();
    const unitName = toSentenceCaseSquadronGroups_(rawUnitName);
    const descriptionLabel = String((distList && distList.description) || '').trim();

    return {
      name: nameMeta.name,
      description: unitName && descriptionLabel
        ? `${unitName} - ${descriptionLabel}`
        : (unitName || descriptionLabel || '')
    };
  }

  function patchGroup_(group, targetEmail, metadata) {
    const currentEmail = String((group && group.email) || '').trim().toLowerCase();
    const currentName = String((group && group.name) || '').trim();
    const currentDescription = String((group && group.description) || '').trim();
    const patch = {};

    if (currentEmail !== String(targetEmail || '').trim().toLowerCase()) {
      patch.email = String(targetEmail || '').trim().toLowerCase();
    }
    if (currentName !== String((metadata && metadata.name) || '').trim()) {
      patch.name = String((metadata && metadata.name) || '').trim();
    }
    if (currentDescription !== String((metadata && metadata.description) || '').trim()) {
      patch.description = String((metadata && metadata.description) || '').trim();
    }

    if (Object.keys(patch).length === 0) {
      return false;
    }

    executeWithRetry(() => AdminDirectory.Groups.patch(patch, group.id || currentEmail));
    return true;
  }

  Object.values(squadrons)
    .filter(squadron =>
      squadron &&
      squadron.scope === 'UNIT' &&
      !['0', '000', '999'].includes(String(squadron.unit || '').trim())
    )
    .forEach(squadron => {
      const unitPrefix = `${squadron.wing.toLowerCase()}${String(squadron.unit).padStart(3, '0')}`;

      configuredLists.forEach(distList => {
        const targetEmail = `${unitPrefix}.${distList.suffix}${CONFIG.EMAIL_DOMAIN}`;
        const metadata = buildMetadata_(squadron, distList);

        try {
          const targetGroup = getGroupOrNull_(targetEmail);

          if (targetGroup) {
            const conflictingLegacy = getLegacyEmails_(unitPrefix, distList)
              .map(email => getGroupOrNull_(email))
              .find(group => group && String(group.id || '') !== String(targetGroup.id || ''));

            if (conflictingLegacy) {
              throw new Error(`Target group already exists and legacy group also exists: ${conflictingLegacy.email}`);
            }

            if (patchGroup_(targetGroup, targetEmail, metadata)) {
              Logger.info('Distribution list updated', {
                squadron: squadron.charter,
                suffix: distList.suffix,
                email: targetEmail,
                name: metadata.name,
                description: metadata.description
              });
              summary.updated.push({
                squadron: squadron.charter,
                suffix: distList.suffix,
                email: targetEmail
              });
            } else {
              summary.skipped.push({
                squadron: squadron.charter,
                suffix: distList.suffix,
                email: targetEmail,
                reason: 'Target group already matched'
              });
            }
            return;
          }

          const legacyEmails = getLegacyEmails_(unitPrefix, distList);
          for (let i = 0; i < legacyEmails.length; i++) {
            const legacyEmail = legacyEmails[i];
            const legacyGroup = getGroupOrNull_(legacyEmail);
            if (!legacyGroup) continue;

            patchGroup_(legacyGroup, targetEmail, metadata);
            Logger.info('Distribution list renamed', {
              squadron: squadron.charter,
              suffix: distList.suffix,
              from: legacyEmail,
              to: targetEmail,
              name: metadata.name,
              description: metadata.description
            });
            summary.renamed.push({
              squadron: squadron.charter,
              suffix: distList.suffix,
              from: legacyEmail,
              to: targetEmail
            });
            return;
          }

          summary.skipped.push({
            squadron: squadron.charter,
            suffix: distList.suffix,
            email: targetEmail,
            reason: 'No existing legacy or target group found'
          });
        } catch (err) {
          summary.errors.push({
            squadron: squadron.charter,
            suffix: distList.suffix,
            email: targetEmail,
            error: err.message
          });
        }
      });
    });

  Logger.info('Existing distribution list rename completed', {
    renamed: summary.renamed.length,
    updated: summary.updated.length,
    skipped: summary.skipped.length,
    errors: summary.errors.length
  });

  return summary;
}

/**
 * Updates squadron groups for PCR-CA-346.
 * Use this from the Apps Script Run menu when you need a no-argument entry point.
 *
 * @returns {Object} Result object
 */
function updateSingleSquadron() {
  return updateSingleSquadronGroups('346');
}

/**
 * Updates squadron groups for a single squadron (for testing)
 *
 * @param {string} unitNumber - Unit number (e.g., "100", "205")
 * @returns {Object} Result object
 */
function updateSingleSquadronGroups(unitNumber) {
  Logger.info('Updating groups for single squadron', { unitNumber: unitNumber });

  clearCache();
  const squadrons = getSquadrons();
  const members = getMembers();
  const distributionContext = buildSquadronDistributionContext_();

  // Find the squadron
  const squadron = Object.values(squadrons).find(sq =>
    String(sq.unit) === String(unitNumber) && sq.scope === 'UNIT'
  );

  if (!squadron) {
    Logger.error('Squadron not found', { unitNumber: unitNumber });
    throw new Error(`Squadron ${unitNumber} not found`);
  }

  const result = updateSquadronGroups(squadron, members, squadrons, distributionContext);

  Logger.info('Single squadron update completed', {
    squadron: squadron.charter,
    result: result
  });

  return result;
}

/**
 * Preview squadron groups that would be created/updated
 * Does not make any changes
 *
 * @returns {Object} Preview object with squadron groups
 */
function previewSquadronGroups() {
  Logger.info('Starting squadron groups preview (no changes will be made)');

  clearCache();
  const squadrons = getSquadrons();
  const members = getMembers();
  const distributionContext = buildSquadronDistributionContext_();

  const unitSquadrons = Object.values(squadrons).filter(sq => sq.scope === 'UNIT');

  const preview = {
    totalSquadrons: unitSquadrons.length,
    squadrons: []
  };

  for (const squadron of unitSquadrons) {
    if (!squadron.unit || squadron.unit === 0) continue;

    const unitPrefix = `${squadron.wing.toLowerCase()}${String(squadron.unit).padStart(3, '0')}`;
    const squadronMembers = getSquadronMembers(squadron.orgid, members);
    const distLists = getDistributionListsForSquadron(squadron, squadronMembers);
    const groups = [];

    if (isSquadronGroupTypeEnabled_('public-contact')) {
      groups.push({
        email: `${unitPrefix}${CONFIG.EMAIL_DOMAIN}`,
        type: 'public-contact',
        memberCount: Object.keys(getPublicContactMembers(squadron, squadronMembers)).length
      });
    }

    distLists.forEach(distList => {
      const memberCount = Object.keys(
        getDesiredDistributionMembers_(distList, squadron, squadronMembers, distributionContext)
      ).length;

      groups.push({
        email: `${unitPrefix}.${distList.suffix}${CONFIG.EMAIL_DOMAIN}`,
        type: `distribution-${distList.suffix}`,
        memberCount: memberCount
      });
    });

    const squadronPreview = {
      charter: squadron.charter,
      unitPrefix: unitPrefix,
      totalMembers: squadronMembers.length,
      groups: groups
    };

    preview.squadrons.push(squadronPreview);
  }

  // Log preview
  console.log('\n=== SQUADRON GROUPS PREVIEW ===\n');
  console.log(`Total squadrons: ${preview.totalSquadrons}\n`);

  preview.squadrons.forEach(sq => {
    console.log(`${sq.charter} (${sq.unitPrefix}) - ${sq.totalMembers} members`);
    sq.groups.forEach(group => {
      console.log(`  ${group.email} (${group.type}): ${group.memberCount} members`);
    });
    console.log('');
  });

  Logger.info('Squadron groups preview completed', {
    totalSquadrons: preview.totalSquadrons
  });

  return preview;
}

/**
 * SPLIT FUNCTION ARCHITECTURE FOR SQUADRON GROUPS
 *
 * This file contains functions that split squadron group management into
 * separate, faster operations that can run on different schedules to avoid
 * the 6-minute Apps Script timeout.
 *
 * ADD THESE FUNCTIONS TO YOUR SquadronGroups.gs FILE
 */

// ============================================================================
// SPLIT FUNCTIONS - RUN ON DIFFERENT SCHEDULES
// ============================================================================

/**
 * Updates ONLY public contact groups for all squadrons
 * Moderate speed - typically completes in 2-3 minutes
 *
 * Schedule: Daily at 3:00 AM
 *
 * @returns {Object} Summary of updates
 */
function updatePublicContactGroupsOnly() {
  const start = new Date();
  const maxExecutionTime = SQUADRON_GROUP_CONFIG.MAX_EXECUTION_TIME_MS || 400000;

  Logger.info('Starting public contact groups update');

  if (!isSquadronGroupTypeEnabled_('public-contact')) {
    Logger.info('Public contact groups disabled by toggle');
    return {
      updated: [],
      created: [],
      errors: [],
      timedOut: false,
      processedSquadrons: 0,
      totalSquadrons: 0,
      startTime: start.toISOString(),
      endTime: new Date().toISOString(),
      duration: 0
    };
  }

  clearCache();

  const summary = {
    updated: [],
    created: [],
    errors: [],
    timedOut: false,
    processedSquadrons: 0,
    totalSquadrons: 0,
    startTime: start.toISOString()
  };

  try {
    const members = getMembers();
    const squadrons = getSquadrons();
    const orgContacts = parseFile('OrgContact');
    const unitSquadrons = Object.values(squadrons).filter(sq => sq.scope === 'UNIT');
    summary.totalSquadrons = unitSquadrons.length;

    for (const squadron of unitSquadrons) {
      // Check time
      if (new Date() - start > maxExecutionTime) {
        summary.timedOut = true;
        break;
      }

      try {
        const unitPrefix = `${squadron.wing.toLowerCase()}${String(squadron.unit).padStart(3, '0')}`;
        const squadronMembers = Object.values(members).filter(m => m.orgid === squadron.orgid);

        const result = updatePublicContactGroup(unitPrefix, squadron, squadronMembers, orgContacts);

        if (result.created) {
          summary.created.push(result);
        } else {
          summary.updated.push(result);
        }

        summary.processedSquadrons++;

      } catch (err) {
        Logger.error('Failed to update public contact group', {
          squadron: squadron.charter,
          errorMessage: err.message
        });
        summary.errors.push({
          squadron: squadron.charter,
          error: err.message
        });
        summary.processedSquadrons++;
      }
    }

  } catch (err) {
    Logger.error('Public contact groups update failed', err);
    summary.errors.push({ message: err.message });
  }

  summary.endTime = new Date().toISOString();
  summary.duration = new Date() - start;

  Logger.info('Public contact groups update completed', {
    duration: summary.duration + 'ms',
    updated: summary.updated.length,
    created: summary.created.length,
    errors: summary.errors.length,
    timedOut: summary.timedOut
  });

  return summary;
}

/**
 * Updates ONLY distribution lists for all squadrons
 * Slower operation - may take 3-5 minutes
 *
 * Schedule: Daily at 4:00 AM
 *
 * @returns {Object} Summary of updates
 */
function updateDistributionListsOnly() {
  const start = new Date();
  const maxExecutionTime = SQUADRON_GROUP_CONFIG.MAX_EXECUTION_TIME_MS || 400000;

  Logger.info('Starting distribution lists update');

  clearCache();

  const summary = {
    updated: [],
    created: [],
    errors: [],
    timedOut: false,
    processedSquadrons: 0,
    totalSquadrons: 0,
    startTime: start.toISOString()
  };

  try {
    const members = getMembers();
    const squadrons = getSquadrons();
    const distributionContext = buildSquadronDistributionContext_();
    const unitSquadrons = Object.values(squadrons).filter(sq => sq.scope === 'UNIT');
    summary.totalSquadrons = unitSquadrons.length;

    for (const squadron of unitSquadrons) {
      // Check time
      if (new Date() - start > maxExecutionTime) {
        summary.timedOut = true;
        break;
      }

      try {
        const unitPrefix = `${squadron.wing.toLowerCase()}${String(squadron.unit).padStart(3, '0')}`;
        const squadronMembers = Object.values(members).filter(m => m.orgid === squadron.orgid);

        const result = updateDistributionLists(unitPrefix, squadron, squadronMembers, members, distributionContext);

        summary.created.push(...result.created);
        summary.updated.push(...result.updated);
        summary.errors.push(...result.errors);
        summary.processedSquadrons++;

      } catch (err) {
        Logger.error('Failed to update distribution lists', {
          squadron: squadron.charter,
          errorMessage: err.message
        });
        summary.errors.push({
          squadron: squadron.charter,
          error: err.message
        });
        summary.processedSquadrons++;
      }
    }

  } catch (err) {
    Logger.error('Distribution lists update failed', err);
    summary.errors.push({ message: err.message });
  }

  summary.endTime = new Date().toISOString();
  summary.duration = new Date() - start;

  Logger.info('Distribution lists update completed', {
    duration: summary.duration + 'ms',
    updated: summary.updated.length,
    created: summary.created.length,
    errors: summary.errors.length,
    timedOut: summary.timedOut
  });

  return summary;
}

// ============================================================================
// BATCH PROCESSING - PROCESS SQUADRONS IN CHUNKS
// ============================================================================

/**
 * Updates squadron groups in batches to avoid timeout
 * Processes N squadrons per execution
 *
 * Usage:
 *   - Set up trigger to run every hour
 *   - Automatically tracks progress
 *   - Resumes where it left off
 *   - Resets to start when complete
 *
 * @param {number} batchSize - Number of squadrons to process per run (default: 10)
 * @returns {Object} Summary with continuation info
 */
function updateSquadronGroupsBatch(batchSize = 10) {
  const start = new Date();
  const scriptProperties = PropertiesService.getScriptProperties();

  // Get current position
  let currentIndex = parseInt(scriptProperties.getProperty('SQUADRON_BATCH_INDEX') || '0');

  Logger.info('Starting batch squadron groups update', {
    batchSize: batchSize,
    startingIndex: currentIndex
  });

  clearCache();

  const summary = {
    updated: [],
    created: [],
    errors: [],
    batchStartIndex: currentIndex,
    batchEndIndex: 0,
    totalSquadrons: 0,
    complete: false,
    startTime: start.toISOString()
  };

  try {
    const members = getMembers();
    const squadrons = getSquadrons();
    const distributionContext = buildSquadronDistributionContext_();
    const unitSquadrons = Object.values(squadrons).filter(sq => sq.scope === 'UNIT');

    summary.totalSquadrons = unitSquadrons.length;

    // Calculate batch boundaries
    const endIndex = Math.min(currentIndex + batchSize, unitSquadrons.length);
    const batch = unitSquadrons.slice(currentIndex, endIndex);

    Logger.info('Processing squadron batch', {
      currentIndex: currentIndex,
      endIndex: endIndex,
      batchSize: batch.length,
      totalSquadrons: unitSquadrons.length
    });

    // Process batch
    for (const squadron of batch) {
      try {
        const result = updateSquadronGroups(squadron, members, squadrons, distributionContext);

        summary.created.push(...result.created);
        summary.updated.push(...result.updated);
        summary.errors.push(...result.errors);

      } catch (err) {
        Logger.error('Failed to update squadron', {
          squadron: squadron.charter,
          errorMessage: err.message
        });
        summary.errors.push({
          squadron: squadron.charter,
          error: err.message
        });
      }
    }

    // Update position
    currentIndex = endIndex;
    summary.batchEndIndex = currentIndex;

    // Check if complete
    if (currentIndex >= unitSquadrons.length) {
      summary.complete = true;
      // Clears the companions too, not just the index — both entry points share
      // this position, and a leftover charter would outlive the run that set it.
      clearSquadronGroupsBatchState_();
      Logger.info('Batch processing complete - resetting to start');
    } else {
      scriptProperties.setProperty('SQUADRON_BATCH_INDEX', currentIndex.toString());
      Logger.info('Batch processing continuing', {
        nextIndex: currentIndex,
        remaining: unitSquadrons.length - currentIndex
      });
    }

  } catch (err) {
    Logger.error('Batch update failed', err);
    summary.errors.push({ message: err.message });
  }

  summary.endTime = new Date().toISOString();
  summary.duration = new Date() - start;

  Logger.info('Batch update completed', {
    duration: summary.duration + 'ms',
    updated: summary.updated.length,
    created: summary.created.length,
    errors: summary.errors.length,
    complete: summary.complete,
    progress: `${summary.batchEndIndex}/${summary.totalSquadrons}`
  });

  return summary;
}

/**
 * Resets batch processing to start from beginning
 * Use this if you want to force a full re-run
 */
function resetBatchProgress() {
  clearSquadronGroupsBatchState_();
  Logger.info('Batch progress reset to start');
  console.log('✓ Batch progress reset - next run will start from beginning');
}

/**
 * Checks current batch processing status
 * @returns {Object} Status information
 */
function checkBatchStatus() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const currentIndex = parseInt(scriptProperties.getProperty('SQUADRON_BATCH_INDEX') || '0');

  const squadrons = getSquadrons();
  const unitSquadrons = Object.values(squadrons).filter(sq => sq.scope === 'UNIT');
  const totalSquadrons = unitSquadrons.length;

  const status = {
    currentIndex: currentIndex,
    totalSquadrons: totalSquadrons,
    remaining: totalSquadrons - currentIndex,
    percentComplete: Math.round((currentIndex / totalSquadrons) * 100),
    isComplete: currentIndex >= totalSquadrons
  };

  console.log('\nBatch Processing Status:');
  console.log('========================');
  console.log(`Progress: ${currentIndex}/${totalSquadrons} squadrons (${status.percentComplete}%)`);
  console.log(`Remaining: ${status.remaining} squadrons`);
  console.log(`Status: ${status.isComplete ? 'Complete' : 'In Progress'}`);

  return status;
}

// ============================================================================
// TEST FUNCTIONS
// ============================================================================

/**
 * Test function to create groups for a specific squadron (no membership)
 * Good for testing the creation process
 *
 * @param {string} unitNumber - Unit number (e.g., "100", "205")
 * @returns {Object} Result object
 */
function testCreateSquadronGroups(unitNumber = '346') {
  Logger.info('=== TESTING SQUADRON GROUPS CREATION (NO MEMBERS) ===', {
    unitNumber: unitNumber
  });

  try {
    clearCache();
    const squadrons = getSquadrons();

    // Find the squadron
    const squadron = Object.values(squadrons).find(sq =>
      String(sq.unit) === String(unitNumber) && sq.scope === 'UNIT'
    );

    if (!squadron) {
      throw new Error(`Squadron ${unitNumber} not found`);
    }

    const result = createSquadronGroupsOnly(squadron);

    console.log('\n' + '='.repeat(80));
    console.log(`TEST RESULTS FOR SQUADRON ${unitNumber}`);
    console.log('='.repeat(80));

    console.log(`\nGroups Created: ${result.created.length}`);
    console.log(`Already Existed: ${result.alreadyExisted.length}`);
    console.log(`Errors: ${result.errors.length}`);

    if (result.created.length > 0) {
      console.log('\nCreated:');
      result.created.forEach(g => {
        console.log(`  ✓ ${g.groupEmail}`);
      });
    }

    if (result.alreadyExisted.length > 0) {
      console.log('\nAlready Existed:');
      result.alreadyExisted.forEach(g => {
        console.log(`  ↻ ${g.groupEmail}`);
      });
    }

    if (result.errors.length > 0) {
      console.log('\nErrors:');
      result.errors.forEach(e => {
        console.log(`  ✗ ${e.groupEmail}: ${e.error}`);
      });
    }

    console.log('\n' + '='.repeat(80) + '\n');

    Logger.info('=== TEST COMPLETED ===', { result: result });
    return result;

  } catch (err) {
    Logger.error('=== TEST FAILED ===', err);
    console.log(`\n✗ Test failed: ${err.message}`);
    throw err;
  }
}

/**
 * Helper function to list all available squadrons for testing
 * Call this first to find a valid squadron number to test with
 *
 * @param {boolean} showDetails - Whether to show detailed info for each squadron
 * @returns {Array} Array of squadron objects
 */
function listAvailableSquadrons(showDetails = false) {
  Logger.info('=== LISTING AVAILABLE SQUADRONS ===');

  try {
    clearCache();
    const squadrons = getSquadrons();
    const unitSquadrons = Object.values(squadrons)
      .filter(sq => sq.scope === 'UNIT')
      .sort((a, b) => a.unit - b.unit);

    console.log('\n' + '='.repeat(80));
    console.log('AVAILABLE SQUADRONS FOR TESTING');
    console.log('='.repeat(80));
    console.log(`\nTotal UNIT squadrons: ${unitSquadrons.length}\n`);

    if (showDetails) {
      console.log('Unit#  Charter          Name                                   OrgID');
      console.log('-'.repeat(80));
      unitSquadrons.forEach(sq => {
        const unitNum = String(sq.unit).padStart(3, '0');
        const charter = sq.charter.padEnd(15);
        const name = sq.name.substring(0, 35).padEnd(35);
        console.log(`${unitNum}    ${charter}  ${name}  ${sq.orgid}`);
      });
    } else {
      // Just show unit numbers in a compact format
      const unitNumbers = unitSquadrons.map(sq => String(sq.unit).padStart(3, '0'));
      console.log('Available unit numbers:');

      // Display in rows of 10
      for (let i = 0; i < unitNumbers.length; i += 10) {
        const row = unitNumbers.slice(i, i + 10);
        console.log('  ' + row.join(', '));
      }

      console.log('\nTo see full details, run: listAvailableSquadrons(true)');
    }

    console.log('\n' + '='.repeat(80));
    console.log(`\nTo test with a squadron, run: testUpdateSquadronGroups("${unitSquadrons[0].unit}")`);
    console.log('='.repeat(80) + '\n');

    Logger.info('=== SQUADRON LIST COMPLETED ===', {
      totalSquadrons: unitSquadrons.length
    });

    return unitSquadrons;

  } catch (err) {
    Logger.error('Failed to list squadrons', err);
    console.log(`\n✗ Error: ${err.message}`);
    throw err;
  }
}

/**
 * Test function to update groups for a specific squadron
 * Use this to test the system with a single squadron before going live
 *
 * First run listAvailableSquadrons() to see which squadrons exist
 *
 * @param {string} unitNumber - Unit number (e.g., "100", "205")
 * @returns {Object} Result object with details
 */
function testUpdateSquadronGroups(unitNumber = '346') {
  Logger.info('=== STARTING SINGLE SQUADRON TEST ===', { unitNumber: unitNumber });

  try {
    const result = updateSingleSquadronGroups(unitNumber);

    Logger.info('=== TEST COMPLETED SUCCESSFULLY ===', {
      squadron: result.squadron || 'Unknown',
      groupsCreated: result.created ? result.created.length : 0,
      groupsUpdated: result.updated ? result.updated.length : 0,
      errors: result.errors ? result.errors.length : 0
    });

    // Display results
    console.log('\n' + '='.repeat(80));
    console.log('TEST RESULTS FOR SQUADRON ' + unitNumber);
    console.log('='.repeat(80));

    if (result.created && result.created.length > 0) {
      console.log('\nGROUPS CREATED:');
      result.created.forEach(g => {
        console.log(`  ✓ ${g.groupEmail} (${g.memberCount} members)`);
      });
    }

    if (result.updated && result.updated.length > 0) {
      console.log('\nGROUPS UPDATED:');
      result.updated.forEach(g => {
        console.log(`  ↻ ${g.groupEmail} (${g.memberCount} members, +${g.changes?.added || 0}/-${g.changes?.removed || 0})`);
      });
    }

    if (result.errors && result.errors.length > 0) {
      console.log('\nERRORS:');
      result.errors.forEach(e => {
        console.log(`  ✗ ${e.groupEmail || e.suffix}: ${e.error}`);
      });
    }

    console.log('\n' + '='.repeat(80) + '\n');

    return result;

  } catch (err) {
    Logger.error('=== TEST FAILED ===', {
      unitNumber: unitNumber,
      errorMessage: err.message,
      stack: err.stack
    });
    throw err;
  }
}

/**
 * Test function to preview all squadron groups without making changes
 * Shows exactly what would be created for each squadron
 *
 * @param {boolean} detailedOutput - Whether to show detailed member lists
 * @returns {Object} Preview object
 */
function testPreviewSquadronGroups(detailedOutput = false) {
  Logger.info('=== STARTING SQUADRON GROUPS PREVIEW ===');

  try {
    const preview = previewSquadronGroups();

    console.log('\n' + '='.repeat(80));
    console.log('SQUADRON GROUPS PREVIEW - NO CHANGES MADE');
    console.log('='.repeat(80));
    console.log(`\nTotal Squadrons: ${preview.totalSquadrons}`);
    console.log(`Total Groups That Would Be Created/Updated: ${preview.totalSquadrons * 7}\n`);

    if (detailedOutput) {
      preview.squadrons.forEach(sq => {
        console.log(`\n${sq.charter} (${sq.unitPrefix}) - ${sq.totalMembers} members`);
        console.log('-'.repeat(60));
        sq.groups.forEach(group => {
          console.log(`  ${group.email}`);
          console.log(`    Type: ${group.type}`);
          console.log(`    Members: ${group.memberCount}`);
        });
      });
    } else {
      // Summary view
      const sampleSquadrons = preview.squadrons.slice(0, 5);
      console.log('Sample Squadrons (first 5):');
      sampleSquadrons.forEach(sq => {
        console.log(`\n  ${sq.charter}: ${sq.groups.length} groups, ${sq.totalMembers} members`);
      });
      console.log(`\n... and ${preview.totalSquadrons - 5} more squadrons`);
      console.log('\nRun testPreviewSquadronGroups(true) for detailed output');
    }

    console.log('\n' + '='.repeat(80) + '\n');

    Logger.info('=== PREVIEW COMPLETED ===', {
      totalSquadrons: preview.totalSquadrons
    });

    return preview;

  } catch (err) {
    Logger.error('=== PREVIEW FAILED ===', {
      errorMessage: err.message,
      stack: err.stack
    });
    throw err;
  }
}

/**
 * Test function to verify CAPWATCH data is loaded correctly
 * Checks that required files and data structures are present
 *
 * @returns {Object} Validation results
 */
function testCapwatchDataLoading() {
  Logger.info('=== TESTING CAPWATCH DATA LOADING ===');

  const results = {
    success: true,
    checks: [],
    errors: []
  };

  try {
    // Test 1: Load squadrons
    console.log('\nTest 1: Loading squadrons...');
    const squadrons = getSquadrons();
    const unitSquadrons = Object.values(squadrons).filter(sq => sq.scope === 'UNIT');
    results.checks.push({
      test: 'Load squadrons',
      passed: unitSquadrons.length > 0,
      details: `Found ${unitSquadrons.length} unit squadrons`
    });
    console.log(`  ✓ Found ${unitSquadrons.length} unit squadrons`);

    // Test 2: Load members
    console.log('\nTest 2: Loading members...');
    const members = getMembers();
    const memberCount = Object.keys(members).length;
    results.checks.push({
      test: 'Load members',
      passed: memberCount > 0,
      details: `Found ${memberCount} members`
    });
    console.log(`  ✓ Found ${memberCount} members`);

    // Test 3: Check squadron has members
    if (unitSquadrons.length > 0) {
      console.log('\nTest 3: Checking squadron membership...');
      const testSquadron = unitSquadrons[0];
      const squadronMembers = Object.values(members).filter(m => m.orgid === testSquadron.orgid);
      results.checks.push({
        test: 'Squadron has members',
        passed: squadronMembers.length > 0,
        details: `${testSquadron.charter} has ${squadronMembers.length} members`
      });
      console.log(`  ✓ ${testSquadron.charter} has ${squadronMembers.length} members`);
    }

    // Test 4: Check duty positions are loaded
    console.log('\nTest 4: Checking duty positions...');
    const membersWithDutyPositions = Object.values(members).filter(m =>
      m.dutyPositionIds && m.dutyPositionIds.length > 0
    );
    results.checks.push({
      test: 'Duty positions loaded',
      passed: membersWithDutyPositions.length > 0,
      details: `${membersWithDutyPositions.length} members have duty positions`
    });
    console.log(`  ✓ ${membersWithDutyPositions.length} members have duty positions`);

    // Test 5: Check email contacts
    console.log('\nTest 5: Checking email contacts...');
    const membersWithEmail = Object.values(members).filter(m => m.email);
    results.checks.push({
      test: 'Email contacts loaded',
      passed: membersWithEmail.length > 0,
      details: `${membersWithEmail.length} members have email addresses`
    });
    console.log(`  ✓ ${membersWithEmail.length} members have email addresses`);

    // Test 6: Check org contacts
    console.log('\nTest 6: Checking org contacts...');
    const orgContacts = parseFile('OrgContact');
    results.checks.push({
      test: 'Org contacts loaded',
      passed: orgContacts.length > 0,
      details: `Found ${orgContacts.length} org contact records`
    });
    console.log(`  ✓ Found ${orgContacts.length} org contact records`);

    // Test 7: Check parent contacts
    console.log('\nTest 7: Checking parent contacts...');
    const mbrContacts = parseFile('MbrContact');
    const parentContacts = mbrContacts.filter(c => c[1] === 'CADET PARENT EMAIL');
    results.checks.push({
      test: 'Parent contacts loaded',
      passed: parentContacts.length > 0,
      details: `Found ${parentContacts.length} parent contact records`
    });
    console.log(`  ✓ Found ${parentContacts.length} parent contact records`);

  } catch (err) {
    results.success = false;
    results.errors.push({
      error: err.message,
      stack: err.stack
    });
    Logger.error('Data loading test failed', err);
  }

  // Summary
  console.log('\n' + '='.repeat(80));
  console.log('DATA LOADING TEST SUMMARY');
  console.log('='.repeat(80));

  const passedTests = results.checks.filter(c => c.passed).length;
  const totalTests = results.checks.length;

  console.log(`\nTests Passed: ${passedTests}/${totalTests}`);

  if (results.errors.length > 0) {
    console.log('\nERRORS:');
    results.errors.forEach(e => {
      console.log(`  ✗ ${e.error}`);
    });
  }

  console.log('\n' + '='.repeat(80) + '\n');

  Logger.info('=== DATA LOADING TEST COMPLETED ===', {
    passed: passedTests,
    total: totalTests,
    success: results.success
  });

  return results;
}

/**
 * Test function to verify group settings are applied correctly
 * Creates a test group and checks its settings
 *
 * @returns {Object} Test results
 */
function testGroupSettings() {
  Logger.info('=== TESTING GROUP SETTINGS ===');

  const testGroupEmail = 'test-squadron-groups@miwg.cap.gov';

  try {
    console.log('\nCreating test group...');

    // Create test group
    const group = getOrCreateGroup(
      testGroupEmail,
      'Test Squadron Group',
      'This is a test group for verifying squadron group settings',
      {
        whoCanJoin: 'INVITED_CAN_JOIN',
        whoCanViewMembership: 'ALL_MEMBERS_CAN_VIEW',
        whoCanViewGroup: 'ALL_MEMBERS_CAN_VIEW',
        whoCanPostMessage: 'ANYONE_CAN_POST',
        allowExternalMembers: 'true',
        spamModerationLevel: 'MODERATE',
        enableCollaborativeInbox: 'true',
        includeInGlobalAddressList: 'true'
      }
    );

    console.log(`  ✓ Test group created: ${testGroupEmail}`);
    console.log(`  Group ID: ${group.id}`);
    console.log(`  Created: ${group.created ? 'Yes' : 'No (already existed)'}`);

    console.log('\n⚠ IMPORTANT: Please verify the following in Google Admin Console:');
    console.log('  1. Group exists in admin.google.com/ac/groups');
    console.log('  2. Collaborative inbox is enabled');
    console.log('  3. Group settings match configuration');
    console.log(`  4. Search for: ${testGroupEmail}`);

    console.log('\n⚠ Remember to delete the test group when done:');
    console.log(`  admin.google.com/ac/groups → Search "${testGroupEmail}" → Delete`);

    Logger.info('=== GROUP SETTINGS TEST COMPLETED ===', {
      testGroup: testGroupEmail,
      created: group.created
    });

    return {
      success: true,
      testGroup: testGroupEmail,
      groupId: group.id,
      created: group.created
    };

  } catch (err) {
    Logger.error('=== GROUP SETTINGS TEST FAILED ===', err);
    console.log(`\n✗ Test failed: ${err.message}`);
    throw err;
  }
}

/**
 * Test function to verify public contact membership calculation
 * Shows which members would be added to public contact groups
 *
 * @param {string} unitNumber - Unit number to test
 * @returns {Object} Test results
 */
function testPublicContactMembership(unitNumber = '346') {
  Logger.info('=== TESTING PUBLIC CONTACT MEMBERSHIP ===', { unitNumber: unitNumber });

  try {
    clearCache();
    const squadrons = getSquadrons();
    const members = getMembers();

    // Find the squadron
    const squadron = Object.values(squadrons).find(sq =>
      String(sq.unit) === String(unitNumber) && sq.scope === 'UNIT'
    );

    if (!squadron) {
      throw new Error(`Squadron ${unitNumber} not found`);
    }

    console.log('\n' + '='.repeat(80));
    console.log(`PUBLIC CONTACT MEMBERSHIP TEST - ${squadron.charter}`);
    console.log('='.repeat(80));

    // Get squadron members
    const squadronMembers = Object.values(members).filter(m => m.orgid === squadron.orgid);
    console.log(`\nTotal squadron members: ${squadronMembers.length}`);

    // Get public contact members
    const publicContactMembers = getPublicContactMembers(squadron, squadronMembers);

    console.log(`\nPublic contact members: ${Object.keys(publicContactMembers).length}`);
    console.log('\nMembership breakdown:');

    for (const email in publicContactMembers) {
      const member = publicContactMembers[email];
      console.log(`  ${member.role === 'OWNER' ? '👑' : '👤'} ${email}`);
      console.log(`     Role: ${member.role}`);
      console.log(`     Reason: ${member.reason}`);
    }

    // Check for recruiting mailbox
    const recruitingMailbox = SQUADRON_GROUP_CONFIG.PUBLIC_CONTACT.RECRUITING_MAILBOX;
    console.log(`\nWing recruiting mailbox: ${recruitingMailbox || 'NOT CONFIGURED'}`);

    // Show qualifying positions
    console.log(`\nQualifying duty positions:`);
    SQUADRON_GROUP_CONFIG.PUBLIC_CONTACT.DUTY_POSITIONS.forEach(pos => {
      console.log(`  • ${pos}`);
    });

    console.log('\n' + '='.repeat(80) + '\n');

    Logger.info('=== PUBLIC CONTACT MEMBERSHIP TEST COMPLETED ===', {
      squadron: squadron.charter,
      memberCount: Object.keys(publicContactMembers).length
    });

    return {
      squadron: squadron.charter,
      publicContactMembers: publicContactMembers,
      memberCount: Object.keys(publicContactMembers).length
    };

  } catch (err) {
    Logger.error('=== PUBLIC CONTACT MEMBERSHIP TEST FAILED ===', err);
    console.log(`\n✗ Test failed: ${err.message}`);
    throw err;
  }
}

/**
 * Test function to verify execution time tracking
 * Simulates processing and shows time remaining
 *
 * @returns {void}
 */
function testExecutionTimeTracking() {
  Logger.info('=== TESTING EXECUTION TIME TRACKING ===');

  const start = new Date();
  const maxExecutionTime = SQUADRON_GROUP_CONFIG.MAX_EXECUTION_TIME_MS || 400000;

  console.log('\n' + '='.repeat(80));
  console.log('EXECUTION TIME TRACKING TEST');
  console.log('='.repeat(80));
  console.log(`\nMax execution time: ${maxExecutionTime}ms (${maxExecutionTime/1000}s)`);
  console.log('Simulating squadron processing...\n');

  // Simulate processing squadrons
  for (let i = 1; i <= 10; i++) {
    const elapsed = new Date() - start;
    const remaining = maxExecutionTime - elapsed;
    const percentComplete = (elapsed / maxExecutionTime) * 100;

    console.log(`Squadron ${i}/10:`);
    console.log(`  Elapsed: ${elapsed}ms (${Math.round(percentComplete)}%)`);
    console.log(`  Remaining: ${remaining}ms`);

    if (elapsed > maxExecutionTime) {
      console.log(`  ⚠ Would timeout - stopping gracefully`);
      break;
    }

    // Simulate work
    Utilities.sleep(500);
  }

  const finalElapsed = new Date() - start;
  console.log(`\nTotal simulation time: ${finalElapsed}ms`);
  console.log(`Would have timed out: ${finalElapsed > maxExecutionTime ? 'YES' : 'NO'}`);

  console.log('\n' + '='.repeat(80) + '\n');

  Logger.info('=== EXECUTION TIME TRACKING TEST COMPLETED ===', {
    simulationTime: finalElapsed,
    maxTime: maxExecutionTime
  });
}

/**
 * Run all test functions in sequence
 * Comprehensive test suite for the entire system
 *
 * @returns {Object} All test results
 */
function runAllTests() {
  console.log('\n' + '='.repeat(80));
  console.log('RUNNING COMPLETE TEST SUITE');
  console.log('='.repeat(80) + '\n');

  const results = {
    startTime: new Date().toISOString(),
    tests: []
  };

  // Test 1: Data Loading
  console.log('\n### TEST 1: CAPWATCH Data Loading ###\n');
  try {
    const dataTest = testCapwatchDataLoading();
    results.tests.push({ name: 'Data Loading', success: dataTest.success, results: dataTest });
  } catch (err) {
    results.tests.push({ name: 'Data Loading', success: false, error: err.message });
  }

  // Find a valid squadron for remaining tests
  let testSquadronNumber = null;
  try {
    const squadrons = getSquadrons();
    const unitSquadrons = Object.values(squadrons).filter(sq => sq.scope === 'UNIT');
    if (unitSquadrons.length > 0) {
      testSquadronNumber = String(unitSquadrons[0].unit);
      console.log(`\nUsing squadron ${testSquadronNumber} for remaining tests...\n`);
    }
  } catch (err) {
    console.log('\n⚠ Warning: Could not find test squadron\n');
  }

  if (!testSquadronNumber) {
    console.log('\n✗ Cannot continue tests - no squadrons found\n');
    console.log('Please ensure CAPWATCH data is loaded and contains squadrons.\n');
    return results;
  }

  // Test 2: Public Contact Membership
  console.log(`\n### TEST 2: Public Contact Membership (Squadron ${testSquadronNumber}) ###\n`);
  try {
    const publicContactTest = testPublicContactMembership(testSquadronNumber);
    results.tests.push({ name: 'Public Contact Membership', success: true, results: publicContactTest });
  } catch (err) {
    results.tests.push({ name: 'Public Contact Membership', success: false, error: err.message });
  }

  // Test 3: Preview Groups
  console.log('\n### TEST 3: Preview Squadron Groups ###\n');
  try {
    const previewTest = testPreviewSquadronGroups(false);
    results.tests.push({ name: 'Preview Groups', success: true, results: previewTest });
  } catch (err) {
    results.tests.push({ name: 'Preview Groups', success: false, error: err.message });
  }

  // Test 4: Execution Time Tracking
  console.log('\n### TEST 4: Execution Time Tracking ###\n');
  try {
    testExecutionTimeTracking();
    results.tests.push({ name: 'Execution Time Tracking', success: true });
  } catch (err) {
    results.tests.push({ name: 'Execution Time Tracking', success: false, error: err.message });
  }

  // Summary
  console.log('\n' + '='.repeat(80));
  console.log('TEST SUITE SUMMARY');
  console.log('='.repeat(80));

  const passedTests = results.tests.filter(t => t.success).length;
  const totalTests = results.tests.length;

  console.log(`\nTests Passed: ${passedTests}/${totalTests}\n`);

  results.tests.forEach(test => {
    const status = test.success ? '✓' : '✗';
    console.log(`  ${status} ${test.name}`);
    if (!test.success && test.error) {
      console.log(`      Error: ${test.error}`);
    }
  });

  console.log('\n' + '='.repeat(80));
  console.log('\nNext Steps:');
  console.log(`  1. If all tests passed, try: testUpdateSquadronGroups("${testSquadronNumber}")`);
  console.log('  2. Verify the groups in Google Admin Console');
  console.log('  3. When ready, run: updateAllSquadronGroups()');
  console.log('='.repeat(80) + '\n');

  results.endTime = new Date().toISOString();

  return results;
}
