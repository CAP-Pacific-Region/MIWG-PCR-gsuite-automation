/*******************************************************
 * Group Membership Synchronization Module
 *
 * Version: 1.9.0
 * Filename: UpdateGroups.gs
 * Saved: 2026-07-27
 * Changes: 1.9.0: a new "Add Lite" column in the Groups sheet lets a row include
 *   cadet-lite members, addressed by their personal CAPWATCH address. Without it
 *   this path could not see them at all — getMembers() filtered them out before
 *   the desired set was built — so SquadronGroups added them to every unit .all
 *   and this path removed them again: 1,643 removals a night on the CAWG cadet
 *   tenant, undone an hour later, with members off their unit list in between.
 *   The wing-wide ca.all never had them at all, which is how this was found.
 *   "Add Lite" IMPLIES external members, because a row asking for accountless
 *   members while leaving "Add EXT" blank asks for two contradictory things and
 *   the losing side is silent. Cadet-lite members are now in the member set by
 *   default and still reach no group without the opt-in: they arrive with a null
 *   address and every group's `isMatch && .email` test skips them, exactly as
 *   when they were absent entirely.
 *   1.8.1: managed groups also get spamModerationLevel MODERATE. This
 *   path already enforced ANYONE_CAN_POST; the moderation setting is what makes
 *   that openness defensible, and SquadronGroups.gs 1.6.0 now manages the same
 *   pair, so the two paths no longer disagree about a group they both touch.
 *   1.8.0: updateEmailGroups() takes an optional deadline and resume
 *   position, and new updateEmailGroupsBatch() drives it in slices that fit
 *   inside the Apps Script execution limit. A run's length tracks the number of
 *   membership CHANGES, so the day a rule changes and thousands of memberships
 *   move with it, one pass is killed partway with no record of where it got to.
 *   State parks in Drive; checkEmailGroupsBatchStatus() reports it,
 *   resetEmailGroupsBatchProgress() discards it. No-argument behavior unchanged.
 *   1.7.0: 'professionalLevel' now places a member on their HIGHEST
 *   completed level only — a Level V holder is in all-level-v and nowhere else,
 *   where 1.6.0 put them in every level group they had ever passed. Adds
 *   'professionalLevelInProgress' for members holding some parts of a level but
 *   not all (Level 2 Part 1 without Part 2). The level ladder is derived from
 *   PL_Paths names, so a level re-split upstream needs no code change.
 *   1.6.0: new 'professionalLevel' attribute reads PD levels from the PL_*
 *   tables (PL_Paths, PL_MemberPathCredit, PL_Lookup), where the post-2018
 *   program actually records them. Achievements.txt still lists the RETIRED
 *   Level II-V (AchvIDs 131-134), so rows keyed on those resolved cleanly and
 *   matched nobody. Level 2 is two CAPWATCH paths and requires both.
 *   listProfessionalLevelPaths() prints the real path names.
 *   1.5.0: group-echelon DLs are no longer minted for positions that
 *   echelon does not establish. Duty and rank rules seed one DL per group-echelon
 *   org before matching members, so a wing-only office — Director of Information
 *   Technology (an IT Officer below wing), Inspector General (groups have none) —
 *   produced one permanently empty DL per group, every run. An empty seed is now
 *   kept only when the group already exists, which is the vacant-seat case that
 *   seeding is for; otherwise it is dropped. cleanupEmptyEchelonGroups() removes
 *   the ones already created.
 *   1.4.0: three group-membership fixes, all of which failed silently.
 *   - dutyPositionIdsWingHQ matched the MEMBER'S HOME UNIT instead of the org the
 *     duty is held at, so a wing office DL held only those wing staff who are also
 *     Wing HQ members — typically the assistant, not the primary.
 *   - Duty titles are now compared after the CAPR 30-1 renames
 *     (DUTY_TITLE_OVERRIDES), so a row keyed on a current title also matches
 *     eServices rows still carrying the retired one.
 *   - achievements Values may be achievement NAMES, not just numeric AchvIDs;
 *     MbrAchievements only ever carries the ID, so name rows (the Education &
 *     Training levels) matched nothing and produced empty groups.
 *   An unknown Attribute no longer leaves an empty group behind, and
 *   previewEmailGroupRows() / listAchievementNames() show what a row resolves to
 *   without changing anything. See PCR_CHANGELOG.md.
 *   1.3.9: Managed wing-scope display name returns CONFIG.WING_ABBREVIATION
 *   instead of literal 'CAWG' (inside the existing WING==='CA' display branch —
 *   no behavior change for CA). See PCR_CHANGELOG.md.
 *   1.3.8: AdminDirectory.Users.list standardized to customer:"my_customer".
 *
 * Manages Google Groups memberships based on CAPWATCH data and configuration:
 * - Reads group configuration from automation spreadsheet
 * - Builds group memberships based on member attributes (type, rank, duty positions, committee assignments, etc.)
 * - Creates both wing-level and group-level groups automatically
 * - Calculates membership deltas (add/remove changes)
 * - Applies changes to Google Workspace groups
 * - Auto-creates groups that don't exist
 * - Supports manual member additions via spreadsheet
 * - Tracks and logs errors to spreadsheet for review
 * - Fixed external contacts and parent/guardian emails
 * - FIxed Group name and description
 * - Added new column to Groups tab EXT where you can create a Group as allow external members
 * - Added 'dutyPositionIdsGroupScope' attribute: one wing-level DL of the
 *   group-echelon-unit commanders only (e.g. ca.group-commanders)
 *******************************************************/

var workspaceUsers = {};
var workspaceEmailMap = {};

// Desired metadata (name/description) computed from the Groups sheet for logging and group creation
var desiredGroupMeta = {};
// Map base group name -> Attribute from the Groups sheet (built in getEmailGroupDeltas)
var groupAttributeByName = {};
// Map base group name -> allow external members boolean (from Groups sheet Add EXT column)
var groupAllowExternalByName = {};
// Map base group name -> include cadet-lite boolean (from Groups sheet "Add Lite" column)
var groupIncludeCadetLiteByName = {};

const DRY_RUN = false; // change to false for real updates

/**
 * Applies the Groups sheet to Google Groups.
 *
 * Called with no arguments this behaves as it always has: compute the deltas, then
 * apply every one of them in a single execution.
 *
 * `options` makes the same pass **interruptible**, which is what updateEmailGroupsBatch()
 * uses. A membership change costs an API call plus its pacing delay, so a run that
 * moves thousands of members — the day a rule changes, for instance — can outlast the
 * Apps Script execution limit. Rather than a second copy of the add/remove logic, the
 * one loop learned to stop on a deadline and say exactly where it stopped.
 *
 * @param {Object} [options]
 * @param {number} [options.deadlineMs] - Epoch ms to stop by; 0/absent = run to completion
 * @param {Object} [options.resume] - Saved state from a previous paused run:
 *   { deltas, groupIndex, memberIndex, errorEmails, totals }
 * @returns {{complete: boolean, groupIndex: number, memberIndex: number, totalGroups: number,
 *   errorEmails: Object, totals: {added: number, removed: number, errors: number}}}
 */
function updateEmailGroups(options) {
  const opts = options || {};
  const deadlineMs = Number(opts.deadlineMs || 0);
  const resume = opts.resume || null;

  const start = new Date();
  if (!resume) clearCache();

  let deltas = resume ? resume.deltas : getEmailGroupDeltas();
  let errorEmails = (resume && resume.errorEmails) || {};
  let totalAdded = (resume && resume.totals && resume.totals.added) || 0;
  let totalRemoved = (resume && resume.totals && resume.totals.removed) || 0;
  let totalErrors = (resume && resume.totals && resume.totals.errors) || 0;

  // One flat, ordered list so a position in the run is a single number. Object key
  // order is insertion order here, and the deltas are rebuilt from the same sheet in
  // the same order, so the list is stable across executions.
  const work = [];
  for (const category in deltas) {
    for (const group in deltas[category]) work.push([category, group]);
  }

  let groupIndex = (resume && Number(resume.groupIndex)) || 0;
  let memberIndex = (resume && Number(resume.memberIndex)) || 0;
  let processedCategories = groupIndex;
  const totalCategories = work.length;
  const outOfTime = () => deadlineMs > 0 && Date.now() >= deadlineMs;

  // For dry-run summary
  let dryRunSummary = [];

  for (; groupIndex < work.length; groupIndex++) {
    if (outOfTime()) break;

    const category = work[groupIndex][0];
    const group = work[groupIndex][1];
    processedCategories++;
    {
      let added = 0;
      let removed = 0;
      const groupEmail = group + CONFIG.EMAIL_DOMAIN;
      const baseGroupName = group.includes('.') ? group.split('.').slice(1).join('.') : group;

      // Metadata and settings are checked once per group, and skipped when resuming
      // into the middle of one — they were applied before the pause.
      if (memberIndex === 0) {
        const metaForGroup = desiredGroupMeta[groupEmail.toLowerCase()] || {};
        applyGroupMeta_(groupEmail, metaForGroup);
        applyManagedGroupSettings_(groupEmail, {
          allowExternalMembers: !!groupAllowExternalByName[baseGroupName],
          whoCanViewMembership: 'ALL_IN_DOMAIN_CAN_VIEW',
          whoCanPostMessage: 'ANYONE_CAN_POST',
          spamModerationLevel: 'MODERATE'
        });
      }

      let dryRunMembers = [];

      const memberEmails = Object.keys(deltas[category][group]);
      let pausedInGroup = false;

      for (; memberIndex < memberEmails.length; memberIndex++) {
        if (outOfTime()) { pausedInGroup = true; break; }
        const email = memberEmails[memberIndex];

        switch(deltas[category][group][email]) {
          case -1:
            // Remove member
            try {
              const finalEmail = workspaceEmailMap[email.replace(/\D/g, '')] || email;
              if (DRY_RUN) {
                Logger.info('💡 [Dry-Run] Would remove member', {
                  member: email,
                  group: groupEmail
                });
                dryRunMembers.push({ email: finalEmail, action: 'REMOVE' });
              } else {
                executeWithRetry(() =>
                  AdminDirectory.Members.remove(groupEmail, finalEmail)
                );
                removed++;
                  Logger.info('Removed member from group', {
                    member: email,
                    group: groupEmail
                  });
              }
            } catch (e) {
              Logger.error('Failed to remove member from group', {
                member: email,
                group: groupEmail,
                category: category,
                errorMessage: e.message,
                errorCode: e.details?.code,
                errorReason: e.details?.errors?.[0]?.reason
              });

              // Track removal errors too
              if (!errorEmails[email]) {
                errorEmails[email] = {
                  email: email,
                  attempts: [],
                  firstSeen: new Date().toISOString()
                };
              }
              errorEmails[email].attempts.push({
                group: group,
                groupEmail: groupEmail,
                category: category,
                action: 'REMOVE',
                errorCode: e.details?.code || 'Unknown',
                errorMessage: e.message || 'Unknown error',
                timestamp: new Date().toISOString()
              });

              totalErrors++;
            }
            break;
          case 1:
            // Add member
            try {
              const finalEmail = workspaceEmailMap[email.replace(/\D/g, '')] || email;

              // Skip external emails except for groups whose Attribute is 'contact'
              if (!finalEmail.endsWith(CONFIG.EMAIL_DOMAIN) && groupAttributeByName[category] !== 'contact') continue;

              if (DRY_RUN) {
                Logger.info('💡 [Dry-Run] Would add member', {
                  member: finalEmail,
                  group: groupEmail
                });
                dryRunMembers.push({ email: finalEmail, action: 'ADD' });
                // Continue to next member, do not actually add
                continue;
              }

              executeWithRetry(() =>
                AdminDirectory.Members.insert({
                  email: finalEmail,
                  role: 'MEMBER'
                }, groupEmail)
              );
              added++;
                Logger.info('Added member to group', {
                  member: email,
                  group: groupEmail
                });

                // Throttle between API insert calls
                Utilities.sleep(CONFIG.API_DELAY_MS);

                // Periodic quota cooldown
                if (added > 0 && added % 25 === 0) {
                  Logger.info('Pausing briefly to allow API quota refill', { added });
                  Utilities.sleep(15000); // 15 sec every 25 adds
                }
            } catch (e) {
              // Check if member is already in group (409 = Conflict/Duplicate)
              if (e.details?.code === 409) {
                Logger.warn('Member already in group', {
                  member: email,
                  group: groupEmail,
                  category: category
                });
              }
              else if (e.details?.code === 404) {
                Logger.warn('Cannot add external member - not found', {
                  member: email,
                  group: groupEmail,
                  category: category,
                  note: 'Email may not exist or group settings prevent external members'
                });

                // Track detailed error info
                if (!errorEmails[email]) {
                  errorEmails[email] = {
                    email: email,
                    attempts: [],
                    firstSeen: new Date().toISOString()
                  };
                }
                errorEmails[email].attempts.push({
                  group: group,
                  groupEmail: groupEmail,
                  category: category,
                  errorCode: 404,
                  errorMessage: 'Resource Not Found',
                  timestamp: new Date().toISOString()
                });
              }
              // All other errors
              else {
                Logger.error('Failed to add member to group', {
                  member: email,
                  group: groupEmail,
                  category: category,
                  errorMessage: e.message,
                  errorCode: e.details?.code,
                  errorReason: e.details?.errors?.[0]?.reason,
                  fullError: JSON.stringify(e.details)
                });

                // Track detailed error info
                if (!errorEmails[email]) {
                  errorEmails[email] = {
                    email: email,
                    attempts: [],
                    firstSeen: new Date().toISOString()
                  };
                }
                errorEmails[email].attempts.push({
                  group: group,
                  groupEmail: groupEmail,
                  category: category,
                  errorCode: e.details?.code || 'Unknown',
                  errorMessage: e.message || 'Unknown error',
                  errorReason: e.details?.errors?.[0]?.reason || 'Unknown',
                  timestamp: new Date().toISOString()
                });
              }

              totalErrors++;
            }
            break;
          case 0:
            // Member already in group
            break;
        }
      }

      totalAdded += added;
      totalRemoved += removed;

      if (DRY_RUN && dryRunMembers.length > 0) {
        dryRunSummary.push({
          group: groupEmail,
          category: category,
          members: dryRunMembers
        });

        try {
          let folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
          let dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
          let safeGroup = group.replace(/[^a-zA-Z0-9.-]/g, '_');
          let fileName = `DryRun-${safeGroup}-${dateStr}.csv`;

          // Use same columns as members_template.csv
          let csvHeader = 'Group Email [Required],Member Email,Member Type,Member Role\n';
          let csvContent = csvHeader;

          dryRunMembers.forEach(m => {
            let memberType = m.action === 'ADD' ? 'User' : 'Removed';
            let memberRole = 'MEMBER';
            csvContent += `${groupEmail},${m.email},${memberType},${memberRole}\n`;
          });

          let file = folder.createFile(fileName, csvContent, MimeType.CSV);
          Logger.info('💡 [Dry-Run] Group CSV saved', {
            fileName: fileName,
            url: file.getUrl(),
            memberCount: dryRunMembers.length
          });
        } catch (e) {
          Logger.error('💡 [Dry-Run] Failed to save CSV for group', {
            group: groupEmail,
            error: e.message
          });
        }
      }

      const meta = desiredGroupMeta[groupEmail.toLowerCase()] || {};
      Logger.info(pausedInGroup ? 'Paused partway through group' : 'Updated group', {
        groupId: group,
        group: groupEmail,
        name: meta.name || '',
        description: meta.description || '',
        added: added,
        removed: removed,
        membersDone: pausedInGroup ? `${memberIndex}/${memberEmails.length}` : undefined
      });

      if (pausedInGroup) break;   // leave groupIndex here; memberIndex says where to resume
      memberIndex = 0;            // next group starts at its first member
    }
    if (processedCategories % 5 === 0 || processedCategories === totalCategories) {
      Logger.info('Progress update', {
        processed: processedCategories,
        total: totalCategories,
        percentComplete: Math.round((processedCategories / totalCategories) * 100)
      });
    }
  }

  const complete = groupIndex >= work.length;

  // A paused run keeps its errors in the returned state and writes the sheet once,
  // at the end — saveErrorEmails() clears and rewrites the tab, so calling it per
  // execution would leave only the last batch's errors on it.
  if (complete) saveErrorEmails(errorEmails);

  // Dry-run: Save summary file and log
  if (DRY_RUN) {
    try {
      let folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
      let dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
      let fileName = 'DryRun-Groups-' + dateStr + '.json';
      let content = JSON.stringify(dryRunSummary, null, 2);
      let file = folder.createFile(fileName, content, MimeType.PLAIN_TEXT);
      Logger.info('💡 [Dry-Run] Summary saved', {
        url: file.getUrl(),
        groupCount: dryRunSummary.length
      });
    } catch (e) {
      Logger.error('💡 [Dry-Run] Failed to save summary', {
        error: e.message
      });
    }
  }

  Logger.info(complete ? 'Email group update completed' : 'Email group update paused on deadline', {
    duration: new Date() - start + 'ms',
    groupsDone: `${groupIndex}/${work.length}`,
    totalAdded: totalAdded,
    totalRemoved: totalRemoved,
    totalErrors: totalErrors,
    errorEmailsCount: Object.keys(errorEmails).length
  });

  return {
    complete: complete,
    groupIndex: groupIndex,
    memberIndex: memberIndex,
    totalGroups: work.length,
    errorEmails: errorEmails,
    totals: { added: totalAdded, removed: totalRemoved, errors: totalErrors },
    // Handed back so a paused run can be parked without recomputing it.
    deltas: deltas
  };
}

const EMAIL_GROUPS_BATCH_STATE_FILE_ = 'EmailGroupsBatchState.json';

/**
 * Wall-clock budget for one slice, in minutes.
 *
 * These tenants allow a 30-minute execution, so 25 leaves five minutes of headroom
 * for the final state save — which writes the whole delta set to Drive and is the
 * one step that must not be interrupted. A tenant on the 6-minute tier should pass
 * 5 explicitly.
 */
const EMAIL_GROUPS_BATCH_DEFAULT_BUDGET_MIN_ = 25;
const EMAIL_GROUPS_BATCH_STALE_HOURS_ = 12;

/**
 * updateEmailGroups() in slices, for when one pass cannot finish inside the Apps
 * Script execution limit.
 *
 * Every membership change is an API call with pacing behind it, so the run's length
 * tracks the number of CHANGES, not the number of groups. A steady day is minutes;
 * the day a rule changes and thousands of memberships move with it is not, and that
 * run gets killed partway with no record of where it got to.
 *
 * The first call computes the deltas and parks them in Drive alongside the group
 * metadata; each call then applies as much as fits in its time budget and saves its
 * position. Run it again — by hand, or by pointing the daily trigger at it — until
 * it reports complete. Re-applying a slice is harmless anyway: an add that already
 * happened comes back 409 and a removal 404, both of which are caught.
 *
 *   updateEmailGroupsBatch()        // 25 minutes, sized for this tenant's 30-minute limit
 *   updateEmailGroupsBatch(5)       // a shorter slice, e.g. on the 6-minute tier
 *   checkEmailGroupsBatchStatus()   // how far along, without touching anything
 *   resetEmailGroupsBatchProgress() // discard the parked run and start fresh
 *
 * @param {number} [budgetMinutes=25] - Wall-clock budget for THIS execution
 * @returns {{complete: boolean, groupIndex: number, totalGroups: number}}
 */
function updateEmailGroupsBatch(budgetMinutes) {
  const budgetMs = Math.max(1, Number(budgetMinutes || EMAIL_GROUPS_BATCH_DEFAULT_BUDGET_MIN_)) * 60 * 1000;
  const deadlineMs = Date.now() + budgetMs;

  const saved = loadEmailGroupsBatchState_();
  let resume = null;

  if (saved) {
    // Restore the maps the apply pass reads out of module scope.
    desiredGroupMeta = saved.desiredGroupMeta || {};
    groupAllowExternalByName = saved.groupAllowExternalByName || {};
    groupAttributeByName = saved.groupAttributeByName || {};
    workspaceEmailMap = saved.workspaceEmailMap || {};

    resume = {
      deltas: saved.deltas,
      groupIndex: saved.groupIndex,
      memberIndex: saved.memberIndex,
      errorEmails: saved.errorEmails || {},
      totals: saved.totals
    };

    Logger.info('Resuming parked email group run', {
      startedAt: saved.startedAt,
      resumingAt: `${saved.groupIndex}/${saved.totalGroups}`,
      memberIndex: saved.memberIndex
    });
  }

  const result = updateEmailGroups({ deadlineMs: deadlineMs, resume: resume });

  if (result.complete) {
    clearEmailGroupsBatchState_();
    Logger.info('Email group batch finished', {
      groups: result.totalGroups,
      added: result.totals.added,
      removed: result.totals.removed,
      errors: result.totals.errors
    });
    console.log(`✅ Complete — ${result.totalGroups} groups, ` +
      `+${result.totals.added} / -${result.totals.removed}, ${result.totals.errors} errors.`);
  } else {
    saveEmailGroupsBatchState_({
      startedAt: (saved && saved.startedAt) || new Date().toISOString(),
      savedAt: new Date().toISOString(),
      groupIndex: result.groupIndex,
      memberIndex: result.memberIndex,
      totalGroups: result.totalGroups,
      totals: result.totals,
      errorEmails: result.errorEmails,
      deltas: result.deltas,
      desiredGroupMeta: desiredGroupMeta,
      groupAllowExternalByName: groupAllowExternalByName,
      groupAttributeByName: groupAttributeByName,
      workspaceEmailMap: workspaceEmailMap
    });
    console.log(`⏸ Paused at group ${result.groupIndex}/${result.totalGroups}. ` +
      `Run updateEmailGroupsBatch() again to continue.`);
  }

  return { complete: result.complete, groupIndex: result.groupIndex, totalGroups: result.totalGroups };
}

/**
 * Read-only: how far the parked run got.
 * @returns {void}
 */
function checkEmailGroupsBatchStatus() {
  const saved = loadEmailGroupsBatchState_();
  if (!saved) {
    console.log('No parked email group run. The next updateEmailGroupsBatch() starts fresh.');
    return;
  }
  console.log(`Parked run started ${saved.startedAt}, last saved ${saved.savedAt}`);
  console.log(`  position: group ${saved.groupIndex}/${saved.totalGroups}, member ${saved.memberIndex}`);
  console.log(`  so far:   +${saved.totals.added} / -${saved.totals.removed}, ${saved.totals.errors} errors`);
}

/**
 * Discards a parked run so the next batch recomputes from the sheet. Changes nothing
 * in Workspace — the memberships already applied stay applied.
 * @returns {void}
 */
function resetEmailGroupsBatchProgress() {
  clearEmailGroupsBatchState_();
  console.log('Parked email group run discarded. The next updateEmailGroupsBatch() starts fresh.');
}

/**
 * @returns {Object|null} Parked state, or null when there is none or it has gone stale
 */
function loadEmailGroupsBatchState_() {
  try {
    const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
    const files = folder.getFilesByName(EMAIL_GROUPS_BATCH_STATE_FILE_);
    if (!files.hasNext()) return null;

    const state = JSON.parse(files.next().getBlob().getDataAsString());
    if (!state || !state.deltas) return null;

    // Deltas were computed against a directory that has since moved on. Past a point
    // it is more honest to recompute than to apply yesterday's answer.
    const ageHours = (Date.now() - new Date(state.savedAt).getTime()) / 3600000;
    if (ageHours > EMAIL_GROUPS_BATCH_STALE_HOURS_) {
      Logger.warn('Parked email group run is stale; recomputing instead of resuming', {
        savedAt: state.savedAt,
        ageHours: Math.round(ageHours)
      });
      clearEmailGroupsBatchState_();
      return null;
    }

    return state;
  } catch (e) {
    Logger.warn('Could not read parked email group run; starting fresh', { errorMessage: e.message });
    return null;
  }
}

/**
 * @param {Object} state
 * @returns {void}
 */
function saveEmailGroupsBatchState_(state) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
    const content = JSON.stringify(state);
    const files = folder.getFilesByName(EMAIL_GROUPS_BATCH_STATE_FILE_);

    if (files.hasNext()) files.next().setContent(content);
    else folder.createFile(EMAIL_GROUPS_BATCH_STATE_FILE_, content, MimeType.PLAIN_TEXT);

    Logger.info('Parked email group run saved', {
      position: `${state.groupIndex}/${state.totalGroups}`,
      memberIndex: state.memberIndex,
      bytes: content.length
    });
  } catch (e) {
    Logger.error('Failed to park the email group run - the next call will start over', {
      errorMessage: e.message
    });
  }
}

/**
 * @returns {void}
 */
function clearEmailGroupsBatchState_() {
  try {
    const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
    const files = folder.getFilesByName(EMAIL_GROUPS_BATCH_STATE_FILE_);
    while (files.hasNext()) files.next().setTrashed(true);
  } catch (e) {
    Logger.warn('Could not clear the parked email group run', { errorMessage: e.message });
  }
}

/**
 * Loads CAPWATCH committee membership and attaches it to the in memory `members` map.
 *
 * Adds:
 *   members[capid].committeeAssignments = [{ orgid: ..., name: ... }, ...]
 *
 * Only WING and GROUP scoped committees are attached (no squadron/unit committees).
 *
 * Source:
 *   parseFile('MbrCommittee') rows shaped like: [CAPID, Committee, Chair, ORGID, DateAssigned]
 *
 * @param {Object<string, Object>} members - Members object indexed by CAPID
 * @param {Object<string, Object>} squadrons - Squadrons object indexed by orgid
 * @returns {void}
 */
function attachCommitteesToMembers(members, squadrons) {
  let rows = [];
  try {
    rows = parseFile('MbrCommittee');
  } catch (e) {
    Logger.warn('MbrCommittee not available; committeeAssignments will be empty', { error: e.message });
    return;
  }

  // Initialize committeeAssignments arrays for every member we know about
  for (const capid in members) {
    if (!Array.isArray(members[capid].committeeAssignments)) {
      members[capid].committeeAssignments = [];
    }
  }

  // CAPWATCH MbrCommittee columns:
  // [0]=CAPID, [1]=Committee, [2]=Chair, [3]=ORGID, [4]=DateAssigned
  let attached = 0;
  let skippedNoMember = 0;
  let skippedMissing = 0;
  let skippedNonWingGroup = 0;

  for (let i = 0; i < rows.length; i++) {
    const capid = (rows[i][0] || '').toString().trim();
    const committeeName = (rows[i][1] || '').toString().trim();
    const orgid = (rows[i][3] || '').toString().trim();

    if (!capid || !committeeName || !orgid) {
      skippedMissing++;
      continue;
    }

    if (!members[capid]) {
      skippedNoMember++;
      continue;
    }

    const org = squadrons && squadrons[orgid] ? squadrons[orgid] : null;
    if (!org || (org.scope !== 'WING' && org.scope !== 'GROUP')) {
      // Explicit requirement: ONLY wing + group committees, no squadrons/units
      skippedNonWingGroup++;
      continue;
    }

    const assigns = members[capid].committeeAssignments;
    const exists = assigns.some(a => a && a.orgid === orgid && a.name === committeeName);
    if (!exists) {
      assigns.push({ orgid: orgid, name: committeeName });
      attached++;
    }
  }

  Logger.info('Attached committeeAssignments to members (WING+GROUP only)', {
    membersCount: Object.keys(members).length,
    rows: rows.length,
    attached: attached,
    skippedNoMember: skippedNoMember,
    skippedMissing: skippedMissing,
    skippedNonWingGroup: skippedNonWingGroup
  });
}

/**
 * Converts a name to title style capitalization.
 * - Converts ALL CAPS words to Title Case (e.g., "CALIFORNIA" -> "California")
 * - Preserves known acronyms (e.g., CAP, FAA, USAF) and WG-style acronyms (e.g., CAWG, HIWG)
 * - Preserves tokens with digits
 * @param {string} s
 * @returns {string}
 */
function toSentenceCase_(s) {
  const str = (s || '').toString().trim();
  if (!str) return '';

  const preserve = new Set([
    'CAP', 'USAF', 'FAA', 'DOT', 'TSA', 'ICAO', 'EASA', 'HQ'
  ]);

  function isWingAcronym_(tok) {
    // CAWG, HIWG, ORWG, NVWG, PCR, etc.
    return /^[A-Z]{2,4}WG$/.test(tok) || tok === 'PCR';
  }

  function titleToken_(tok) {
    if (!tok) return tok;
    if (/\d/.test(tok)) return tok;

    // Strip trailing punctuation for evaluation, re-attach later
    const m = tok.match(/^(.+?)([.,;:)]?)$/);
    const core = m ? m[1] : tok;
    const punct = m ? m[2] : '';

    const upper = core.toUpperCase();

    // Preserve known acronyms / wing acronyms
    if (preserve.has(upper) || isWingAcronym_(upper)) return upper + punct;

    // If it's ALL CAPS, convert to Title Case
    if (core === upper) {
      return (upper.charAt(0) + upper.slice(1).toLowerCase()) + punct;
    }

    // Otherwise just capitalize first letter and lower the rest
    return (core.charAt(0).toUpperCase() + core.slice(1).toLowerCase()) + punct;
  }

  return str
    .split(/\s+/)
    .map(titleToken_)
    .join(' ');
}

function isCAWGTenant_() {
  return String((CONFIG && CONFIG.WING) || '').trim().toUpperCase() === 'CA';
}

function stripLeadingHonorific_(value) {
  return String(value || '')
    .replace(/^(lt\.?\s*col|col|maj|capt|1st\s*lt|2nd\s*lt|lt)\s+/i, '')
    .trim();
}

function abbreviateManagedGroupCommonShortName_(value) {
  const normalized = String(value || '').trim();
  if (!normalized) return '';

  if (/^los angeles county$/i.test(normalized)) return 'LA County';
  if (/^san francisco bay$/i.test(normalized)) return 'SF Bay';

  return normalized;
}

function abbreviateManagedGroupOrgDisplayName_(org) {
  if (!org || !org.name) return '';

  const fullName = toSentenceCase_(String(org.name || '').trim());
  if (!fullName) return '';

  if (!isCAWGTenant_()) return fullName;

  const scope = String(org.scope || '').trim().toUpperCase();
  const unit = String(org.unit || '').trim().replace(/^0+/, '');

  if (scope === 'WING') {
    return CONFIG.WING_ABBREVIATION;
  }

  if (scope === 'GROUP') {
    const match = fullName.match(/^(.*)\bGroup\s+(\d+)\b$/i);
    if (match) {
      const shortName = abbreviateManagedGroupCommonShortName_(String(match[1] || '').trim());
      const number = String(match[2] || '').trim();
      return shortName ? `Grp ${number} ${shortName}` : `Grp ${number}`;
    }
    return unit ? `Grp ${unit} ${fullName}` : fullName;
  }

  const leadingNumberUnitMatch = fullName.match(/^(\d+(?:st|nd|rd|th))\s+(.*?)(?:\s+(?:Challenger\s+)?)?(?:Cadet|Composite)?\s*Sq(?:dn|uadron)?\b$/i);
  if (leadingNumberUnitMatch) {
    const number = String(leadingNumberUnitMatch[1] || '').trim().toUpperCase();
    const shortName = abbreviateManagedGroupCommonShortName_(
      stripLeadingHonorific_(String(leadingNumberUnitMatch[2] || '').trim())
    );
    return shortName ? `Sqdn ${number} ${shortName}` : `Sqdn ${number}`;
  }

  const unitMatch = fullName.match(/^(.*?)(?:\s+(?:Challenger\s+)?)?(?:Cadet|Composite)?\s*Sq(?:dn|uadron)?\s+(\d+)\b$/i)
    || fullName.match(/^(.*?)\s+Squadron\s+(\d+)\b$/i);
  if (unitMatch) {
    const shortName = abbreviateManagedGroupCommonShortName_(
      stripLeadingHonorific_(String(unitMatch[1] || '').trim())
    );
    const number = String(unitMatch[2] || '').trim();
    return shortName ? `Sqdn ${number} ${shortName}` : `Sqdn ${number}`;
  }

  return unit ? `Sqdn ${unit} ${fullName}` : fullName;
}

/**
 * Ensures an existing Google Group has the desired name/description.
 * Uses PATCH so only provided fields are updated.
 * Dry-run aware.
 * @param {string} groupEmail Full group email address
 * @param {{name?: string, description?: string}} meta Desired metadata
 * @returns {void}
 */
function applyGroupMeta_(groupEmail, meta) {
  const email = (groupEmail || '').toString().toLowerCase();
  if (!email || !meta) return;

  const desiredName = (meta.name || '').toString().trim();
  const desiredDesc = (meta.description || '').toString().trim();
  if (!desiredName && !desiredDesc) return;

  try {
    const existing = AdminDirectory.Groups.get(email);
    const currentName = (existing.name || '').toString().trim();
    const currentDesc = (existing.description || '').toString().trim();

    const patch = {};
    if (desiredName && desiredName !== currentName) patch.name = desiredName;
    if (desiredDesc && desiredDesc !== currentDesc) patch.description = desiredDesc;

    if (Object.keys(patch).length === 0) return;

    if (DRY_RUN) {
      Logger.info('💡 [Dry-Run] Would update group metadata', {
        group: email,
        fromName: currentName,
        toName: desiredName,
        fromDescription: currentDesc,
        toDescription: desiredDesc
      });
      return;
    }

    executeWithRetry(() => AdminDirectory.Groups.patch(patch, email));
    Logger.info('Updated group metadata', { group: email, ...patch });
  } catch (e) {
    Logger.warn('Failed to update group metadata', {
      group: email,
      errorMessage: e.message,
      errorCode: e.details?.code
    });
  }
}

/**
 * Calculates email group membership deltas by comparing desired state with current state
 * Returns object with delta values: 1 = add, 0 = no change, -1 = remove
 * @returns {Object} Groups object with delta values for each member
 */
function getEmailGroupDeltas() {
  const start = new Date();
  let groups = {};
  let groupsConfig = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID).getSheetByName('Groups').getDataRange().getValues();

// Build Group Name -> Attribute lookup for use during updateEmailGroups filtering
groupAttributeByName = {};
groupAllowExternalByName = {};
// Not carried in the batch state: it is only read while deltas are computed, and
// a resumed slice applies deltas it already has.
groupIncludeCadetLiteByName = {};

const groupsHeader = (groupsConfig[0] || []).map(h => (h || '').toString().trim().toLowerCase());
const addExtIdx = groupsHeader.indexOf('add ext');
const addLiteIdx = groupsHeader.indexOf('add lite');

function isTruthyAddExt_(v) {
  const t = (v || '').toString().trim().toLowerCase();
  return t === 'y' || t === 'yes' || t === 'x' || t === 'true';
}

for (let r = 1; r < groupsConfig.length; r++) {
  const gName = (groupsConfig[r][1] || '').toString().trim(); // Group Name
  const attr = (groupsConfig[r][2] || '').toString().trim();  // Attribute
  if (!gName) continue;
  groupAttributeByName[gName] = attr;

  const wantsLite = addLiteIdx > -1 ? isTruthyAddExt_(groupsConfig[r][addLiteIdx]) : false;
  const wantsExt = addExtIdx > -1 ? isTruthyAddExt_(groupsConfig[r][addExtIdx]) : false;

  // "Add Lite" IMPLIES external members. A cadet-lite member is addressed by a
  // personal CAPWATCH address, so a row that asks for them and leaves "Add EXT"
  // blank is asking for two things that contradict each other — and the losing
  // side is silent: the adds fail, or worse, allowExternalMembers is set false
  // here while another writer sets it true, and the flag flips daily.
  groupAllowExternalByName[gName] = wantsExt || wantsLite;
  groupIncludeCadetLiteByName[gName] = wantsLite;
}

  // Map base group name -> spreadsheet description (if provided)
  // Expected Groups sheet columns:
  // [0]=Category, [1]=Group Name, [2]=Attribute, [3]=Values, [4]=Description
  const descriptionByGroupName = {};
  for (let r = 1; r < groupsConfig.length; r++) {
    const gName = (groupsConfig[r][1] || '').toString().trim();
    const desc = (groupsConfig[r][4] || '').toString().trim();
    if (gName) {
      descriptionByGroupName[gName] = desc;
    }
  }
  let squadrons = getSquadrons();

  // Cadet-lite members are INCLUDED here and still reach no group by default.
  // createMemberObject sets email:null, and only buildWorkspaceEmailMapForGroups_
  // fills it — from the Workspace directory, where an accountless member has no
  // entry. So they arrive addressless and every group's `isMatch && .email` test
  // skips them, exactly as when they were filtered out entirely. What changes is
  // that a row opting into "Add Lite" can now give them an address.
  let members = getMembers(CONFIG.MEMBER_TYPES.ACTIVE, true, true);
  // --- Build CAPWATCH → Workspace email map ---
  attachCommitteesToMembers(members, squadrons);
  buildWorkspaceEmailMapForGroups_(members);

  // --- Build Workspace user lookup map (for internal members only) ---
  workspaceUsers = {};
  let pageToken = '';
  try {
    do {
      const res = AdminDirectory.Users.list({
        customer: "my_customer",
        maxResults: 500,
        projection: 'basic',
        fields: 'users(primaryEmail),nextPageToken',
        pageToken: pageToken
      });
      if (res.users) {
        res.users.forEach(u => {
          workspaceUsers[u.primaryEmail.toLowerCase()] = true;
        });
      }
      pageToken = res.nextPageToken;
    } while (pageToken);
    Logger.info('Loaded Workspace user list', {
      count: Object.keys(workspaceUsers).length
    });
  } catch (err) {
    Logger.error('Failed to build Workspace user map', { error: err.message });
  }

  // Personal addresses, built once and only if some row actually asked for them.
  // Reuses the squadron module's map so both paths reach a cadet-lite member at
  // the same address — two writers disagreeing about which address is "theirs"
  // is how a member ends up added and removed on a loop.
  const anyRowWantsCadetLite = Object.keys(groupIncludeCadetLiteByName)
    .some(function (n) { return groupIncludeCadetLiteByName[n]; });
  const capwatchEmailByCapid = anyRowWantsCadetLite
    ? buildSquadronCapwatchPrimaryEmailByCapidMap_()
    : {};
  let membersWithCadetLite = null;

  if (anyRowWantsCadetLite) {
    membersWithCadetLite = withCadetLiteAddresses_(members, capwatchEmailByCapid);
    const reachable = Object.keys(membersWithCadetLite)
      .filter(function (c) { return !members[c].email && membersWithCadetLite[c].email; }).length;
    Logger.info('Cadet-lite members addressable for opted-in groups', {
      reachable: reachable,
      note: 'Accountless members with a CAPWATCH address. Groups without "Add Lite" are unaffected.'
    });
  }

  // Build desired group membership state
  for(let i = 1; i < groupsConfig.length; i++) {
    const groupName = groupsConfig[i][1];
    const rowMembers = (membersWithCadetLite && groupIncludeCadetLiteByName[String(groupName || '').trim()])
      ? membersWithCadetLite
      : members;
    const generatedGroups = getGroupMembers(
      groupName,
      groupsConfig[i][2],
      groupsConfig[i][3],
      rowMembers,
      squadrons,
    );

    // Multiple sheet rows may intentionally target the same base group name.
    // Merge their generated memberships instead of letting the last row win.
    if (!groups[groupName]) groups[groupName] = {};
    for (const groupId in generatedGroups) {
      if (!groups[groupName][groupId]) groups[groupName][groupId] = {};
      for (const email in generatedGroups[groupId]) {
        groups[groupName][groupId][email] = generatedGroups[groupId][email];
      }
    }
  }

  // Preserve manual members from "User Additions" by treating them as desired
  // memberships before the current-vs-desired delta pass.
  const manualByGroup = getManualGroupMembersFromUserAdditions_();
  const mergeStats = mergeManualGroupMembersIntoDesired_(groups, manualByGroup);
  Logger.info('Manual User Additions merged into desired memberships', mergeStats);

  // Calculate deltas by comparing with current state
  for (const category in groups) {
    for (const group in groups[category]) {
      // Use Groups sheet Description column for group creation whenever provided.
      // group is like "hiwg.<name>", "hi073.<name>", etc. Base name is everything after the first dot.
      const baseGroupName = group.includes('.') ? group.split('.').slice(1).join('.') : group;

      let spreadsheetDescription = (descriptionByGroupName[baseGroupName] || '').toString().trim();

      // Duty-position groups: if Description is blank, fall back to Values column as description.
      if (!spreadsheetDescription && baseGroupName.indexOf('dty.') > -1) {
        const cfgRow = groupsConfig.find(row => (row[1] || '').toString().trim() === baseGroupName);
        const valuesCol = cfgRow ? (cfgRow[3] || '').toString().trim() : '';
        spreadsheetDescription = valuesCol || 'Unknown Duty Position';
      }

      // Compute friendly group metadata:
      // - description: from spreadsheet (or fallback)
      // - name: "<Org Name> - <description>"
      const groupEmail = (group + CONFIG.EMAIL_DOMAIN).toLowerCase();
      const metaDescriptionRaw = (spreadsheetDescription || baseGroupName).toString().trim();
      const metaDescription = metaDescriptionRaw;
      let metaNameSuffix = metaDescriptionRaw;
      const groupAttribute = String(groupAttributeByName[baseGroupName] || '').trim().toLowerCase();

      // Achievement descriptions often follow "ABBR - Full Name".
      // Use the short code for the group name while keeping the full
      // description unchanged for directory details.
      if (groupAttribute === 'achievements' && metaDescriptionRaw.indexOf(' - ') > -1) {
        metaNameSuffix = String(metaDescriptionRaw.split(' - ')[0] || '').trim() || metaDescriptionRaw;
      }

      // Determine the org display name based on the groupId prefix (wing-level "ca.*" or unit-level "ca008.*")
      const groupPrefix = (group.split('.')[0] || '').toString().trim().toLowerCase();
      let orgRecord = null;

      try {
        if (groupPrefix === CONFIG.WING.toLowerCase()) {
          // Wing-level group: find the WING org record
          orgRecord = Object.values(squadrons).find(o =>
            o && String(o.scope || '').toUpperCase() === 'WING' &&
            String(o.wing || '').toLowerCase() === CONFIG.WING.toLowerCase()
          );
        } else {
          // Unit/group prefix like "ca008" or "ca445": match by wing+unit
          const wing = groupPrefix.substring(0, 2);
          const unit = groupPrefix.substring(2);
          orgRecord = Object.values(squadrons).find(o =>
            o &&
            String(o.wing || '').toLowerCase() === wing &&
            String(o.unit || '') === unit
          );
        }
      } catch (e) {
        orgRecord = null;
      }

      const orgDisplayName = orgRecord && orgRecord.name ? toSentenceCase_(String(orgRecord.name || '')) : '';
      const shortOrgDisplayName = abbreviateManagedGroupOrgDisplayName_(orgRecord);
      const metaName = (shortOrgDisplayName ? `${shortOrgDisplayName} - ${metaNameSuffix}` : metaNameSuffix);
      const metaFullDescription = (orgDisplayName ? `${orgDisplayName} - ${metaDescription}` : metaDescription);

      desiredGroupMeta[groupEmail] = {
        name: metaName,
        description: metaFullDescription
      };

      const allowExternal = !!groupAllowExternalByName[baseGroupName];
      const currentMembers = getCurrentGroup(group, squadrons, desiredGroupMeta[groupEmail], allowExternal);
      for (let i = 0; i < currentMembers.length; i++) {
        const currentEmail = currentMembers[i].email;
        const currentRole = (currentMembers[i].role || 'MEMBER').toString().toUpperCase();

        if (groups[category][group][currentEmail]) {
          // Member already in group - no change needed
          groups[category][group][currentEmail] = 0;
        } else if (currentRole === 'MEMBER') {
          // Only auto-remove plain members. Leave MANAGER/OWNER entries alone
          // unless they are explicitly managed elsewhere (for example User Additions).
          groups[category][group][currentEmail] = -1;
        }
      }
    }
  }

  saveEmailGroups(groups);
  Logger.info('Group deltas generated', {
    duration: new Date() - start + 'ms',
    categories: Object.keys(groups).length
  });
  return groups;
}

/**
 * Populates the module-level workspaceEmailMap (CAPID -> Workspace primary email)
 * and rewrites each member's email to their Workspace address where one exists.
 *
 * Shared by the live delta pass and previewEmailGroupRows(), so a preview resolves
 * the same addresses a real run would.
 *
 * @param {Object<string, Object>} members - Members object indexed by CAPID
 * @returns {Object<string, string>} The CAPID -> email map
 */
/**
 * Returns a copy of `members` in which the accountless ones carry their CAPWATCH
 * personal address, so a group can actually reach them.
 *
 * WHY A COPY. The result is handed to one Groups-sheet row and thrown away.
 * Mutating the shared member map would leak these addresses into every row
 * processed afterwards, quietly adding 1,600-odd external members to groups
 * nobody opted in — the kind of change that looks fine in a diff and is
 * discovered in a mailbox.
 *
 * A member with no Workspace account and no CAPWATCH address stays addressless
 * and reaches no group. That is the honest outcome: there is nowhere to send.
 *
 * @param {Object} members - CAPID -> member object
 * @param {Object} capwatchEmailByCapid - CAPID -> personal address
 * @returns {Object} Copy with accountless members addressed where possible
 */
function withCadetLiteAddresses_(members, capwatchEmailByCapid) {
  const out = {};
  Object.keys(members).forEach(function (capid) {
    const member = members[capid];
    if (member && !member.email) {
      const fallback = capwatchEmailByCapid[capid];
      if (fallback) {
        // Shallow copy: only the address differs, and the original must not.
        out[capid] = Object.assign({}, member, { email: fallback });
        return;
      }
    }
    out[capid] = member;
  });
  return out;
}

function buildWorkspaceEmailMapForGroups_(members) {
  workspaceEmailMap = {};
  let token = '';

  try {
    do {
      const page = AdminDirectory.Users.list({
        customer: "my_customer",
        maxResults: 500,
        projection: 'full',
        fields: 'users(primaryEmail,externalIds),nextPageToken',
        pageToken: token
      });
      if (page.users) {
        page.users.forEach(u => {
          const capidField = (u.externalIds || []).find(x => x.type === 'organization');
          if (capidField && capidField.value) {
            workspaceEmailMap[capidField.value.toString()] = u.primaryEmail.toLowerCase();
            if (members[capidField.value]) members[capidField.value].email = u.primaryEmail.toLowerCase();
          }
        });
      }
      token = page.nextPageToken;
    } while (token);
    Logger.info('Workspace CAPID→Email map built', { count: Object.keys(workspaceEmailMap).length });
  } catch (err) {
    Logger.error('Failed to build Workspace CAPID→Email map', { message: err.message });
  }

  return workspaceEmailMap;
}

/**
 * Read-only: shows what each Groups sheet row would produce, without creating a
 * group, adding a member or removing one.
 *
 * Answers the two questions the execution log cannot: which group IDs a row
 * generates, and who lands in them. A row that produces a group with zero members
 * is the signature of a Values column that matches no CAPWATCH record.
 *
 *   previewEmailGroupRows()                  // every row
 *   previewEmailGroupRows('recruiting')      // rows whose name/attribute/values match
 *   previewEmailGroupRows('level', true)     // and list the members
 *
 * @param {string} [filter] - Case-insensitive substring match on the row
 * @param {boolean} [showMembers=false] - Print each resolved address
 * @returns {void}
 */
function previewEmailGroupRows(filter, showMembers = false) {
  clearCache();
  const needle = String(filter || '').trim().toLowerCase();

  const groupsConfig = SpreadsheetApp
    .openById(CONFIG.AUTOMATION_SPREADSHEET_ID)
    .getSheetByName('Groups')
    .getDataRange()
    .getValues();

  const squadrons = getSquadrons();
  const members = getMembers();
  attachCommitteesToMembers(members, squadrons);
  buildWorkspaceEmailMapForGroups_(members);

  console.log('\n' + '='.repeat(80));
  console.log('GROUPS SHEET PREVIEW - no groups created, no memberships changed');
  console.log('='.repeat(80) + '\n');

  let rowsShown = 0;
  let emptyGroups = 0;

  for (let i = 1; i < groupsConfig.length; i++) {
    const category = String(groupsConfig[i][0] || '').trim();
    const groupName = String(groupsConfig[i][1] || '').trim();
    const attribute = String(groupsConfig[i][2] || '').trim();
    const values = String(groupsConfig[i][3] || '').trim();
    if (!groupName) continue;

    const haystack = `${category} ${groupName} ${attribute} ${values}`.toLowerCase();
    if (needle && haystack.indexOf(needle) === -1) continue;

    rowsShown++;
    console.log(`Row ${i + 1}: ${groupName}`);
    console.log(`  Category: ${category || '(blank)'}   Attribute: ${attribute || '(blank)'}`);
    console.log(`  Values:   ${values || '(blank)'}`);

    let generated = {};
    try {
      generated = getGroupMembers(groupName, attribute, values, members, squadrons) || {};
    } catch (e) {
      console.log(`  ✗ Threw: ${e.message}`);
      console.log('');
      continue;
    }

    const groupIds = Object.keys(generated);
    if (!groupIds.length) {
      console.log('  → creates NO groups');
      console.log('');
      continue;
    }

    groupIds.sort();
    for (let g = 0; g < groupIds.length; g++) {
      const memberEmails = Object.keys(generated[groupIds[g]]);
      const marker = memberEmails.length === 0 ? ' ⚠ EMPTY' : '';
      if (memberEmails.length === 0) emptyGroups++;
      console.log(`  → ${groupIds[g]}${CONFIG.EMAIL_DOMAIN}: ${memberEmails.length} member(s)${marker}`);
      if (showMembers) {
        memberEmails.sort().forEach(email => console.log(`       ${email}`));
      }
    }
    console.log('');
  }

  console.log('='.repeat(80));
  console.log(`Rows shown: ${rowsShown}   Groups that would be empty: ${emptyGroups}`);
  console.log('An empty group usually means the Values column matches no CAPWATCH record.');
  console.log('='.repeat(80) + '\n');
}

/**
 * Deletes group-echelon DLs that exist, are empty, and are for a position that
 * echelon does not establish — the ones minted before pruneEmptySeededGroups_()
 * existed (`ca0X.dty.director-it` at every group, and the same for a rule keyed
 * on a wing-only office such as Inspector General).
 *
 * Candidates are computed from the Groups sheet, so nothing outside the managed
 * naming scheme is ever considered. A group is offered only when it is **empty**,
 * which is also why deleting it loses nothing: an empty DL has no manual
 * additions to lose, and one whose position really does exist at that echelon is
 * recreated the moment somebody holds it.
 *
 *   cleanupEmptyEchelonGroups()        // preview
 *   cleanupEmptyEchelonGroups(false)   // delete
 *
 * @param {boolean} [dryRun=true] - When true, only reports
 * @returns {{toDelete: string[], deleted: string[], errors: Object[]}}
 */
function cleanupEmptyEchelonGroups(dryRun = true) {
  clearCache();
  const summary = { toDelete: [], deleted: [], errors: [] };

  const groupsConfig = SpreadsheetApp
    .openById(CONFIG.AUTOMATION_SPREADSHEET_ID)
    .getSheetByName('Groups')
    .getDataRange()
    .getValues();

  const squadrons = getSquadrons();
  const echelonOrgs = Object.values(squadrons).filter(org =>
    org &&
    String(org.scope || '').toUpperCase() === 'GROUP' &&
    String(org.wing || '').toLowerCase() === CONFIG.WING.toLowerCase() &&
    String(org.unit || '') !== '000' &&
    String(org.unit || '') !== '001'
  );

  console.log('\n' + '='.repeat(80));
  console.log(dryRun ? 'EMPTY GROUP-ECHELON DLs - PREVIEW' : 'EMPTY GROUP-ECHELON DLs - DELETING');
  console.log('='.repeat(80) + '\n');

  for (let i = 1; i < groupsConfig.length; i++) {
    const groupName = String(groupsConfig[i][1] || '').trim();
    const attribute = String(groupsConfig[i][2] || '').trim();
    if (!groupName) continue;

    // Only the rules that create per-group-echelon DLs.
    if (['dutyPositionIds', 'rank', 'dutyPositionLevelStaff', 'achievements', 'committeeIds']
      .indexOf(attribute) === -1) continue;

    for (let g = 0; g < echelonOrgs.length; g++) {
      const org = echelonOrgs[g];
      const groupEmail = (String(org.wing).toLowerCase() + String(org.unit) + '.' + groupName + CONFIG.EMAIL_DOMAIN).toLowerCase();

      if (!managedGroupExists_(groupEmail)) continue;

      let memberCount = 0;
      try {
        const page = AdminDirectory.Members.list(groupEmail, { maxResults: 1 });
        memberCount = (page.members || []).length;
      } catch (e) {
        summary.errors.push({ groupEmail: groupEmail, error: e.message });
        continue;
      }

      if (memberCount > 0) continue;

      if (dryRun) {
        summary.toDelete.push(groupEmail);
        console.log(`  would delete  ${groupEmail}   (row ${i + 1}, ${attribute})`);
        continue;
      }

      try {
        executeWithRetry(() => AdminDirectory.Groups.remove(groupEmail));
        summary.deleted.push(groupEmail);
        console.log(`  deleted       ${groupEmail}`);
      } catch (e) {
        summary.errors.push({ groupEmail: groupEmail, error: e.message });
        console.log(`  ✗ ${groupEmail}: ${e.message}`);
      }
    }
  }

  console.log('\n' + '='.repeat(80));
  if (dryRun) {
    console.log(`Empty group-echelon DLs found: ${summary.toDelete.length}`);
    console.log('Review the list, then run: cleanupEmptyEchelonGroups(false)');
  } else {
    console.log(`Deleted: ${summary.deleted.length}   Errors: ${summary.errors.length}`);
  }
  console.log('='.repeat(80) + '\n');

  Logger.info('Empty echelon DL cleanup finished', {
    dryRun: dryRun,
    toDelete: summary.toDelete.length,
    deleted: summary.deleted.length,
    errors: summary.errors.length
  });

  return summary;
}

/**
 * Loads manual member/group mappings from "User Additions".
 * Expected columns (same as updateAdditionalGroupMembers):
 * - [1] Email
 * - [3] Groups (comma-separated group IDs, optionally full group emails)
 *
 * @returns {Object<string, Object<string, number>>} groupId -> { email: 1 }
 */
function getManualGroupMembersFromUserAdditions_() {
  const out = {};
  try {
    const sheet = SpreadsheetApp
      .openById(CONFIG.AUTOMATION_SPREADSHEET_ID)
      .getSheetByName('User Additions');
    if (!sheet) {
      Logger.warn('User Additions tab not found; skipping manual preserve merge');
      return out;
    }

    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const email = (rows[i][1] || '').toString().trim().toLowerCase();
      const groupsCell = (rows[i][3] || '').toString().trim();
      if (!email || !groupsCell) continue;

      const groupTokens = groupsCell.split(',')
        .map(g => (g || '').toString().trim().toLowerCase())
        .filter(Boolean);

      for (let j = 0; j < groupTokens.length; j++) {
        let groupId = groupTokens[j];
        if (groupId.endsWith(CONFIG.EMAIL_DOMAIN.toLowerCase())) {
          groupId = groupId.slice(0, -CONFIG.EMAIL_DOMAIN.length);
        }
        if (!groupId) continue;

        if (!out[groupId]) out[groupId] = {};
        out[groupId][email] = 1;
      }
    }

    Logger.info('Manual members loaded from User Additions', {
      groups: Object.keys(out).length
    });
  } catch (e) {
    Logger.warn('Failed to load User Additions for preserve merge', {
      errorMessage: e.message
    });
  }
  return out;
}

/**
 * Merges manual group members into desired memberships across matching groups.
 *
 * @param {Object} groups Desired memberships by category/group/email
 * @param {Object<string, Object<string, number>>} manualByGroup groupId -> { email: 1 }
 * @returns {{groupsMatched:number, groupsUnmatched:number, membersMerged:number}}
 */
function mergeManualGroupMembersIntoDesired_(groups, manualByGroup) {
  const index = {};
  for (const category in groups) {
    for (const groupId in groups[category]) {
      if (!index[groupId]) index[groupId] = [];
      index[groupId].push(category);
    }
  }

  let groupsMatched = 0;
  let groupsUnmatched = 0;
  let membersMerged = 0;

  for (const groupId in manualByGroup) {
    const categories = index[groupId] || [];
    if (!categories.length) {
      groupsUnmatched++;
      continue;
    }
    groupsMatched++;

    for (let i = 0; i < categories.length; i++) {
      const category = categories[i];
      const desired = groups[category][groupId];
      for (const email in manualByGroup[groupId]) {
        if (!desired[email]) {
          desired[email] = 1;
          membersMerged++;
        }
      }
    }
  }

  return { groupsMatched, groupsUnmatched, membersMerged };
}

/**
 * Builds group membership lists based on member attributes
 * Creates wing, group, and (for member-type only) unit-level groups
 * @param {string} groupName - Base name of the group
 * @param {string} attribute - Member attribute to filter by (type, rank, dutyPositionIds, etc.)
 * @param {string} attributeValues - Comma-separated list of values to match
 * @param {Object} members - Members object indexed by CAPID
 * @param {Object} squadrons - Squadrons object indexed by orgid
 * @returns {Object} Groups object with member emails
 */
function getGroupMembers(groupName, attribute, attributeValues, members, squadrons) {
  let groups = {};
  let wingGroupId = CONFIG.WING.toLowerCase() + '.' + groupName;
  let values = (attributeValues == null ? '' : String(attributeValues)).split(',');
  values = values.map(v => (v || '').toString().trim()).filter(v => v.length > 0);
  // True if the member attribute (string or array) matches ANY of the configured values.
  function matchesAny_(memberValue, allowedValues) {
    if (!memberValue) return false;
    if (typeof memberValue === 'string') {
      return allowedValues.indexOf(memberValue) > -1;
    }
    if (Array.isArray(memberValue)) {
      return allowedValues.some(v => memberValue.indexOf(v) > -1);
    }
    return false;
  }

  // Manual/IAOD rows may store multiple duty IDs as a single comma-separated string
  // (optionally with trailing markers like "(A)" / "(P)").
  //
  // Both sides of the comparison run through the same CAPR 30-1 renames the
  // signature generator applies (DUTY_TITLE_OVERRIDES in UpdateMembers.gs), so a
  // Groups row keyed on a current title still matches the eServices rows that
  // carry the retired one — "Recruiting & Retention Officer" is the live example,
  // and a row keyed on "Recruiting Officer" silently missed those holders.
  function normalizeDutyId_(s) {
    const raw = (s || '')
      .toString()
      .replace(/\s*\((a|p)\)\s*$/i, '')
      .trim()
      .replace(/\s+/g, ' ');
    if (!raw) return '';

    const canonical = (typeof formatDutyTitle_ === 'function') ? formatDutyTitle_(raw) : raw;
    return canonical.toLowerCase();
  }

  function expandDutyIds_(memberValue) {
    const out = [];
    if (!memberValue) return out;

    const pushToken_ = (v) => {
      const parts = (v || '').toString().split(',');
      for (let i = 0; i < parts.length; i++) {
        const norm = normalizeDutyId_(parts[i]);
        if (norm) out.push(norm);
      }
    };

    if (typeof memberValue === 'string') {
      pushToken_(memberValue);
      return out;
    }

    if (Array.isArray(memberValue)) {
      for (let i = 0; i < memberValue.length; i++) {
        pushToken_(memberValue[i]);
      }
      return out;
    }

    return out;
  }

  function matchesAnyDutyId_(memberValue, allowedValues) {
    const memberIds = expandDutyIds_(memberValue);
    if (!memberIds.length) return false;

    const allowed = {};
    for (let i = 0; i < allowedValues.length; i++) {
      const norm = normalizeDutyId_(allowedValues[i]);
      if (norm) allowed[norm] = 1;
    }

    for (let i = 0; i < memberIds.length; i++) {
      if (allowed[memberIds[i]]) return true;
    }
    return false;
  }

  let charterOrgLookup_ = null;

  function getOrgByCharter_(charter) {
    const normalized = (charter || '').toString().trim().toUpperCase();
    if (!normalized) return null;

    if (!charterOrgLookup_) {
      charterOrgLookup_ = {};
      for (const orgid in squadrons) {
        const org = squadrons[orgid];
        const orgCharter = (org && org.charter ? String(org.charter) : '').trim().toUpperCase();
        if (orgCharter) {
          charterOrgLookup_[orgCharter] = org;
        }
      }
    }

    return charterOrgLookup_[normalized] || null;
  }

  function getDutyAssignmentOrg_(dutyPosition) {
    const value = dutyPosition && dutyPosition.value ? String(dutyPosition.value) : '';
    const match = value.match(/\(([^()]+)\)\s*$/);
    return match ? getOrgByCharter_(match[1]) : null;
  }

  let groupId;
  groups[wingGroupId] = {};

  switch (attribute) {
    case 'type':
    case 'dutyPositionIds':
    case 'rank':
      // Pre-seed parent-group variants for configured group-level lists so they
      // continue to exist even when a parent group currently has zero members.
      // Keep unit-level member-type groups dynamic; only real parent GROUP orgs
      // are pre-created here. Seeds that stay empty are pruned after the member
      // loop — see pruneEmptySeededGroups_().
      const seededEchelonGroupIds = [];
      if (attribute !== 'type') {
        for (const orgid in squadrons) {
          const org = squadrons[orgid];
          if (
            org &&
            String(org.scope || '').toUpperCase() === 'GROUP' &&
            String(org.wing || '').toLowerCase() === CONFIG.WING.toLowerCase() &&
            String(org.unit || '') !== '000' &&
            String(org.unit || '') !== '001'
          ) {
            const seededGroupId = String(org.wing || '').toLowerCase() + String(org.unit || '') + '.' + groupName;
            if (!groups[seededGroupId]) groups[seededGroupId] = {};
            seededEchelonGroupIds.push(seededGroupId);
          }
        }
      }

      for (const member in members) {
        const isMatch = (attribute === 'dutyPositionIds')
          ? matchesAnyDutyId_(members[member][attribute], values)
          : matchesAny_(members[member][attribute], values);

        if (
          isMatch &&
          members[member].email
        ) {
          // Wing-level group
          groups[wingGroupId][members[member].email] = 1;

          // Group-level group (only if parent org is a real GROUP)
          const parent = squadrons[members[member].group];
          if (parent && parent.scope === 'GROUP') {
            groupId =
              squadrons[members[member].orgid].wing.toLowerCase() +
              parent.unit +
              '.' +
              groupName;
            if (!groups[groupId]) groups[groupId] = {};
            groups[groupId][members[member].email] = 1;
          }

          // Unit-level groups (only for member-type categories)
          if (attribute === 'type') {
            const org = squadrons[members[member].orgid];
            if (org && org.unit && org.scope === 'UNIT' && org.unit !== '001') {
              const unitGroupId = org.wing.toLowerCase() + org.unit + '.' + groupName;
              if (!groups[unitGroupId]) groups[unitGroupId] = {};
              groups[unitGroupId][members[member].email] = 1;
            }
          }
        }
      }

      pruneEmptySeededGroups_(groups, seededEchelonGroupIds, groupName);
      break;

    case 'dutyPositionIdsWingHQ':
      // Wing HQ ONLY (unit 001) for the configured duty position titles.
      // Spreadsheet usage:
      //   Category: duty-position
      //   Group Name: <your DL base name>
      //   Attribute: dutyPositionIdsWingHQ
      //   Values: Director of Safety,Safety Officer
      // Result:
      //   Creates ONLY: <wing>.<groupName> with Wing HQ duty holders
      //   (e.g., hi.all-safety-hq)
      //
      // Matches on the org the DUTY is held at, not the member's home unit. Wing
      // staff are overwhelmingly assigned to a squadron and hold their wing duty on
      // top of that; testing the home unit is why a wing office DL could come back
      // holding the assistant and not the primary. A Wing HQ member who holds the
      // same title at a squadron is excluded — the same rule read from the other side.

      groupId = wingGroupId;
      if (!groups[groupId]) groups[groupId] = {};

      let wingHqDutyOrgUnresolved = 0;
      let wingHqDutyOrgFromHomeUnit = 0;

      for (const member in members) {
        if (!members[member].email || !Array.isArray(members[member].dutyPositions)) continue;

        for (let i = 0; i < members[member].dutyPositions.length; i++) {
          const dutyPosition = members[member].dutyPositions[i];
          if (!matchesAnyDutyId_(dutyPosition.id, values)) continue;

          // A Manual Members row records a duty with no charter, because the sheet
          // has nowhere to put one — such a duty is held at the member's own org by
          // definition, so fall back to that rather than dropping the row.
          let dutyOrg = getDutyAssignmentOrg_(dutyPosition);
          if (!dutyOrg) {
            dutyOrg = squadrons[members[member].orgid] || null;
            if (dutyOrg) wingHqDutyOrgFromHomeUnit++;
          }
          if (!dutyOrg) {
            wingHqDutyOrgUnresolved++;
            continue;
          }

          if (
            String(dutyOrg.scope || '').toUpperCase() === 'WING' &&
            String(dutyOrg.unit || '') === '001' &&
            String(dutyOrg.wing || '').toUpperCase() === CONFIG.WING.toUpperCase()
          ) {
            groups[groupId][members[member].email] = 1;
            break;
          }
        }
      }

      Logger.info('Wing HQ duty group resolved', {
        group: groupId,
        values: values.join(', '),
        members: Object.keys(groups[groupId]).length,
        dutyOrgFromHomeUnit: wingHqDutyOrgFromHomeUnit,
        dutyOrgUnresolved: wingHqDutyOrgUnresolved
      });

      // Prevent creating an empty group
      if (Object.keys(groups[groupId]).length === 0) {
        Logger.warn('Wing HQ duty group matched nobody - not creating it', {
          group: groupId,
          values: values.join(', '),
          note: 'Values must match the CAPWATCH duty title exactly, after the CAPR 30-1 renames'
        });
        delete groups[groupId];
      }
      break;

    case 'dutyPositionIdsGroupScope':
      // Wing-level DL containing members who hold one of the given duty IDs at a
      // GROUP-echelon org ONLY. Wing HQ holders and squadron/unit holders are
      // intentionally excluded, and NO per-group child DLs are created.
      //
      // Built for the "group commanders" list: the commanders of the group-echelon
      // units (CC at GROUP scope) collapsed into a single wing-level list. Squadron
      // commanders (CC at UNIT scope) and any Wing HQ CC are excluded because the
      // duty's own assigned org scope is checked, not the member's home unit.
      //
      // Spreadsheet usage:
      //   Category:   duty-position
      //   Group Name: group-commanders
      //   Attribute:  dutyPositionIdsGroupScope
      //   Values:     CC
      // Result: creates ONLY <wing>.group-commanders (e.g., ca.group-commanders).
      // The wing Deputy Commander over the groups is not encoded distinctly in
      // CAPWATCH, so add that person via the "User Additions" tab, not here.
      groupId = wingGroupId;
      if (!groups[groupId]) groups[groupId] = {};

      for (const member in members) {
        if (!members[member].email || !Array.isArray(members[member].dutyPositions)) continue;

        for (let i = 0; i < members[member].dutyPositions.length; i++) {
          const dutyPosition = members[member].dutyPositions[i];
          if (!matchesAnyDutyId_(dutyPosition.id, values)) continue;

          const dutyOrg = getDutyAssignmentOrg_(dutyPosition);
          if (
            dutyOrg &&
            String(dutyOrg.scope || '').toUpperCase() === 'GROUP' &&
            String(dutyOrg.wing || '').toUpperCase() === CONFIG.WING.toUpperCase() &&
            String(dutyOrg.unit || '') !== '000' &&
            String(dutyOrg.unit || '') !== '001'
          ) {
            groups[wingGroupId][members[member].email] = 1;
            break;
          }
        }
      }

      // Prevent creating an empty group
      if (Object.keys(groups[groupId]).length === 0) {
        delete groups[groupId];
      }
      break;

    case 'dutyPositionIdsAndLevel':
      // Prevent creation of Wing HQ-level (hi001.* or 000.*) duty lists
      groupId = groupName;

      // Only build duty groups for Group- and Squadron-level orgs (not Wing HQ or placeholders)
      if (!groups[groupId]) groups[groupId] = {};

      for (const member in members) {
        const org = squadrons[members[member].orgid];
        // Only process if org is not Wing HQ or placeholder units
        if (
          org &&
          org.scope !== 'WING' &&
          org.unit !== '000' &&
          org.unit !== '001' &&
          members[member][attribute] &&
          (
            (typeof members[member][attribute] === 'string' &&
              values.indexOf(members[member][attribute]) > -1) ||
            (Array.isArray(members[member][attribute]) &&
              members[member][attribute].indexOf(values[0]) > -1)
          ) &&
          members[member].email
        ) {
          groups[groupId][members[member].email] = 1;
        }
      }
      // If no members were added, remove the empty group (prevents hi001.* creation)
      if (Object.keys(groups[groupId]).length === 0) {
        delete groups[groupId];
      }
      break;

    case 'dutyPositionLevel':
      groupId = wingGroupId;
      if (groupId && !groups[groupId]) {
        groups[groupId] = {};
      }
      for(const member in members) {
        for (let i = 0; i < members[member].dutyPositions.length; i++) {
          if (members[member].dutyPositions[i].level === values[0] && members[member].email) {
            groups[groupId][members[member].email] = 1;
            break;
          }
        }
      }
      break;

    case 'dutyPositionLevelStaff':
      const staffLevel = (values[0] || '').toString().trim().toUpperCase();
      const seededStaffGroupIds = [];

      if (staffLevel === 'WING') {
        for (const member in members) {
          if (!members[member].email || !Array.isArray(members[member].dutyPositions)) continue;

          for (let i = 0; i < members[member].dutyPositions.length; i++) {
            const dutyPosition = members[member].dutyPositions[i];
            const dutyOrg = getDutyAssignmentOrg_(dutyPosition);

            if (
              String(dutyPosition.level || '').toUpperCase() === 'WING' &&
              dutyOrg &&
              String(dutyOrg.scope || '').toUpperCase() === 'WING' &&
              String(dutyOrg.unit || '') === '001' &&
              String(dutyOrg.wing || '').toUpperCase() === CONFIG.WING.toUpperCase()
            ) {
              groups[wingGroupId][members[member].email] = 1;
              break;
            }
          }
        }
      } else if (staffLevel === 'GROUP') {
        delete groups[wingGroupId];

        for (const orgid in squadrons) {
          const org = squadrons[orgid];
          if (
            org &&
            String(org.scope || '').toUpperCase() === 'GROUP' &&
            String(org.wing || '').toUpperCase() === CONFIG.WING.toUpperCase() &&
            String(org.unit || '') !== '000' &&
            String(org.unit || '') !== '001'
          ) {
            groupId = String(org.wing || '').toLowerCase() + String(org.unit || '') + '.' + groupName;
            if (!groups[groupId]) groups[groupId] = {};
            seededStaffGroupIds.push(groupId);
          }
        }

        for (const member in members) {
          if (!members[member].email || !Array.isArray(members[member].dutyPositions)) continue;

          for (let i = 0; i < members[member].dutyPositions.length; i++) {
            const dutyPosition = members[member].dutyPositions[i];
            const dutyOrg = getDutyAssignmentOrg_(dutyPosition);

            if (
              String(dutyPosition.level || '').toUpperCase() === 'GROUP' &&
              dutyOrg &&
              String(dutyOrg.scope || '').toUpperCase() === 'GROUP' &&
              String(dutyOrg.wing || '').toUpperCase() === CONFIG.WING.toUpperCase() &&
              String(dutyOrg.unit || '') !== '000' &&
              String(dutyOrg.unit || '') !== '001'
            ) {
              groupId = String(dutyOrg.wing || '').toLowerCase() + String(dutyOrg.unit || '') + '.' + groupName;
              if (!groups[groupId]) groups[groupId] = {};
              groups[groupId][members[member].email] = 1;
              break;
            }
          }
        }

        pruneEmptySeededGroups_(groups, seededStaffGroupIds, groupName);
      } else {
        delete groups[wingGroupId];
      }
      break;

    case 'achievements':
      // MbrAchievements.txt records the achievement by NUMERIC AchvID ([1]), never by
      // name — Level I is 96, not the string "Level I". A Values column holding names
      // (the natural thing to type for the Education & Training levels) therefore
      // matched no row at all and produced an empty group with no error. Values may
      // now be either form: numeric IDs pass through, names are resolved against
      // Achievements.txt.
      let achievements = parseFile('MbrAchievements');
      const achievementIds = resolveAchievementValuesToIds_(values, groupName);
      const achievementStatusCounts = {};
      let achievementRowsMatched = 0;
      let achievementRowsNotAMember = 0;
      let achievementRowsNoEmail = 0;

      for(let i = 0; i < achievements.length; i++) {
        const achvCapid = achievements[i][0];
        const achvId = String(achievements[i][1] == null ? '' : achievements[i][1]).trim();
        const achvStatus = String(achievements[i][2] == null ? '' : achievements[i][2]).trim().toUpperCase();

        if (!achievementIds[achvId]) continue;

        achievementRowsMatched++;
        const statusKey = achvStatus || '(blank)';
        achievementStatusCounts[statusKey] = (achievementStatusCounts[statusKey] || 0) + 1;

        if (['ACTIVE', 'TRAINING'].indexOf(achvStatus) === -1) continue;
        if (!members[achvCapid]) {
          achievementRowsNotAMember++;
          continue;
        }
        if (!members[achvCapid].email) {
          achievementRowsNoEmail++;
          continue;
        }

        groups[wingGroupId][members[achvCapid].email] = 1;
        // Group-level achievement DLs: ONLY when the member's parent org is a real GROUP.
        // Prevents duplicate Wing HQ groups like "hi001.*".
        const parent = members[achvCapid].group ? squadrons[members[achvCapid].group] : null;
        if (parent && parent.scope === 'GROUP' && parent.unit && parent.unit !== '001' && parent.unit !== '000') {
          groupId =
            squadrons[members[achvCapid].orgid].wing.toLowerCase() +
            parent.unit +
            '.' +
            groupName;
          if (!groups[groupId]) {
            groups[groupId] = {};
          }
          groups[groupId][members[achvCapid].email] = 1;
        }
      }

      Logger.info('Achievement group resolved', {
        group: wingGroupId,
        values: values.join(', '),
        resolvedAchvIds: Object.keys(achievementIds).join(', '),
        members: Object.keys(groups[wingGroupId]).length,
        matchingRows: achievementRowsMatched,
        rowsByStatus: achievementStatusCounts,
        skippedNotAMember: achievementRowsNotAMember,
        skippedNoWorkspaceEmail: achievementRowsNoEmail
      });

      if (Object.keys(groups[wingGroupId]).length === 0) {
        // Leave the group out of the desired state entirely rather than creating it
        // empty — and, just as important, rather than letting a misconfigured row
        // empty a group that already has the right people in it.
        Logger.warn('Achievement group matched no members - leaving it unmanaged this run', {
          group: wingGroupId,
          values: values.join(', '),
          resolvedAchvIds: Object.keys(achievementIds).join(', ') || '(none)',
          matchingRows: achievementRowsMatched,
          rowsByStatus: achievementStatusCounts,
          note: achievementRowsMatched === 0
            ? 'No MbrAchievements row carries these achievements. Check the Values column against Achievements.txt.'
            : 'Rows exist but no row is ACTIVE/TRAINING for a member with a Workspace account.'
        });
        for (const emptyGroupId in groups) {
          if (Object.keys(groups[emptyGroupId]).length === 0) delete groups[emptyGroupId];
        }
      }
      break;

    case 'professionalLevel':
      // Professional development levels — Level 1 through Level 5.
      //
      // These are NOT in MbrAchievements. The post-2018 PD program lives in its own
      // CAPWATCH subsystem, the PL_* tables: PL_Paths names each path, and
      // PL_MemberPathCredit records who has credit for it and whether it is approved.
      // Achievements.txt still LISTS "Level II".."Level V" (AchvIDs 131-134) from the
      // retired program, which is the trap — a Groups row keyed on those IDs resolves
      // cleanly and then matches nobody, because no member record uses them any more.
      //
      // A member appears ONLY in the group for their HIGHEST completed level. A
      // Level V holder is in all-level-v and nowhere else — the levels are rungs,
      // not badges you accumulate, and a wing that mails "all-level-ii" means the
      // people sitting at Level II, not everyone who ever passed through it.
      //
      // Spreadsheet usage:
      //   Category:   education-training
      //   Group Name: all-level-ii
      //   Attribute:  professionalLevel
      //   Values:     Level 2          (a PathID also works; "Level II" is accepted)
      //
      // Level 2 is two paths in CAPWATCH ("Level 2 Part 1" and "Part 2") and counts
      // as completed only when BOTH are approved. For the people partway through,
      // see professionalLevelInProgress below. Several values are an OR: "Level 4,
      // Level 5" is anyone whose highest is either.
      //
      // A value naming something off the level ladder — "Squadron Commander
      // Training", "TLC Basic" — has no rung to be highest, so it keeps plain
      // "holds this path" semantics.

    case 'professionalLevelInProgress':
      // The other side of the same table: members who hold SOME parts of a level
      // but not all of them. Built for "finished Level II Part 1, hasn't finished
      // Part 2" — the list a wing sends a nudge to.
      //
      // Spreadsheet usage:
      //   Group Name: all-level-ii-part-1-only
      //   Attribute:  professionalLevelInProgress
      //   Values:     Level 2
      //
      // A level CAPWATCH stores as a single path can have no in-progress state, and
      // says so rather than creating a group that can never fill.
      const wantsInProgress = attribute === 'professionalLevelInProgress';
      const levelSpec = resolveProfessionalLevelSpec_(values, groupName, wantsInProgress);
      const levelLadder = getProfessionalLevelLadder_();
      const approvedByCapid = getApprovedPathCreditsByCapid_();

      let levelMembersMatched = 0;
      let levelSupersededByHigher = 0;
      let levelPartialSeen = 0;

      for (const member in members) {
        if (!members[member].email) continue;

        const earned = approvedByCapid[String(member)] || {};
        const standing = summarizeMemberLevelStanding_(earned, levelLadder);
        if (standing.partial.length) levelPartialSeen++;

        let qualifies = false;

        if (wantsInProgress) {
          qualifies = levelSpec.levels.some(n => standing.partial.indexOf(n) > -1);
        } else {
          qualifies = levelSpec.levels.some(n => standing.highest === n) ||
            levelSpec.paths.some(p => p.pathIds.length && p.pathIds.every(id => earned[id]));

          // Completed this level, but has since moved past it. Counted so the log
          // explains a group that shrank the day this rule arrived.
          if (!qualifies && levelSpec.levels.some(n => standing.completed.indexOf(n) > -1)) {
            levelSupersededByHigher++;
          }
        }

        if (!qualifies) continue;

        levelMembersMatched++;
        groups[wingGroupId][members[member].email] = 1;

        const parent = members[member].group ? squadrons[members[member].group] : null;
        if (parent && parent.scope === 'GROUP' && parent.unit && parent.unit !== '001' && parent.unit !== '000') {
          groupId = squadrons[members[member].orgid].wing.toLowerCase() + parent.unit + '.' + groupName;
          if (!groups[groupId]) groups[groupId] = {};
          groups[groupId][members[member].email] = 1;
        }
      }

      Logger.info('Professional level group resolved', {
        group: wingGroupId,
        mode: wantsInProgress ? 'in progress' : 'highest completed level',
        values: values.join(', '),
        resolvedLevels: levelSpec.levels.join(', ') || '(none)',
        resolvedPaths: levelSpec.paths.map(p => `${p.label}=[${p.pathIds.join('+')}]`).join(', ') || '(none)',
        members: levelMembersMatched,
        excludedHoldingAHigherLevel: levelSupersededByHigher,
        membersWithAnyPartialLevel: levelPartialSeen,
        capidsWithAnyApprovedPath: Object.keys(approvedByCapid).length
      });

      if (Object.keys(groups[wingGroupId]).length === 0) {
        Logger.warn('Professional level group matched no members - leaving it unmanaged this run', {
          group: wingGroupId,
          mode: wantsInProgress ? 'in progress' : 'highest completed level',
          values: values.join(', '),
          resolvedLevels: levelSpec.levels.join(', ') || '(none)',
          excludedHoldingAHigherLevel: levelSupersededByHigher,
          note: Object.keys(approvedByCapid).length === 0
            ? 'PL_MemberPathCredit.txt is missing or empty in the CAPWATCH folder.'
            : (levelSupersededByHigher > 0
              ? 'Everyone who completed this level has gone past it, and each is in the group for their own highest level.'
              : 'Paths resolved and credits loaded, but nobody with an account is at this level.')
        });
        for (const emptyGroupId in groups) {
          if (Object.keys(groups[emptyGroupId]).length === 0) delete groups[emptyGroupId];
        }
      }
      break;

    case 'contact':
      // Always include ALL cadets (Workspace primary emails) at wing level only
      for (const member in members) {
        const m = members[member];
        if (!m) continue;

        // Exclude Wing HQ (unit 001)
        const org = squadrons[m.orgid];
        if (org && org.scope === 'WING' && org.unit === '001') continue;

        if (m.email && (m.type || '').toString().trim() === 'CADET') {
          groups[wingGroupId][m.email] = 1;
        }
      }
      let contacts = parseFile('MbrContact');
      for (let i = 0; i < contacts.length; i++) {
        if (members[contacts[i][0]] &&
            values.indexOf(contacts[i][1]) > -1 &&
            contacts[i][6] == 'False') {
          // Exclude Wing HQ (unit 001)
          const org = squadrons[members[contacts[i][0]].orgid];
          if (org && org.scope === 'WING' && org.unit === '001') continue;
          let contact = sanitizeEmail(contacts[i][3]);
          if (contact) {
            groups[wingGroupId][contact] = 1;
            groupId = members[contacts[i][0]].group ?
              (squadrons[members[contacts[i][0]].orgid].wing.toLowerCase() +
               squadrons[members[contacts[i][0]].group].unit + '.' + groupName) : '';
            if (groupId) {
              if (!groups[groupId]) {
                groups[groupId] = {};
              }
              groups[groupId][contact] = 1;
            }
          } else {
            Logger.warn('Invalid contact email - skipping', {
              capsn: contacts[i][0],
              rawEmail: contacts[i][3],
              contactType: contacts[i][1]
            });
          }
        }
      }
      break;

    case 'committeeIds':
      for (const member in members) {
        const email = members[member].email;
        if (!email) continue;

        const assigns = members[member].committeeAssignments;
        if (!Array.isArray(assigns) || assigns.length === 0) continue;

        for (let i = 0; i < assigns.length; i++) {
          const a = assigns[i];
          if (!a || !a.name || !a.orgid) continue;
          if (values.indexOf(a.name) === -1) continue;

          const committeeOrg = squadrons[a.orgid];
          if (!committeeOrg) continue;

          // Wing-scoped committee DL
          if (committeeOrg.scope === 'WING') {
            groups[wingGroupId][email] = 1;
            continue;
          }

          // Group-scoped committee DL (no squadron/unit committees by requirement)
          if (committeeOrg.scope === 'GROUP') {
            groupId =
              committeeOrg.wing.toLowerCase() +
              committeeOrg.unit +
              '.' +
              groupName;
            if (!groups[groupId]) groups[groupId] = {};
            groups[groupId][email] = 1;
          }
        }
      }
      break;

    case 'manualOnly':
      // Create exactly the group IDs listed in Values without deriving members
      // from local CAPWATCH data. User Additions can then supply nested external
      // groups or other managed members later in the pipeline.
      delete groups[wingGroupId];

      if (!values.length) {
        groups[wingGroupId] = {};
        break;
      }

      for (let i = 0; i < values.length; i++) {
        let explicitGroupId = String(values[i] || '').trim().toLowerCase();
        if (!explicitGroupId) continue;
        if (explicitGroupId.endsWith(CONFIG.EMAIL_DOMAIN.toLowerCase())) {
          explicitGroupId = explicitGroupId.slice(0, -CONFIG.EMAIL_DOMAIN.length);
        }
        if (!explicitGroupId) continue;
        if (!groups[explicitGroupId]) groups[explicitGroupId] = {};
      }
      break;

    default:
      // An unrecognized Attribute used to fall through here having already had its
      // wing-level group pre-created at the top of this function, so a typo in the
      // Groups sheet silently produced a real, permanently empty Google Group.
      // Nothing is created for an attribute this code cannot honor.
      delete groups[wingGroupId];
      Logger.warn('Unknown attribute type - no group created for this row', {
        attribute: attribute,
        groupName: groupName,
        wouldHaveCreated: wingGroupId + CONFIG.EMAIL_DOMAIN
      });
  }
  return groups;
}

/**
 * Drops pre-seeded group-echelon DLs that nobody qualifies for.
 *
 * A duty rule seeds one DL per group-echelon org up front, so that a DL whose
 * holder has just left stays in the desired state and gets that person removed
 * from it. That is right for a position the echelon **has** and is merely vacant,
 * and wrong for one it does not have: CAP puts a **Director of Information
 * Technology at wing and above and an IT Officer below it**, and groups have **no
 * Inspector General**, so those rows minted one empty DL per group, every run,
 * forever.
 *
 * The two cases are told apart by asking whether the group already exists. An
 * empty seed for a DL that is already in the tenant is kept (a real position,
 * currently vacant, whose stale members still need clearing); an empty seed for a
 * group that does not exist is dropped rather than created.
 *
 * @param {Object} groups - Generated groups, mutated in place
 * @param {string[]} seededGroupIds - Group ids seeded before members were matched
 * @param {string} groupName - Base group name, for logging
 * @returns {void}
 */
function pruneEmptySeededGroups_(groups, seededGroupIds, groupName) {
  if (!seededGroupIds || !seededGroupIds.length) return;

  const dropped = [];
  const keptVacant = [];

  for (let i = 0; i < seededGroupIds.length; i++) {
    const seededGroupId = seededGroupIds[i];
    if (!groups[seededGroupId] || Object.keys(groups[seededGroupId]).length > 0) continue;

    if (managedGroupExists_(seededGroupId + CONFIG.EMAIL_DOMAIN)) {
      keptVacant.push(seededGroupId);
      continue;
    }

    delete groups[seededGroupId];
    dropped.push(seededGroupId);
  }

  if (dropped.length || keptVacant.length) {
    Logger.info('Pruned empty group-echelon DLs', {
      groupName: groupName,
      notCreated: dropped.length,
      keptExistingButVacant: keptVacant.length,
      notCreatedIds: dropped.join(', '),
      note: 'A position that echelon does not establish produces no group; an existing vacant one is kept so its former holder is still removed.'
    });
  }
}

let _existingGroupEmailsCache = null;

/**
 * Set of every group address in the tenant, lowercased, built once per run.
 *
 * Read lazily — nothing calls it unless a seeded DL comes back empty, which on a
 * settled wing is rare. A failure to list is not fatal: an unseen group is
 * treated as absent, so the run creates nothing it should not and simply skips
 * clearing a vacant DL until the next run.
 *
 * @returns {Object<string, boolean>} group email -> true
 */
function getExistingGroupEmails_() {
  if (_existingGroupEmailsCache) return _existingGroupEmailsCache;

  const emails = {};
  let pageToken = '';

  try {
    do {
      const page = AdminDirectory.Groups.list({
        customer: 'my_customer',
        maxResults: 200,
        fields: 'groups(email),nextPageToken',
        pageToken: pageToken || undefined
      });
      (page.groups || []).forEach(g => {
        const email = String((g && g.email) || '').trim().toLowerCase();
        if (email) emails[email] = true;
      });
      pageToken = page.nextPageToken || '';
    } while (pageToken);

    Logger.info('Existing group index built', { count: Object.keys(emails).length });
  } catch (e) {
    Logger.warn('Failed to list existing groups; vacant echelon DLs are skipped this run', {
      errorMessage: e.message
    });
  }

  _existingGroupEmailsCache = emails;
  return emails;
}

/**
 * @param {string} groupEmail
 * @returns {boolean} True if the group already exists in the tenant
 */
function managedGroupExists_(groupEmail) {
  return !!getExistingGroupEmails_()[String(groupEmail || '').trim().toLowerCase()];
}

const _capwatchTableCache_ = {};

/**
 * Reads a CAPWATCH file WITH its header, as an array of objects.
 *
 * parseFile() drops the header row, which leaves column meaning encoded as a
 * magic index — the habit that left Member.txt's Expiration column unverified at
 * index 16 for months. The PL_* tables are new here and have no such folklore, so
 * they are read by name from the start.
 *
 * @param {string} fileName - Without extension, e.g. 'PL_MemberPathCredit'
 * @returns {Array<Object<string,string>>} One object per data row
 */
function readCapwatchTable_(fileName) {
  if (_capwatchTableCache_[fileName]) return _capwatchTableCache_[fileName];

  let out = [];
  try {
    const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
    const files = folder.getFilesByName(fileName + '.txt');

    if (!files.hasNext()) {
      Logger.warn('CAPWATCH file not found', { fileName: fileName + '.txt' });
      _capwatchTableCache_[fileName] = out;
      return out;
    }

    const rows = Utilities.parseCsv(files.next().getBlob().getDataAsString());
    if (!rows || rows.length < 2) {
      Logger.warn('CAPWATCH file has no data rows', { fileName: fileName + '.txt' });
      _capwatchTableCache_[fileName] = out;
      return out;
    }

    const header = rows[0].map(h => String(h || '').trim());
    for (let i = 1; i < rows.length; i++) {
      const obj = {};
      for (let c = 0; c < header.length; c++) {
        if (header[c]) obj[header[c]] = rows[i][c] === undefined ? '' : String(rows[i][c]).trim();
      }
      out.push(obj);
    }
  } catch (e) {
    Logger.warn('Failed to read CAPWATCH file', { fileName: fileName + '.txt', errorMessage: e.message });
  }

  _capwatchTableCache_[fileName] = out;
  return out;
}

/**
 * CAPID -> { PathID: true } for every APPROVED PL path credit.
 *
 * "Approved" is read from PL_Lookup rather than hardcoded: the StatusID that means
 * APPROVED is 8 today, and the lookup table is right there in the extract.
 *
 * @returns {Object<string, Object<string, boolean>>}
 */
function getApprovedPathCreditsByCapid_() {
  const approvedIds = {};
  readCapwatchTable_('PL_Lookup').forEach(row => {
    if (String(row.LookupType || '').trim() === 'ApprovalStatus' &&
        String(row.LookupValue || '').trim().toUpperCase() === 'APPROVED' &&
        row.LookupID) {
      approvedIds[String(row.LookupID)] = true;
    }
  });

  if (!Object.keys(approvedIds).length) {
    approvedIds['8'] = true;
    Logger.warn('PL_Lookup had no APPROVED ApprovalStatus row; falling back to StatusID 8');
  }

  const byCapid = {};
  let approvedRows = 0;
  readCapwatchTable_('PL_MemberPathCredit').forEach(row => {
    const capid = String(row.CAPID || '').trim();
    const pathId = String(row.PathID || '').trim();
    if (!capid || !pathId) return;
    if (!approvedIds[String(row.StatusID || '').trim()]) return;

    if (!byCapid[capid]) byCapid[capid] = {};
    byCapid[capid][pathId] = true;
    approvedRows++;
  });

  Logger.info('PL path credits loaded', {
    members: Object.keys(byCapid).length,
    approvedRows: approvedRows,
    approvedStatusIds: Object.keys(approvedIds).join(', ')
  });

  return byCapid;
}

/**
 * The level ladder: level number -> every PL path that level is made of.
 *
 * Read out of PL_Paths by name, so "Level 2 Part 1" and "Level 2 Part 2" both land
 * under 2 and a level added or re-split upstream is picked up without a code change.
 * Paths that are not levels (TLC, commander training, cadet achievements) are not
 * rungs and are absent here.
 *
 * @returns {Object<string, string[]>} e.g. { '1': ['4'], '2': ['7','8'], '3': ['3'] }
 */
function getProfessionalLevelLadder_() {
  const ladder = {};

  readCapwatchTable_('PL_Paths').forEach(row => {
    const id = String(row.PathID || '').trim();
    if (!id) return;

    const match = normalizeAchievementLabel_(row.PathName).match(/^level (\d+)( part \d+)?$/);
    if (!match) return;

    const level = match[1];
    if (!ladder[level]) ladder[level] = [];
    if (ladder[level].indexOf(id) === -1) ladder[level].push(id);
  });

  Object.keys(ladder).forEach(level => ladder[level].sort());
  return ladder;
}

/**
 * Where one member stands on the ladder.
 *
 * @param {Object<string, boolean>} earned - Approved PathIDs for this member
 * @param {Object<string, string[]>} ladder - From getProfessionalLevelLadder_()
 * @returns {{completed: string[], partial: string[], highest: string}} `highest` is
 *   '' for a member who has completed no level; `partial` holds levels with some
 *   parts approved but not all.
 */
function summarizeMemberLevelStanding_(earned, ladder) {
  const completed = [];
  const partial = [];
  let highest = 0;

  Object.keys(ladder).forEach(level => {
    const pathIds = ladder[level];
    const have = pathIds.filter(id => earned[id]).length;

    if (have === pathIds.length) {
      completed.push(level);
      if (Number(level) > highest) highest = Number(level);
    } else if (have > 0) {
      partial.push(level);
    }
  });

  return {
    completed: completed,
    partial: partial,
    highest: highest ? String(highest) : ''
  };
}

/**
 * Resolves Values entries into level rungs and plain paths.
 *
 * A value naming a level — "Level 2", "Level II", or the PathID of any path that
 * level is made of — becomes a rung, matched against the member's HIGHEST completed
 * level. Anything else that names a real path ("TLC Basic") stays a plain
 * "holds this path" requirement, since a non-level has no rung to be highest.
 *
 * @param {string[]} values - Trimmed Values entries from the Groups sheet
 * @param {string} groupName - Base group name, for logging
 * @param {boolean} [forInProgress=false] - Warn about values that cannot be partial
 * @returns {{levels: string[], paths: Array<{label: string, pathIds: string[]}>}}
 */
function resolveProfessionalLevelSpec_(values, groupName, forInProgress = false) {
  const paths = readCapwatchTable_('PL_Paths');
  const ladder = getProfessionalLevelLadder_();

  const byId = {};
  const byNormalizedName = {};
  const levelByPathId = {};

  paths.forEach(row => {
    const id = String(row.PathID || '').trim();
    const name = String(row.PathName || '').trim();
    if (!id) return;
    byId[id] = name;
    const norm = normalizeAchievementLabel_(name);
    if (norm) byNormalizedName[norm] = id;
  });

  Object.keys(ladder).forEach(level => {
    ladder[level].forEach(id => { levelByPathId[id] = level; });
  });

  const spec = { levels: [], paths: [] };
  const unresolved = [];
  const notSplittable = [];

  const addLevel = (level) => {
    if (spec.levels.indexOf(level) === -1) spec.levels.push(level);
    if (forInProgress && (ladder[level] || []).length < 2) notSplittable.push('Level ' + level);
  };

  for (let i = 0; i < values.length; i++) {
    const value = String(values[i] || '').trim();
    if (!value) continue;

    const norm = normalizeAchievementLabel_(value);

    // "Level 2" / "Level II"
    const named = norm.match(/^level (\d+)$/);
    if (named && ladder[named[1]]) {
      addLevel(named[1]);
      continue;
    }

    // A PathID, or a full path name. Either becomes a rung when it belongs to one.
    const pathId = (/^\d+$/.test(value) && byId[value]) ? value : byNormalizedName[norm];
    if (pathId) {
      if (levelByPathId[pathId]) addLevel(levelByPathId[pathId]);
      else spec.paths.push({ label: byId[pathId] || value, pathIds: [pathId] });
      continue;
    }

    unresolved.push(value);
  }

  if (unresolved.length) {
    Logger.warn('Professional level values could not be resolved to a PL path', {
      groupName: groupName,
      unresolved: unresolved.join(', '),
      pathsIndexed: paths.length,
      note: paths.length
        ? 'Value matches no PathName in PL_Paths.txt. Run listProfessionalLevelPaths() to see the real names.'
        : 'PL_Paths.txt is missing or empty in the CAPWATCH folder.'
    });
  }

  if (notSplittable.length) {
    Logger.warn('Level has no parts, so it has no in-progress state', {
      groupName: groupName,
      levels: notSplittable.join(', '),
      note: 'CAPWATCH stores this level as a single path — a member either holds it or does not.'
    });
  }

  return spec;
}

/**
 * Read-only: prints the PL paths CAPWATCH knows about, so a Groups sheet Values
 * column can be written against the real names.
 *
 * @param {string} [filter] - Case-insensitive substring, e.g. 'level'
 * @returns {void}
 */
function listProfessionalLevelPaths(filter) {
  const needle = String(filter || '').trim().toLowerCase();
  const paths = readCapwatchTable_('PL_Paths');

  if (!paths.length) {
    console.log('PL_Paths.txt not found or empty in the CAPWATCH data folder.');
    return;
  }

  console.log(`${paths.length} PL paths${needle ? ` matching "${needle}"` : ''}:\n`);
  paths
    .filter(row => !needle || String(row.PathName || '').toLowerCase().includes(needle))
    .sort((a, b) => String(a.PathName).localeCompare(String(b.PathName)))
    .forEach(row => console.log(`  PathID ${String(row.PathID).padEnd(5)} ${row.PathName}`));
}

/**
 * Resolves a Groups sheet Values list to the numeric AchvIDs MbrAchievements uses.
 *
 * MbrAchievements.txt identifies an achievement only by its numeric AchvID; the
 * human name lives in Achievements.txt. Both forms are accepted here so a Values
 * column can read "Level II" instead of a magic number, which is what the
 * Education & Training level rows were written with.
 *
 * Names are matched case-insensitively, with whitespace and punctuation collapsed
 * and standalone Roman numerals folded to digits, so "Level II", "level ii" and
 * "Level 2" all reach the same achievement.
 *
 * @param {string[]} values - Trimmed Values entries from the Groups sheet
 * @param {string} groupName - Base group name, for logging
 * @returns {Object<string, string>} AchvID -> the value that resolved to it
 */
function resolveAchievementValuesToIds_(values, groupName) {
  const ids = {};
  const unresolved = [];
  const labelIndex = getAchievementIdsByLabel_();

  for (let i = 0; i < values.length; i++) {
    const value = String(values[i] || '').trim();
    if (!value) continue;

    if (/^\d+$/.test(value)) {
      ids[value] = value;
      continue;
    }

    const matchedId = labelIndex[normalizeAchievementLabel_(value)];
    if (matchedId) {
      ids[matchedId] = value;
    } else {
      unresolved.push(value);
    }
  }

  if (unresolved.length) {
    Logger.warn('Achievement values could not be resolved to an AchvID', {
      groupName: groupName,
      unresolved: unresolved.join(', '),
      achievementsIndexed: Object.keys(labelIndex).length,
      note: Object.keys(labelIndex).length
        ? 'Value matches no name in Achievements.txt. Run listAchievementNames() to see the real strings.'
        : 'Achievements.txt is missing or has no recognizable name column, so only numeric AchvIDs can be used.'
    });
  }

  return ids;
}

let _achievementIdsByLabelCache = null;

/**
 * Builds a normalized achievement-name -> AchvID index from Achievements.txt.
 *
 * Read with its header rather than through parseFile(), which drops the header
 * row and so leaves no way to tell which column carries the name — CAPWATCH has
 * used more than one spelling for it.
 *
 * @returns {Object<string, string>} normalized label -> AchvID
 */
function getAchievementIdsByLabel_() {
  if (_achievementIdsByLabelCache) return _achievementIdsByLabelCache;

  const index = {};

  try {
    const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
    const files = folder.getFilesByName('Achievements.txt');

    if (!files.hasNext()) {
      Logger.warn('Achievements.txt not found - achievement names cannot be resolved', {
        folderId: CONFIG.CAPWATCH_DATA_FOLDER_ID
      });
      _achievementIdsByLabelCache = index;
      return index;
    }

    const rows = Utilities.parseCsv(files.next().getBlob().getDataAsString());
    if (!rows || rows.length < 2) {
      _achievementIdsByLabelCache = index;
      return index;
    }

    const header = rows[0].map(h => String(h || '').trim());
    const idIdx = header.findIndex(h => /^achvid$/i.test(h));
    const labelIdxs = header
      .map((h, i) => (/^(achv|achvname|name|text|title|description)$/i.test(h) ? i : -1))
      .filter(i => i > -1);

    if (idIdx === -1 || !labelIdxs.length) {
      Logger.warn('Achievements.txt has no usable AchvID/name columns', {
        header: header.join(', ')
      });
      _achievementIdsByLabelCache = index;
      return index;
    }

    for (let i = 1; i < rows.length; i++) {
      const achvId = String(rows[i][idIdx] || '').trim();
      if (!achvId) continue;

      for (let c = 0; c < labelIdxs.length; c++) {
        const label = normalizeAchievementLabel_(rows[i][labelIdxs[c]]);
        // First writer wins: an achievement's own name column beats a longer
        // description that happens to normalize to the same string.
        if (label && !index[label]) index[label] = achvId;
      }
    }

    Logger.info('Achievement name index built', {
      achievements: rows.length - 1,
      labels: Object.keys(index).length
    });
  } catch (e) {
    Logger.warn('Failed to index Achievements.txt', { errorMessage: e.message });
  }

  _achievementIdsByLabelCache = index;
  return index;
}

/**
 * Normalizes an achievement name for comparison: case, whitespace, punctuation,
 * and Roman numerals ("Level II" and "Level 2" are the same achievement).
 *
 * @param {string} value
 * @returns {string}
 */
function normalizeAchievementLabel_(value) {
  const romanToArabic = { i: '1', ii: '2', iii: '3', iv: '4', v: '5' };

  return String(value == null ? '' : value)
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .trim()
    .replace(/\s+/g, ' ')
    .split(' ')
    .map(token => (Object.prototype.hasOwnProperty.call(romanToArabic, token) ? romanToArabic[token] : token))
    .join(' ');
}

/**
 * Read-only: prints the achievements CAPWATCH knows about, so a Groups sheet
 * Values column can be written against the real strings.
 *
 * @param {string} [filter] - Case-insensitive substring, e.g. 'level'
 * @returns {void}
 */
function listAchievementNames(filter) {
  const needle = String(filter || '').trim().toLowerCase();
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files = folder.getFilesByName('Achievements.txt');

  if (!files.hasNext()) {
    console.log('Achievements.txt not found in the CAPWATCH data folder.');
    return;
  }

  const rows = Utilities.parseCsv(files.next().getBlob().getDataAsString());
  const header = (rows[0] || []).map(h => String(h || '').trim());
  const idIdx = header.findIndex(h => /^achvid$/i.test(h));
  const labelIdxs = header
    .map((h, i) => (/^(achv|achvname|name|text|title|description)$/i.test(h) ? i : -1))
    .filter(i => i > -1);

  console.log(`Achievements.txt columns: ${header.join(', ')}`);
  console.log(`${rows.length - 1} achievements${needle ? ` matching "${needle}"` : ''}:\n`);

  for (let i = 1; i < rows.length; i++) {
    const labels = labelIdxs.map(c => String(rows[i][c] || '').trim()).filter(Boolean);
    const line = `  AchvID ${String(rows[i][idIdx] || '').padEnd(6)} ${labels.join(' | ')}`;
    if (!needle || line.toLowerCase().includes(needle)) console.log(line);
  }
}

/**
 * Saves email groups data to file for tracking and debugging
 * @param {Object} emailGroups - Groups object with member emails
 * @returns {void}
 */
function saveEmailGroups(emailGroups) {
  let folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  let files = folder.getFilesByName('EmailGroups.txt');

  if (files.hasNext()) {
    let file = files.next();
    let content = JSON.stringify(emailGroups);
    file.setContent(content);
    Logger.info('Email groups saved', {
      fileName: 'EmailGroups.txt',
      categories: Object.keys(emailGroups).length
    });
  } else {
    Logger.warn('EmailGroups.txt file not found', {
      folderId: CONFIG.CAPWATCH_DATA_FOLDER_ID
    });
  }
}

/**
 * Saves problematic email addresses to spreadsheet for manual review
 * Includes detailed error information, CAPID mapping, and multiple attempts per email
 * @param {Object} errorEmails - Object mapping email addresses to error details
 * @returns {void}
 */
function saveErrorEmails(errorEmails) {
  if (Object.keys(errorEmails).length === 0) {
    Logger.info('No error emails to save');
    return;
  }

  try {
    // Map emails to CAPIDs
    const contacts = parseFile('MbrContact');
    const emailMap = contacts.reduce(function(map, obj) {
      const cleanEmail = (obj[3] || '').trim().toLowerCase();
      if (cleanEmail) {
        map[cleanEmail] = obj[0];
      }
      return map;
    }, {});

    const sheet = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID)
      .getSheetByName('Error Emails');

    // Clear existing data (keep header row)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }

    // Set up headers if not present
    const headers = sheet.getRange(1, 1, 1, 9).getValues()[0];
    if (headers[0] !== 'Email' || headers.length < 9) {
      sheet.getRange(1, 1, 1, 9).setValues([[
        'Email',
        'CAPID',
        'Error Count',
        'Groups Affected',
        'Error Codes',
        'Last Error Message',
        'Categories',
        'First Seen',
        'Last Seen'
      ]]);

      // Format header row
      sheet.getRange(1, 1, 1, 9)
        .setFontWeight('bold')
        .setBackground('#4285f4')
        .setFontColor('#ffffff');
    }

    // Build rows with detailed information
    const values = [];

    for (const email in errorEmails) {
      const errorInfo = errorEmails[email];
      const attempts = errorInfo.attempts || [];

      if (attempts.length === 0) continue;

      // Extract unique values from attempts
      const groups = [...new Set(attempts.map(a => a.group))].join(', ');
      const errorCodes = [...new Set(attempts.map(a => a.errorCode))].join(', ');
      const categories = [...new Set(attempts.map(a => a.category))].join(', ');

      // Get last error message
      const lastAttempt = attempts[attempts.length - 1];
      const lastErrorMessage = lastAttempt.errorMessage || 'Unknown';

      // Get timestamps
      const firstSeen = errorInfo.firstSeen || attempts[0].timestamp || 'Unknown';
      const lastSeen = lastAttempt.timestamp || 'Unknown';

      // Format dates for spreadsheet
      const firstSeenDate = firstSeen !== 'Unknown' ? new Date(firstSeen) : 'Unknown';
      const lastSeenDate = lastSeen !== 'Unknown' ? new Date(lastSeen) : 'Unknown';

      // Look up CAPID
      const capid = emailMap[email.toLowerCase()] || 'Unknown';

      values.push([
        email,
        capid,
        attempts.length,
        groups,
        errorCodes,
        lastErrorMessage,
        categories,
        firstSeenDate,
        lastSeenDate
      ]);
    }

    // Sort by error count (descending) then by email
    values.sort((a, b) => {
      if (b[2] !== a[2]) return b[2] - a[2]; // Sort by error count
      return a[0].localeCompare(b[0]); // Then by email
    });

    // Write to spreadsheet
    if (values.length > 0) {
      sheet.getRange(2, 1, values.length, 9).setValues(values);

      // Format the data
      const dataRange = sheet.getRange(2, 1, values.length, 9);
      dataRange.setVerticalAlignment('top');

      // Format date columns
      if (values.length > 0) {
        sheet.getRange(2, 8, values.length, 2).setNumberFormat('yyyy-mm-dd hh:mm:ss');
      }

      // Add conditional formatting for error count
      const errorCountRange = sheet.getRange(2, 3, values.length, 1);
      const rules = sheet.getConditionalFormatRules();

      // High errors (5+) = Red
      const redRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(5)
        .setBackground('#f4cccc')
        .setRanges([errorCountRange])
        .build();

      // Medium errors (2-4) = Yellow
      const yellowRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(2, 4)
        .setBackground('#fff2cc')
        .setRanges([errorCountRange])
        .build();

      rules.push(redRule);
      rules.push(yellowRule);
      sheet.setConditionalFormatRules(rules);

      // Auto-resize columns
      for (let i = 1; i <= 9; i++) {
        sheet.autoResizeColumn(i);
      }

      Logger.info('Error emails saved to spreadsheet', {
        count: values.length,
        totalAttempts: values.reduce((sum, row) => sum + row[2], 0),
        sheetName: 'Error Emails'
      });
    }

  } catch (e) {
    Logger.error('Failed to save error emails', {
      errorMessage: e.message,
      errorCount: Object.keys(errorEmails).length
    });
  }
}

/**
 * Retrieves current members of a Google Group
 * Creates the group if it doesn't exist
 * @param {string} groupId - Group identifier (without domain)
 * @param {Object} squadrons - Squadrons object indexed by orgid
 * @param {{name?: string, description?: string}} [meta] - Desired metadata for the group
 * @returns {{email:string, role:string}[]} Array of current group members with roles
 */
function getCurrentGroup(groupId, squadrons, meta = {}, allowExternalMembers = false) {
  const email = groupId + CONFIG.EMAIL_DOMAIN;
  let members = [];
  let nextPageToken = '';

  try {
    do {
      let page = AdminDirectory.Members.list(email, {
        maxResults: GROUP_MEMBER_PAGE_SIZE,
        pageToken: nextPageToken
      });
      if (page.members) {
        members = members.concat(page.members.map(function(member) {
          return {
            email: (member.email || '').toLowerCase(),
            role: (member.role || 'MEMBER').toString().toUpperCase()
          };
        }));
      }
      nextPageToken = page.nextPageToken;
    } while(nextPageToken);

  } catch(e) {
    if (e.details?.code === ERROR_CODES.NOT_FOUND) {
      // Group not found - create it (dry-run aware)
      try {
        // Prefer the already-computed desired metadata so new groups are created
        // with the same final name/description existing groups are patched to.
        let finalName = (meta && meta.name ? meta.name : '').toString().trim();
        let finalDescription = (meta && meta.description ? meta.description : '').toString().trim();

        if (!finalDescription) {
          const org = Object.values(squadrons).find(o => groupId.includes(o.unit));
          const baseName = groupId.split('.').slice(1).join('.');
          const orgName = org ? org.name.toLowerCase().replace(/\b\w/g, c => c.toUpperCase()) : '';
          const formattedGroupName = baseName.replace(/-/g, '.');
          finalDescription = org ? `${orgName} – ${formattedGroupName}` : formattedGroupName;
        }

        if (!finalName) {
          finalName = finalDescription || groupId;
        }

        if (DRY_RUN) {
          Logger.info('💡 [Dry-Run] Would create group', {
            groupId: groupId,
            name: finalName,
            description: finalDescription,
            email: groupId + CONFIG.EMAIL_DOMAIN,
            allowExternalMembers: allowExternalMembers
          });
          return [];
        } else {
          let newGroup = AdminDirectory.Groups.insert({
            email: groupId + CONFIG.EMAIL_DOMAIN,
            name: finalName,
            description: finalDescription
          });

          // Apply "Allow external members" when requested by Groups sheet Add EXT column.
          if (allowExternalMembers) {
            applyAllowExternalMembersSetting_(newGroup.email || (groupId + CONFIG.EMAIL_DOMAIN), true);
          }

          Logger.info('Group created', {
            groupEmail: newGroup.email,
            name: newGroup.name,
            description: newGroup.description,
            allowExternalMembers: allowExternalMembers
          });
        }
      } catch(createError) {
        Logger.error('Failed to create group', {
          groupId: groupId,
          errorMessage: createError.message,
          errorCode: createError.details?.code
        });
      }
    } else {
      Logger.error('Error retrieving group members', {
        groupId: groupId,
        errorMessage: e.message,
        errorCode: e.details?.code
      });
    }
  }

  return members;
}

/**
 * Best-effort apply of the Google Group setting "allowExternalMembers".
 * Requires the Advanced Google Service "Admin SDK Groups Settings API"
 * (AdminGroupsSettings) to be enabled.
 *
 * @param {string} groupEmail
 * @param {boolean} allowExternalMembers
 * @returns {void}
 */
function applyAllowExternalMembersSetting_(groupEmail, allowExternalMembers) {
  try {
    if (typeof AdminGroupsSettings === 'undefined' || !AdminGroupsSettings.Groups || !AdminGroupsSettings.Groups.patch) {
      Logger.warn('AdminGroupsSettings API not available; cannot set allowExternalMembers', {
        group: groupEmail,
        allowExternalMembers: allowExternalMembers
      });
      return;
    }

    executeWithRetry(() =>
      AdminGroupsSettings.Groups.patch({
        allowExternalMembers: allowExternalMembers ? 'true' : 'false'
      }, groupEmail)
    );

    Logger.info('Applied allowExternalMembers setting', {
      group: groupEmail,
      allowExternalMembers: allowExternalMembers
    });
  } catch (e) {
    Logger.warn('Failed to apply allowExternalMembers setting', {
      group: groupEmail,
      allowExternalMembers: allowExternalMembers,
      errorMessage: e.message
    });
  }
}

/**
 * Best-effort apply of managed Google Group settings for UpdateGroups-managed groups.
 * Currently used to keep membership visibility open to the whole domain while
 * preserving the per-group allowExternalMembers behavior from the Groups sheet.
 *
 * @param {string} groupEmail
 * @param {{allowExternalMembers?: boolean, whoCanViewMembership?: string,
 *          whoCanPostMessage?: string, spamModerationLevel?: string}} settings
 * @returns {void}
 */
function applyManagedGroupSettings_(groupEmail, settings) {
  try {
    const desired = {};

    if (typeof settings.allowExternalMembers === 'boolean') {
      desired.allowExternalMembers = settings.allowExternalMembers ? 'true' : 'false';
    }
    if (settings.whoCanViewMembership) {
      desired.whoCanViewMembership = settings.whoCanViewMembership;
    }
    if (settings.whoCanPostMessage) {
      desired.whoCanPostMessage = settings.whoCanPostMessage;
    }
    // ANYONE_CAN_POST is open to the internet, so it travels with moderation.
    if (settings.spamModerationLevel) {
      desired.spamModerationLevel = settings.spamModerationLevel;
    }

    if (Object.keys(desired).length === 0) return;

    if (DRY_RUN) {
      Logger.info('💡 [Dry-Run] Would apply managed group settings', {
        group: groupEmail,
        settings: desired
      });
      return;
    }

    if (typeof AdminGroupsSettings === 'undefined' || !AdminGroupsSettings.Groups || !AdminGroupsSettings.Groups.patch) {
      Logger.warn('AdminGroupsSettings API not available; cannot apply managed group settings', {
        group: groupEmail
      });
      return;
    }

    const existing = AdminGroupsSettings.Groups.get(groupEmail);
    const patch = {};

    for (const key in desired) {
      const currentValue = (existing && existing[key] != null) ? String(existing[key]) : '';
      const desiredValue = String(desired[key]);
      if (currentValue !== desiredValue) {
        patch[key] = desiredValue;
      }
    }

    if (Object.keys(patch).length === 0) {
      Logger.info('Managed group settings already correct', {
        group: groupEmail,
        settings: desired
      });
      return;
    }

    executeWithRetry(() => AdminGroupsSettings.Groups.patch(patch, groupEmail));
    Logger.info('Applied managed group settings', {
      group: groupEmail,
      settings: patch
    });
  } catch (e) {
    Logger.warn('Failed to apply managed group settings', {
      group: groupEmail,
      errorMessage: e.message
    });
  }
}

/**
 * Adds additional members to groups based on manual spreadsheet entries
 * Supports MEMBER, MANAGER, and OWNER roles
 * Does not automatically remove members
 * @returns {void}
 */
function updateAdditionalGroupMembers() {
  const start = new Date();
  let additionalMembers = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID)
    .getSheetByName('User Additions')
    .getDataRange()
    .getValues();
  let errorEmails = {};
  const roles = ['MEMBER', 'MANAGER', 'OWNER'];
  let added = 0;
  let skipped = 0;
  let errors = 0;

  for(let i = 1; i < additionalMembers.length; i++) {
    let groups = additionalMembers[i][3].split(',');
    for(let j = 0; j < groups.length; j++) {
      let groupEmail = groups[j].trim() + CONFIG.EMAIL_DOMAIN;
      let email = additionalMembers[i][1];
      let role = additionalMembers[i][2].toLocaleUpperCase();

      if (roles.indexOf(role) < 0) {
        Logger.warn('Invalid role in spreadsheet - skipping', {
          email: email,
          invalidRole: role,
          validRoles: roles.join(', '),
          row: i + 1
        });
        skipped++;
        continue;
      }

      // Add member to group
      try {
        executeWithRetry(() =>
          AdminDirectory.Members.insert({
            email: email,
            role: role
          }, groupEmail)
        );
        Logger.info('Additional member added to group', {
          email: email,
          group: groupEmail,
          role: role
        });
        added++;

      } catch (e) {
        if (e.details?.code === ERROR_CODES.CONFLICT) {
          Logger.info('Member already in group', {
            email: email,
            group: groupEmail,
            role: role
          });
          skipped++;
        } else {
          Logger.error('Failed to add additional member', {
            email: email,
            group: groupEmail,
            role: role,
            row: i + 1,
            errorMessage: e.message,
            errorCode: e.details?.code
          });
          errors++;

          if ([ERROR_CODES.BAD_REQUEST, ERROR_CODES.NOT_FOUND].indexOf(e.details?.code) > -1) {
            // Track detailed error info
            if (!errorEmails[email]) {
              errorEmails[email] = {
                email: email,
                attempts: [],
                firstSeen: new Date().toISOString()
              };
            }
            errorEmails[email].attempts.push({
              group: groups[j].trim(),
              groupEmail: groupEmail,
              category: 'additional-members',
              errorCode: e.details?.code || 'Unknown',
              errorMessage: e.message || 'Unknown error',
              timestamp: new Date().toISOString()
            });
          }
        }
      }
    }
  }

  Logger.info('Additional group members processed', {
    duration: new Date() - start + 'ms',
    added: added,
    skipped: skipped,
    errors: errors,
    errorEmailsCount: Object.keys(errorEmails).length
  });
}

/**
 * Test function for saveErrorEmails
 * @returns {void}
 */
function testSaveErrorEmails() {
  let errorEmails = {
    'bob.rodenhouse@gmail.com': 'test-group-1',
    'mi190.sdavis@live.com': 'test-group-2',
    'michael-shoemaker@sbcglobal.net': 'test-group-3'
  };
  saveErrorEmails(errorEmails);
}

function testEnhancedErrorTracking() {
   // Create test error structure
   const testErrors = {
     'test1@gmail.com': {
       email: 'test1@gmail.com',
       firstSeen: new Date().toISOString(),
       attempts: [
         {
           group: 'test-group-1',
           groupEmail: `test-group-1${CONFIG.EMAIL_DOMAIN}`,
           category: 'test-category',
           errorCode: 404,
           errorMessage: 'Test error message 1',
           timestamp: new Date().toISOString()
         },
         {
           group: 'test-group-2',
           groupEmail: `test-group-2${CONFIG.EMAIL_DOMAIN}`,
           category: 'test-category-2',
           errorCode: 400,
           errorMessage: 'Test error message 2',
           timestamp: new Date().toISOString()
         }
       ]
     },
     'test2@example.com': {
       email: 'test2@example.com',
       firstSeen: new Date().toISOString(),
       attempts: [
         {
           group: 'test-group-3',
           groupEmail: `test-group-3${CONFIG.EMAIL_DOMAIN}`,
           category: 'test-category-3',
           errorCode: 404,
           errorMessage: 'Test error message 3',
           timestamp: new Date().toISOString()
         }
       ]
     }
   };

   saveErrorEmails(testErrors);
   Logger.info('Test completed - check Error Emails sheet');
 }

function debugGroupsTabSource() {
  const spreadsheetId = String(CONFIG.AUTOMATION_SPREADSHEET_ID || '').trim();
  const sheetName = 'Groups';

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Missing sheet: ${sheetName}`);
  }

  const rows = sheet.getDataRange().getValues();
  const header = rows[0] || [];

  Logger.info('Runtime automation source', {
    spreadsheetId: spreadsheetId,
    spreadsheetName: ss.getName(),
    sheetName: sheetName,
    totalRows: rows.length
  });

  Logger.info('Groups header', {
    c1: header[0],
    c2: header[1],
    c3: header[2],
    c4: header[3],
    c5: header[4],
    c6: header[5]
  });

  const matches = [];
  const exactDtyAll = [];
  const staffRows = [];

  for (let i = 1; i < rows.length; i++) {
    const rowNum = i + 1;
    const category = String(rows[i][0] || '').trim();
    const groupName = String(rows[i][1] || '').trim();
    const attribute = String(rows[i][2] || '').trim();
    const values = String(rows[i][3] || '').trim();
    const description = String(rows[i][4] || '').trim();
    const addExt = String(rows[i][5] || '').trim();

    const rowObj = {
      row: rowNum,
      category: category,
      groupName: groupName,
      attribute: attribute,
      values: values,
      description: description,
      addExt: addExt
    };

    if (groupName === 'dty.all') {
      exactDtyAll.push(rowObj);
    }

    if (
      groupName.indexOf('dty.') > -1 ||
      attribute === 'dutyPositionLevelStaff'
    ) {
      matches.push(rowObj);
    }

    if (
      groupName === 'dty.wing-stf-only' ||
      groupName === 'dty.grp-stf-only' ||
      attribute === 'dutyPositionLevelStaff'
    ) {
      staffRows.push(rowObj);
    }
  }

  Logger.info('Exact dty.all rows', {
    count: exactDtyAll.length,
    rows: exactDtyAll
  });

  Logger.info('Staff rows of interest', {
    count: staffRows.length,
    rows: staffRows
  });

  Logger.info('All duty-position style rows', {
    count: matches.length,
    rows: matches
  });
}
