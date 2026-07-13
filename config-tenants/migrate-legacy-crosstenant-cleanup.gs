/**
 * ONE-TIME migration cleanup for the cross-tenant shared-contacts move.
 * Paste into the NEW project's Apps Script editor (the one now running
 * src/cross-tenant-contacts), select cleanupLegacyCrossTenantContacts, and Run
 * once PER TENANT (seniors, then cadets). Repo-only: .claspignore ships only
 * src/**, so this file is never pushed — it is a manual paste-once tool.
 *
 * WHY: the retired standalone projects tagged their managed Domain Shared
 * Contacts with orgName = 'CAWG' (members) and 'CAWG_CADET_PARENTS_GROUPS'
 * (parent groups). The new module uses different markers (CONFIG.WING = 'CA'
 * and 'CA_PARENTS'), so those legacy contacts would linger as duplicates. This
 * deletes them. Because the markers differ, it can only ever match legacy
 * contacts — never the new module's.
 *
 * PRECONDITIONS
 *   1. Retire the legacy cadets (1fJRqo...) and seniors (1b2JSIB...) projects —
 *      remove their triggers — so they stop recreating contacts.
 *   2. The shared src/ (with the m8/feeds scope + the new module) is pushed and
 *      the project has been re-consented.
 *
 * PROCEDURE
 *   - Leave DRY_RUN = true, Run, and review the logged count + sample.
 *   - Set DRY_RUN = false, Run again to delete.
 *   - The m8 feed has a per-day write cap (~few thousand). A large backlog may
 *     hit it: the run logs "Aborting ... quota exhausted" or "Reached
 *     MAX_DELETES_PER_RUN" and stops cleanly. Just re-run after the daily reset
 *     (~midnight Pacific) — it re-lists and continues where it left off. Repeat
 *     until a DRY_RUN reports matched=0.
 *   - Delete this function from the project when done.
 *
 * Reuses xtListManagedContacts_() and xtDeleteContact_() from the pushed module.
 */
function cleanupLegacyCrossTenantContacts() {
  const DRY_RUN = true; // <-- set to false to actually delete
  const LEGACY_MARKERS = ['CAWG', 'CAWG_CADET_PARENTS_GROUPS'];

  // Safety: never run where this tenant's live marker equals a legacy marker
  // (that would mean the deletion could hit current contacts). CONFIG.WING is
  // 'CA' here, so this never trips for CAWG — it guards against misuse elsewhere.
  if (LEGACY_MARKERS.indexOf(String(CONFIG.WING)) !== -1 ||
      LEGACY_MARKERS.indexOf(String(CONFIG.WING) + '_PARENTS') !== -1) {
    throw new Error('Refusing to run: CONFIG.WING ("' + CONFIG.WING +
      '") collides with a legacy marker; deletion could remove current contacts.');
  }

  const cfg = { selfDomain: CONFIG.DOMAIN };
  const token = ScriptApp.getOAuthToken();

  // The m8 Domain Shared Contacts feed has a per-day write cap (~few thousand).
  // Cap deletes per run to stay under it, and bail out fast when sustained 429s
  // mean the quota is exhausted — re-run after the daily reset to finish.
  const MAX_DELETES_PER_RUN = 2000;
  const ABORT_AFTER_CONSECUTIVE_FAILS = 25;

  let matched = 0, deleted = 0, failed = 0, consecutiveFails = 0, aborted = false;

  for (let mi = 0; mi < LEGACY_MARKERS.length && !aborted; mi++) {
    const marker = LEGACY_MARKERS[mi];
    const rows = xtListManagedContacts_(cfg, marker);
    matched += rows.length;
    console.log('Legacy marker "' + marker + '": ' + rows.length + ' contacts in ' + cfg.selfDomain);

    for (let i = 0; i < rows.length; i++) {
      if (DRY_RUN) {
        if (i < 15) console.log('  would delete: ' + (rows[i].email || rows[i].resourceId));
        continue;
      }
      if (deleted >= MAX_DELETES_PER_RUN) {
        console.log('Reached MAX_DELETES_PER_RUN (' + MAX_DELETES_PER_RUN + '); re-run to continue.');
        aborted = true; break;
      }

      const res = xtDeleteContact_(rows[i], token);
      if (res.ok) {
        deleted++; consecutiveFails = 0;
        if (deleted % 100 === 0) console.log('  progress: deleted ' + deleted);
        if (deleted % 50 === 0) Utilities.sleep(300);
      } else {
        failed++; consecutiveFails++;
        console.log('  DELETE FAILED (' + res.code + '): ' + (rows[i].email || rows[i].resourceId));
        if (consecutiveFails >= ABORT_AFTER_CONSECUTIVE_FAILS) {
          console.log('Aborting: ' + consecutiveFails + ' consecutive failures (likely m8 daily write ' +
            'quota exhausted). Nothing is broken — re-run after the quota resets to finish.');
          aborted = true; break;
        }
      }
    }
  }

  console.log((DRY_RUN ? 'DRY RUN — ' : '') + 'Legacy cleanup ' + (aborted ? 'PAUSED' : 'complete') +
    '. matched=' + matched + ', deleted=' + deleted + ', failed=' + failed);
  return { dryRun: DRY_RUN, matched: matched, deleted: deleted, failed: failed, aborted: aborted };
}
