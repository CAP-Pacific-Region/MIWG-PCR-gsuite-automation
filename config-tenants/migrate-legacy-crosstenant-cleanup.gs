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
  let matched = 0, deleted = 0, failed = 0;

  LEGACY_MARKERS.forEach(function (marker) {
    const rows = xtListManagedContacts_(cfg, marker);
    matched += rows.length;
    console.log('Legacy marker "' + marker + '": ' + rows.length + ' contacts in ' + cfg.selfDomain);

    rows.forEach(function (r, i) {
      if (DRY_RUN) {
        if (i < 15) console.log('  would delete: ' + (r.email || r.resourceId));
        return;
      }
      const res = xtDeleteContact_(r, token);
      if (res.ok) { deleted++; }
      else { failed++; console.log('  DELETE FAILED (' + res.code + '): ' + (r.email || r.resourceId)); }
      if ((i + 1) % 50 === 0) Utilities.sleep(200); // gentle throttle
    });
  });

  console.log((DRY_RUN ? 'DRY RUN — ' : '') + 'Legacy cleanup complete. matched=' +
    matched + ', deleted=' + deleted + ', failed=' + failed);
  return { dryRun: DRY_RUN, matched: matched, deleted: deleted, failed: failed };
}
