# PCR Changelog

Pacific Region (PCR) fork-specific changes to the CAPWATCH / Google Workspace
automation, layered on top of the upstream `cap-miwg/gsuite-automation` project.
Upstream changes are tracked in [CHANGELOG.md](CHANGELOG.md); **this file records
only what the Pacific Region deployment adds or diverges on.**

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).
Individual source files carry their own SemVer version in their header
(see [docs/VERSIONING.md](docs/VERSIONING.md)); the per-file version is noted
next to each entry below.

## [2026-07-11] — Squadron `.all` lists now admit cross-tenant cadet groups

### Fixed

- **`SquadronGroups.gs` (v1.2.9)** — squadron distribution lists (notably the
  `ca###.all` lists) were not receiving the setting that lets the cross-tenant
  cadet group `ca###.cadets@cawgcadets.org` be added as a member, so messages to
  a unit's **All** list never reached cadets. Root cause: `applyGroupSettings()`
  was a log-only stub that built the intended settings (including
  `allowExternalMembers: 'true'`) but never called any API — the header comment
  wrongly claimed "Apps Script doesn't have direct Groups Settings API access,"
  even though the `AdminGroupsSettings` advanced service is enabled and used
  elsewhere (`UpdateGroups.gs`, `groupAdministration.gs`). External-member adds
  therefore failed silently and were swallowed per-member in
  `updateGroupMembership()`. `applyGroupSettings()` now patches
  `allowExternalMembers` through `AdminGroupsSettings.Groups.patch`, only when the
  live value differs (idempotent, `DRY_RUN`-aware). Because `getOrCreateGroup()`
  runs it for existing groups too, the next `updateAllSquadronGroups()` backfills
  `allowExternalMembers=true` across all squadron lists — self-healing, no manual
  console work. Deployed to all three tenants.

### Scope note (why only `allowExternalMembers`)

- The fix deliberately enforces **only** `allowExternalMembers` (narrowed from an
  initial v1.2.8 that applied the whole settings block). The code passes
  `whoCanPostMessage: 'ALL_MEMBERS_CAN_POST'` for every distribution list, but
  that was never applied while the function was a stub — so the live cadet-tenant
  receive lists `ca###.cadets@cawgcadets.org` sit at `ANYONE_CAN_POST`, which is
  exactly what lets them accept mail fanned out from the wing `.all` lists.
  Enforcing the full block would have flipped those receivers to
  `ALL_MEMBERS_CAN_POST` and silently re-broken cadet delivery. Posting/visibility
  policy is therefore left to console/GAM.
- Audit (`groupAdministration_auditReceiveListPosting()`, run on the cadets
  tenant): `.cadets`/`.parents` receivers = `ANYONE_CAN_POST` (correct); the 66
  flagged `ca###.all@cawgcadets.org` are the cadet tenant's own internal
  all-hands at `ALL_IN_DOMAIN_CAN_POST` — not cross-tenant receivers, left as-is.

### Changed — squadron distribution toggles are now tenant-driven

- **`SquadronGroups.gs` (v1.3.0) + `config.gs` (v1.2.1)** — `SQUADRON_DISTRIBUTION_TOGGLES`
  was a hard-coded const, so the cadet tenant was creating senior-only lists
  (`.seniors`, Deputy Commander for Seniors) that don't apply there. The toggles
  now come from `PROFILE_.SQUADRON_DISTRIBUTION_TOGGLES` in `config.gs` (selected
  by the `TENANT_PROFILE` Script Property), read via
  `getSquadronDistributionToggles_()`; the const is a fallback default only. Same
  mechanism as the other per-tenant behavior — a shared-code `clasp push` can't
  make a tenant create the wrong lists.
- **Cadets profile = cadets + parents lists only.** Disabled on the cadet tenant:
  `.seniors` (no seniors here), `.all` (redundant with `.cadets` on a cadet-only
  tenant), and the command-staff lists (Commander / Deputy Commander / Deputy
  Commander for Cadets — those are senior duty positions whose holders have wing
  accounts, so the lists would be empty). This matches the pre-existing
  `SQUADRON_DISTRIBUTION_TYPES` and the `_behavioral_note` in
  `config-tenants/cadets.json`. Seniors profile unchanged; pacific = all off
  (single-unit region, squadron sync not triggered there).
- **Cleanup follow-up:** disabling a toggle stops managing those lists but does
  not delete already-created groups. The existing `ca###.seniors@cawgcadets.org`,
  `ca###.all@cawgcadets.org`, and cadet command-staff groups become orphans on
  the cadet tenant and should be removed (e.g. via
  `groupAdministration_bulkDeleteGroupsFromSheet`).

## [2026-07-09] — Pacific go-live

The reconciled `src/` was deployed to the live "PCR Automation" project (`TENANT_PROFILE=pacific`)
and verified end-to-end; triggers rebuilt under `automation@pcr.cap.gov`. **All three tenants now
run identical source, differentiated only by configuration** — the reconciliation goal.
_This supersedes the "not yet deployed to Pacific / on hold pending 2SV" notes in the entries
below, which were accurate when written._

### Fixed

- **`UnitVisitReport.gs` (v1.0.1)** — `buildRegionUnitVisitReport()` failed with
  "You can't merge frozen and non-frozen columns" in `buildWingTab_`.
  `clear()`/`clearFormats()` don't reset a tab's freeze state, so a pre-existing
  frozen column made the `A1:I1` title merge fail. Now resets frozen rows/columns
  before merging. (Surfaced during the Pacific go-live; PR #11.)

### Deployment notes

- **Push must come from an account internal to `pcr.cap.gov`** (`automation@pcr.cap.gov`, the
  owner). The project is in a `pcr.cap.gov` Shared Drive; an external-org account can pull but gets
  a 403 "Apps Script API not enabled" on push (a cross-org write block, not an API-toggle issue).
- The `contacts` OAuth scope was **verified working** against the M8 Domain Shared Contacts feed.
- Two pre-existing Google **abuse-suspended** accounts (`timothy.verrett`, `rene.mccoy`) can't be
  auto-restored (412); they need an admin restore in the console. Unrelated to the reconciliation.

## [2026-07-09] — Fold region modules into the shared `src/` (identical-code model)

All three tenants now run an identical `src/`; modules a tenant doesn't use are
disabled by per-feature profile flags rather than removed (see
[docs/PACIFIC_DIFF.md](docs/PACIFIC_DIFF.md)). **Behavior-preserving for the seniors
and cadets tenants** (region features flag off). **Not yet deployed to Pacific** —
deployment is on hold pending 2SV.

### Added

- **`src/region/UpdateRegionGroupChats.gs`** (v1.0.0) — region duty groups + duty chat
  spaces (`updateRegionGroupChats()`), gated by `RUN_REGION_GROUP_CHATS`.
- **`src/region/UnitVisitReport.gs`** (v1.0.0) — region-wide unit-visit spreadsheet
  (`buildRegionUnitVisitReport()`), gated by `RUN_UNIT_VISIT_REPORT`. Destination
  spreadsheet/calendar IDs read from Script Properties (no tenant literals).
- **`src/accounts-and-groups/SharedContacts.gs`** (v1.0.0) — "External Contacts" sheet →
  Domain Shared Contacts (`runExternalContactsToDomainSharedContacts()`), gated by
  `RUN_SHARED_CONTACTS`.
- Per-feature flags in `TENANT_PROFILES_` (all `false` for seniors/cadets, `true` for
  pacific); `REGION_CAPWATCH_DATA_FOLDER_ID` identity + `TENANT_UNIT_VISIT_*` properties.
- `https://www.googleapis.com/auth/contacts` OAuth scope (for shared contacts).

### Changed

- **`UpdateChatSpaces.gs` (v2.0.0)** — converged to the Pacific superset as the single
  shared module: adds automation-group + user-additions chat spaces (gated by
  `RUN_AUTOMATION_CHAT_SPACES`, off for the wing) and an empty-vs-null cache-safety fix.
  Two corrections vs the Pacific copy: `buildWorkspaceCapidMaps` keeps
  `customer:"my_customer"` (not `domain:`); `getMembersForChatSpaces_` falls back to
  `INDEFINITE` not `LIFE`.
- **`config.gs` (v1.2.0)** — region feature flags + `REGION_CAPWATCH_DATA_FOLDER_ID`.

### Notes

- **Adding the `contacts` scope requires re-authorization** on all three projects at the
  next `clasp push` / first run. Verify during the Pacific dry-run that this scope suffices
  for the M8 Domain Shared Contacts feed.
- `PCRCAP.ORG.js` (one-off `@pcrcap.org` audit) intentionally **not** folded.

## [2026-07-09] — Pacific tenant profile + profile-driven per-tenant orgs

Code-side reconciliation so the Pacific Region project can run the shared `src/`,
differentiated only by configuration (see [docs/PACIFIC_DIFF.md](docs/PACIFIC_DIFF.md)).
Behavior-preserving for the seniors and cadets tenants. **Not yet deployed to
Pacific** — deployment is on hold pending 2SV for `automation@pcr.cap.gov`.

### Added

- **`pacific` profile** in `TENANT_PROFILES_` (`config.gs`) — single-unit region
  HQ (PCR-PCR-001): senior member types (`SENIOR`/`FIFTY YEAR`/`INDEFINITE`/`CADET`;
  region confirmed no AEM and all `INDEFINITE`), holding unit 1345, org-path sync
  and squadron-group auto-create disabled. Selected with `TENANT_PROFILE=pacific`.

### Changed

- **`config.gs` (v1.1.0)** — `EXCLUDED_ORG_IDS` and `SPECIAL_ORGS.AEM_UNIT` are now
  profile-driven (`PROFILE_.*`) instead of hard-coded CA-wing values, so holding
  units and AEM handling vary per tenant. Seniors/cadets values unchanged
  (`['1297','368']`, AEM_UNIT `''`).
- **`GetCapwatch.gs` (v1.0.0)** — `getCapwatch()` now calls `syncOrgPaths()` only
  when `PROFILE_.SYNC_ORG_PATHS` is true, so single-unit region tenants skip
  org-path auto-mapping (and its IT summary email) entirely.
- **`config-tenants/pacific.json`** — populated with the live non-secret identity
  read via clasp (was an all-blank stub); scriptId note corrected. (PR #9)

## [2026-07-09] — Reconcile live tenants + per-tenant config hardening

Merged via PR #7 (`reconcile-live-hardening`). Reconciles the repository with the
code actually deployed across the **seniors** (`cawgcap.org`) and **cadets**
(`cawgcadets.org`) Workspace tenants, and removes the shared-config clobber
footgun (a `clasp push` overwriting a tenant's `config.gs`).

### Added

- `docs/ADMIN_GUIDE.md` — successor / "hit by a bus" runbook: three-tenant
  deployment, access checklist, Apps Script + clasp workflow, secrets and
  Script-Properties inventory, automation schedule, entry-point reference,
  disaster recovery.
- `config-tenants/{seniors,cadets,pacific}.json` — canonical **non-secret**
  per-tenant identity templates (kept out of every `clasp push` by `.claspignore`).
- `getTenantConfig_()`, `setupTenantConfig()`, `validateTenantConfig()` in
  `config.gs`; `TENANT_PROFILE` + `TENANT_PROFILES_` (`seniors` | `cadets`)
  selecting per-tenant behavior (member types, Cadet-Lite, squadron-group set).

### Changed

- **`config.gs` (v1.0.0)** — Per-tenant identity (domain, ORGID, folder/sheet IDs,
  contact emails) now read from `TENANT_*` **Script Properties**, not literals, so
  a `clasp push` no longer clobbers a tenant's config. No cross-tenant fallback:
  an unconfigured project yields empty identity and fails loudly rather than acting
  on the wrong domain.
- **`SyncOrgPaths.gs` (v1.0.0)** — OrgPath sync recipient resolved per-tenant via
  `getOrgPathSyncEmail_()` (was hardcoded `it@cawgcap.org`).
- **`UpdateMembers.gs` (v1.4.5)** — `testImpersonationToken` uses `console.log`
  (the codebase overrides the global `Logger`, which has no `.log`).

### Fixed

- **`AdminDirectory.Users.list` domain → customer** — standardized all call sites
  to `customer: "my_customer"` (the `{domain: ...}` form returned 400 Bad Request
  on the cadets tenant; identical result on a single-domain customer). Affects
  `ManageLicenses.gs` (v1.0.0), `UpdateMembers.gs` (v1.4.5), `UpdateGroups.gs`
  (v1.3.8), `UpdateChatSpaces.gs` (v1.0.0), `UpdateCalendars.gs` (v1.2.4), and
  `SquadronGroups.gs` (v1.2.7).
- Correct CAPWATCH senior member type is **`INDEFINITE`, not `LIFE`**
  (`config.gs`, `SendRetentionEmail.gs` v1.0.0, and squadron distribution lists).

### Operational

- **Cadets tenant re-enabled end-to-end**: config resolved, license/group previews
  clean (0 destructive actions), watched `updateAllMembers()` applied 22 benign
  org/duty/grade changes + 3 new accounts with 0 errors; time-driven triggers
  recreated under `automation@cawgcadets.org`.
- Both **seniors** and **cadets** projects deployed to this code
  (`npm run push:seniors` / `push:cadets`).

### Known / carried forward

- **Pacific** (`pcr.cap.gov`) tenant not yet reconciled or verified.
- Leaked service-account private key in git history still needs **GCP key rotation**.
- `CONFIG.CUSTOMER_ID` is referenced by a few calls but undefined (latent cleanup).

---

## Earlier PCR-fork changes (pre-2026-07-09)

Reconstructed from git history for continuity; predates this changelog file.

- **Security hardening** (PR #6, `c6b3099`) — randomized temporary Workspace
  passwords (~244-bit entropy, not derivable from public data) and mission-webhook
  hardening.
- **Member eligibility** (PR #5) — corrected `LIFE` → `INDEFINITE` member type and
  switched holding-unit exclusion to ORGID-based (`63301be`); fixed broken email
  templates and mislabeled post-creation errors (`9688c88`).
- **Level I gating** (`c425948`) — senior account provisioning gated on Level I
  completion.
