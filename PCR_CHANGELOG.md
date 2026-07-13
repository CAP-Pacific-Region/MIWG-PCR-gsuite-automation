# PCR Changelog

Pacific Region (PCR) fork-specific changes to the CAPWATCH / Google Workspace
automation, layered on top of the upstream `cap-miwg/gsuite-automation` project.
Upstream changes are tracked in [CHANGELOG.md](CHANGELOG.md); **this file records
only what the Pacific Region deployment adds or diverges on.**

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).
Individual source files carry their own SemVer version in their header
(see [docs/VERSIONING.md](docs/VERSIONING.md)); the per-file version is noted
next to each entry below.

## [Unreleased]

### Added

- **`src/cross-tenant-contacts/CrossTenantContacts.gs`** (v0.1.0, draft) — folds the
  wing's two separate cross-tenant directory-sync projects (cadets `1fJRqo…`, seniors
  `1b2JSIB…`) into the shared `src/` as one **role-relative, Script-Property-configured**
  module. Publishes the **peer** Workspace tenant's members into this tenant's Global
  Address List (seniors ⇄ cadets) as Domain Shared Contacts.
  - **Spreadsheet-free.** Replaces the old export→sheet→import pipeline. The peer roster
    (incl. cadet-lite members with no account) comes from one wing CAPWATCH pull via the
    existing `getMembers()`; the authoritative Workspace email comes from a live read of
    the **peer** directory (read-only peer-tenant service account, DWD, same JWT pattern
    as `getImpersonatedToken_`).
  - **Email waterfall** per member: peer Workspace `primaryEmail` (authoritative; fixes
    collisions/renames) → CAPWATCH `MbrContact` personal email (cadet-lite / no-account)
    → `do.not.contact+<CAPID>@` sentinel (presence-only, opt-in).
  - **Stateless reconcile.** No sync-state sheet: managed contacts are marked by
    `orgName` and carry their content hash in a `gContact:userDefinedField`.
  - **Parent-group sync** (`syncCrossTenantParentContacts`, gated by `RUN_PARENTS`, on for
    seniors) publishes the peer tenant's `*.parents@` distribution groups into the GAL
    under a separate `<WING>_PARENTS` marker.
  - Entry points `syncCrossTenantContacts` / `syncCrossTenantParentContacts`; helpers
    `setupCrossTenantConfig()` / `validateCrossTenantConfig()`. All symbols `xt`-prefixed
    (zero collisions with existing `src/`).
- **`config.gs`** — `PROFILE_.CROSS_TENANT` block per profile (on for seniors/cadets, off
  for pacific): `RUN_INBOUND`, `RUN_PARENTS`, `PEER_TYPES`, `PEER_LABEL`, `EMIT_PLACEHOLDERS`.
- **`XT_PEER_*` Script Properties** — `XT_PEER_DOMAIN` (canonical values added to
  `config-tenants/seniors.json` + `cadets.json`) and the read-only peer SA creds
  `XT_PEER_SA_EMAIL` / `XT_PEER_SA_SUBJECT` / `XT_PEER_SA_KEY` (secret; on-project only).
- **`https://www.google.com/m8/feeds`** OAuth scope in `appsscript.json` (Domain Shared
  Contacts; the manifest previously had only `.../auth/contacts`). Requires one re-consent
  per project.
- **[docs/CROSS_TENANT_CONTACTS.md](docs/CROSS_TENANT_CONTACTS.md)** — architecture, the
  email waterfall, per-project setup, and migration off the two legacy projects.
- **[docs/GCP_PROJECT_MIGRATION.md](docs/GCP_PROJECT_MIGRATION.md)** — one-way migration of a
  tenant's Apps Script project from its default GCP project to a standard project, required
  to enable the Contacts API (m8 feed) for any shared-contacts feature. Surfaced by the
  seniors canary: default projects deny `serviceusage.services.enable`.

### Notes

- Draft — not yet deployed. Requires per-peer read-only service accounts (DWD:
  `admin.directory.user.readonly` + `admin.directory.group.readonly`), the `m8/feeds`
  re-consent, and triggers. **Migration:** the legacy projects tag managed contacts
  `orgName=CAWG` / `CAWG_CADET_PARENTS_GROUPS`; this module uses `CONFIG.WING` (`CA`) /
  `CA_PARENTS`, so decommission the old projects and clean up their contacts (or re-tag)
  to avoid duplicates. See the doc.

### Fixed

- **`UnitVisitReport.gs` (v1.0.1)** — `buildRegionUnitVisitReport()` failed with
  "You can't merge frozen and non-frozen columns" in `buildWingTab_`.
  `clear()`/`clearFormats()` don't reset a tab's freeze state, so a pre-existing
  frozen column made the `A1:I1` title merge fail. Now resets frozen rows/columns
  before merging. (Surfaced during the Pacific go-live.)

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
