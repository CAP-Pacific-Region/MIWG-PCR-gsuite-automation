# Pacific Region Reconciliation & Divergence Report

**Status:** Analysis complete. Deployment to Pacific is **ON HOLD** pending 2SV for
`automation@pcr.cap.gov`. This document is the plan of record for lining the Pacific
Region project up with the shared repo.

**Goal (from the wing/region director):**
> Make the source of all three tenants line up in GitHub, differentiated only by
> configuration.

Seniors (`cawgcap.org`) and cadets (`cawgcadets.org`) already run the identical `src/`,
differentiated by per-project **Script Properties** (`TENANT_*`) and a `TENANT_PROFILE`.
Pacific is the last tenant still on its own diverged copy. This report captures what
diverges, what to keep, what to drop, and how to fold Pacific onto the shared codebase.

---

## 0. Access correction (resolved)

Earlier notes claimed the Pacific project was owned by `it@pcr.cap.gov` and that `clasp`
access was blocked. **That was wrong.** The real cause was a **stale `scriptId`** in
`clasp-targets/pacific.clasp.json` pointing at a dead project
(`1Y_7HIPo83OIOaC5opnPb7hmjvy7SDL2cMkuBhJNOYmM4BgZLXXSPitfi`).

The live project is **"PCR Automation"**, scriptId
`1s2Fmdo0sxWjuPawYBU_dCGYa5qA0h8LuGbQkIGzptzhBlTqL14JqW-T0`, fully readable with the
`automation@pcr.cap.gov` clasp login. `clasp-targets/pacific.clasp.json` has been
corrected in this change. (Mission provisioning lives in a separate project, "PCR Mission
Automaton" — out of scope here.)

---

## 1. Pacific identity (verified from live config)

| Field | Value |
|---|---|
| Domain | `pcr.cap.gov` |
| Email domain | `@pcr.cap.gov` |
| CAPWATCH ORGID | `434` |
| Region / Wing | `PCR` / `PCR` |
| Managed unit | **PCR-PCR-001 only** (single region HQ unit) |
| SA project | `pcr-capwatch` |

These non-secret values are now recorded in `config-tenants/pacific.json` (previously an
all-blank stub) so they can be loaded into Script Properties during migration.

Pacific-specific behavioral values found in its live `config.gs`, and how the shared model
now covers them (see `pacific` profile in `config.gs`, §5):

- `MEMBER_TYPES.ACTIVE = ['CADET','SENIOR','FIFTY YEAR','LIFE','AEM']` — live used legacy
  `LIFE` and `AEM`. **Region confirmed (2026-07-09): no AEM automation, and all senior
  members are typed `INDEFINITE`.** The `pacific` profile therefore uses
  `['SENIOR','FIFTY YEAR','INDEFINITE','CADET']` — no `LIFE`, no `AEM`.
- `EXCLUDED_ORG_IDS = ['1345','']` — different holding unit than seniors/cadets; now the
  `pacific` profile's `EXCLUDED_ORG_IDS: ['1345']`.
- `SPECIAL_ORGS.AEM_UNIT = '182'` — dropped (no AEM); `pacific` profile `AEM_UNIT: ''`.
- `REGION_CAPWATCH_DATA_FOLDER_ID = '1lU9yWHPf1Eij3AEQPmMR8ki7EpgslV9z'` — a region-level
  CAPWATCH folder with no equivalent in the shared config. Confirm whether it's still used;
  not currently modeled.

**Resolved:** Pacific runs on the shared code via a dedicated **`pacific` profile**
(`TENANT_PROFILE=pacific`) — single-unit, no AEM, INDEFINITE, org-path sync + squadron
groups off. See §5.

---

## 2. Shared-module divergence (Pacific live vs repo `src/`)

Line counts are from a normalized-LF diff of Pacific's pulled copy against the repo. Small
counts are cosmetic/whitespace; large counts are real drift.

| Module | Δ lines | Pacific version | Notes |
|---|---:|---|---|
| `utils.gs` | 0 | — | Identical. |
| `UpdateCalendars.gs` | 2 | 1.2.3 | Near-identical; repo is 1.2.4. |
| `GetCapwatch.gs` | 4 | — | Near-identical. Pacific filename typo: `GetCaptwatch.js`. |
| `SendRetentionEmail.gs` | 39 | — | Minor drift. |
| `UpdateMembers.gs` | 180 | 1.4.3 | Repo is 1.4.5 (has `customer:"my_customer"` fix). |
| `UpdateGroups.gs` | 491 | 1.3 | Repo is 1.3.8. Significant drift. |
| `ManageLicenses.gs` | 573 | — (no version) | Pacific 903 lines vs repo 1444. Major drift. |
| `config.gs` | 576 | — | Pacific 335 lines vs repo 653. Pre-refactor (no `getTenantConfig_`). |
| `UpdateChatSpaces.gs` | 1007 | 1.2 | Pacific **larger** (1767 vs 811) — different implementation. |

**Bug still live in Pacific:** `AdminDirectory.Users.list({domain: ...})`, which 400s on
some tenants. Present in Pacific's ManageLicenses (×6), UpdateMembers (×4), UpdateGroups
(×2), UpdateCalendars (×1), UpdateChatSpaces (×1). The repo already fixed all of these to
`customer:"my_customer"`. Adopting the shared `src/` resolves them wholesale.

**Direction:** for every module above, the **repo version wins**. Pacific's drift is older
code plus a few region tweaks that migrate into config/profile, not into forked logic.
`UpdateChatSpaces.gs` is the one to eyeball before overwriting — Pacific's is a larger,
different implementation; confirm no region-only chat behavior is being dropped that the
region still wants (see §3, chat decision).

---

## 3. Pacific-only files — disposition

**Architecture (director direction, 2026-07-09): the shared `src/` is identical for all
three tenants.** A module a tenant doesn't use is **disabled by configuration**, not removed
— the code is present everywhere, gated by a profile flag (and by which triggers a tenant
schedules). This replaces the earlier "relocate region files to a separate project" idea:
folding them in means the reconciling push no longer deletes any region functionality.

So the region-used modules are **folded into `src/` and gated off for the wing**. Verified
2026-07-09 by diffing the live project against the repo.

| File | Decision | Detail |
|---|---|---|
| `updateRegionGroupChats.js` | **Fold into `src/`; gate OFF for wing** | `updateRegionGroupChats()` — region duty groups + duty chat spaces. Present in shared code; enabled only for `pacific`. |
| `UnitVisitReport.js` | **Fold into `src/`; gate OFF for wing** | `buildRegionUnitVisitReport()` — region-wide unit-visit spreadsheet. Present in shared code; enabled only for `pacific`. |
| `SharedContacts.js` | **Fold into `src/`; enable for Pacific** | `runExternalContactsToDomainSharedContacts()` — syncs the "External Contacts" sheet tab into Domain Shared Contacts. Already CONFIG-driven (`AUTOMATION_SPREADSHEET_ID`/`DOMAIN`/`WING`) + uses shared `sanitizeEmail()`. Region uses it; wing handles shared contacts in a separate project, so wing keeps it disabled. |
| `PCRCAP.ORG.js` | **Drop (not automation)** | One-off, read-only audit that lists every `@pcrcap.org` user/group/alias/resource. Hardcoded domain; declares top-level `DOMAIN_TO_FIND` / `CUSTOMER_ID` globals (collision risk). Not part of the shared codebase — re-paste into the editor if the audit is ever needed again. |

**Gating mechanism:** add per-feature profile flags (e.g. `RUN_REGION_GROUP_CHATS`,
`RUN_UNIT_VISIT_REPORT`, `RUN_SHARED_CONTACTS`) to `TENANT_PROFILES_`. Each entry point
guards on its flag (`if (!PROFILE_.RUN_… ) return;`) so it no-ops on tenants where it's off,
even if run manually. Wing profiles (`seniors`/`cadets`) set them `false`; `pacific` `true`.
Triggers are only scheduled where the feature is on.

> ℹ️ **`UpdateChatSpaces` convergence — analyzed 2026-07-09 (function-level diff).** Pacific's
> file is a strict **superset**: +9 functions (automation-group chat spaces, user-additions
> chat spaces, their loaders, a dry-run switch); nothing exists only in the wing. Of the
> shared functions: 6 identical; the rest are either inert (`getMembersForChatSpaces_` dead
> fallback), **PCR-gated** (`WING==='PCR'` committee legacy-rename + "skip unit spaces" — inert
> for the wing), or **benign improvements** (empty-vs-null cache-safety **bugfix**, committee
> description/guidelines, `createChatSpace` history/external-user support).
>
> **Plan — adopt Pacific's version as the single shared file, with three corrections:**
> 1. **Flag-gate the two new features OFF for the wing.** They do **not** self-gate: their
>    loaders fall back to the `Groups` / `User Additions` tabs the wing already has, so they
>    would create unwanted spaces. Guard both calls behind a profile flag (wing off, pacific on).
> 2. **Keep `customer:"my_customer"`** in `buildWorkspaceCapidMaps` — Pacific reintroduced the
>    `domain: CONFIG.DOMAIN` 400 bug we fixed repo-wide; take its `emailToUserId` addition only.
> 3. **Keep the wing's `INDEFINITE` fallback** in `getMembersForChatSpaces_` (Pacific has stale
>    `LIFE`; it's a dead fallback either way).

---

## 4. Repo modules Pacific does not run — leaner variant

Pacific is a **single-unit** region tenant (PCR-PCR-001 only). Several repo modules assume
a wing with many subordinate squadrons and are **not needed** for Pacific:

| Module | Needed on Pacific? | Reason |
|---|---|---|
| `SyncOrgPaths.gs` | **No** | Auto-maps new/deactivated *squadrons*. Pacific has one fixed unit — nothing to sync. |
| `squadron-groups/SquadronGroups.gs` | **No** | Per-squadron groups; Pacific has no subordinate squadrons. |
| `UpdateCAWGCadetGroups.gs` | **No** | CAWG cadet-tenant crossover; N/A to a region. |
| `UpdateResources.gs` | **No** (confirmed) | Region does not manage aircraft/vehicle resources; the file was never even present in the live Pacific project. |
| `groupAdministration.gs` | Optional | Ad-hoc admin utility; harmless if shipped, unused otherwise. |
| `MissionProvisioning.gs` | **No** | Region mission handled by the separate "PCR Mission Automaton". |

Because these are **trigger-invoked**, the clean way to "not run" them on Pacific is simply
**not to create their triggers** — the shared `src/` can still contain them (that's the
point of one codebase). Two integration points need explicit gating so an unwanted module
doesn't run as a side effect:

1. **`getCapwatch()` calls `syncOrgPaths()` at the end.** On a single-unit region this
   should be a no-op, but `syncOrgPaths` filters on `region === 'PCR' && wing === CONFIG.WING`
   and emails IT. Gate it behind a profile flag (e.g. `SYNC_ORG_PATHS: false` for the
   `pacific` profile) rather than relying on it being harmless.
2. Any squadron-group provisioning called from the member/group sync path must be
   likewise gated by the profile.

---

## 5. Config differentiation plan

Fold Pacific onto shared `src/` by moving everything region-specific into Script Properties
and a profile — never into forked code.

**Script Properties (identity)** — load from `config-tenants/pacific.json` via
`setupTenantConfig()`, then `validateTenantConfig()`. Secrets (`SA_PRIVATE_KEY`,
`CAPWATCH_AUTHORIZATION`) stay in Script Properties only, never committed.

**`pacific` profile in `TENANT_PROFILES_` (src/config.gs)** — added (PR #10) to carry the
region behavior the `seniors` profile doesn't:
- `MEMBER_TYPES_ACTIVE = ['SENIOR','FIFTY YEAR','INDEFINITE','CADET']` — no `LIFE`, no
  `AEM` (region confirmed).
- `SYNC_ORG_PATHS: false` and squadron-group auto-create `false` (single-unit).
- `EXCLUDED_ORG_IDS: ['1345']` and `AEM_UNIT: ''` — implemented as **profile-driven**
  (`PROFILE_.*`) rather than new Script Properties, since Pacific is its own profile and
  the values stay version-controlled. `EXCLUDED_ORG_IDS`/`AEM_UNIT` are now profile-driven
  for all three tenants (seniors/cadets keep `['1297','368']` / `''`).
- `REGION_CAPWATCH_DATA_FOLDER_ID`: not modeled; confirm whether it's still used.

**Triggers on Pacific (lean set):** getCapwatch, updateAllMembers, updateEmailGroups,
license lifecycle, retention email. **Not** squadron groups, **not** org-path sync,
**not** cadet groups.

---

## 6. Migration checklist (execute once 2SV clears — do NOT run while on hold)

**Code-side (can land before 2SV, in PRs):**

1. ~~**Add the `pacific` profile.**~~ **Done — PR #10** (`reconcile/pacific-profile`):
   `pacific` profile + profile-driven `EXCLUDED_ORG_IDS`/`AEM_UNIT` + `SYNC_ORG_PATHS` gate.
2. **Fold the region modules into `src/`** with per-feature profile flags (§3):
   `updateRegionGroupChats`, `UnitVisitReport`, `SharedContacts` — each guarded by
   `RUN_*` flags (wing `false`, pacific `true`). Drop `PCRCAP.ORG`. Requires modeling
   `REGION_CAPWATCH_DATA_FOLDER_ID` if the folded region code needs it.
3. **Converge `UpdateChatSpaces`** to a single shared version (adopt Pacific's superset, or
   confirm its extras are inert for the wing) — line-by-line diff first (§3/§7).

**Deploy-side (execute once 2SV clears — do NOT run while on hold):**

4. **Back up** the live "PCR Automation" project (versioned clasp pull, archived).
5. **Set Script Properties** on "PCR Automation" from `config-tenants/pacific.json`
   (`setupTenantConfig()` → `validateTenantConfig()`), plus `TENANT_PROFILE=pacific` and
   the SA secrets.
6. **Dry run:** with the shared code loaded, preview member/group/license/chat actions and
   diff against current behavior before any write.
7. **Push shared `src/`** to Pacific (`npm run push:pacific`). With everything folded in,
   the only files removed are the deliberately-dropped `PCRCAP.ORG` — no region
   functionality is lost.
8. **Recreate the trigger set** under `automation@pcr.cap.gov` — the lean member/group/
   license/retention set **plus** the region features now enabled (`updateRegionGroupChats`,
   `buildRegionUnitVisitReport`, shared-contacts sync).
9. **Verify** a full watched run (0 unexpected changes) and confirm all three tenants run
   byte-identical `src/`.

---

## 7. Open items to confirm with the region

- ~~Does Pacific still run **AEM** automation?~~ **Resolved (2026-07-09): no.** AEM dropped
  from the `pacific` profile.
- ~~Any live members still typed **`LIFE`**?~~ **Resolved: no — all `INDEFINITE`.**
- ~~Is `UpdateResources` used by the region?~~ **Resolved: no** — Pacific won't schedule it.
- ~~Keep `UnitVisitReport` / `updateRegionGroupChats`?~~ **Resolved: region uses both →
  fold into shared `src/`, gated OFF for the wing** (identical-code model, §3).
- ~~Is `SharedContacts` region-used?~~ **Resolved: yes** (verified — CONFIG-driven External
  Contacts → Domain Shared Contacts sync). Fold in; enable for Pacific, off for the wing.
- ~~`PCRCAP.ORG.js`?~~ **Resolved: drop** — one-off read-only `@pcrcap.org` audit, not
  automation.
- ~~**`UpdateChatSpaces` superset**~~ **Resolved (diff done 2026-07-09): converge to Pacific's
  version** as the single shared file, flag-gating the two new features OFF for the wing (their
  loaders fall back to the wing's `Groups`/`User Additions` tabs, so they don't self-gate) and
  keeping `customer:"my_customer"` + the wing's `INDEFINITE` fallback (§3 note).
- **`REGION_CAPWATCH_DATA_FOLDER_ID` (still open)** — used by the folded region code? Must be
  modeled (Script Property or profile) if so.

---

*Related: `config-tenants/pacific.json` (identity), `clasp-targets/pacific.clasp.json`
(corrected scriptId), `docs/ADMIN_GUIDE.md` (secrets & tenant setup), `PCR_CHANGELOG.md`.*
