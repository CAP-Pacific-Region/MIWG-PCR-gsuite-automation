# CAPWATCH → Google Workspace Automation (Pacific Region)

> **Keeps multiple Google Workspace tenants in sync with CAPWATCH membership data for Civil Air Patrol — one shared codebase, several domains.**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?logo=google&logoColor=white)](https://www.google.com/script/start/)
[![Status: Production](https://img.shields.io/badge/Status-Production-green)](https://github.com/CAP-Pacific-Region/MIWG-PCR-gsuite-automation)

> **📖 New administrator taking this over?** Start with the
> **[Administrator & Successor Guide](docs/ADMIN_GUIDE.md)** — the PCR "hit by a bus" runbook
> covering the three-tenant deployment, access checklist, secrets, schedule, and disaster recovery.
> This README is the map; the Admin Guide is the ground truth for operating the live system.

## Overview

This project synchronizes Civil Air Patrol (CAP) Google Workspace environments with **CAPWATCH**
membership data. It began as a fork of the Michigan Wing single-wing automation, but has grown into
a **multi-tenant platform**: a single `src/` codebase is deployed to **three independent Workspace
tenants** and adapts its behavior per tenant through Google Apps Script **Script Properties** — no
per-tenant code branches, no forks.

### The three tenants

| Tenant | Domain | Profile | Role |
|--------|--------|---------|------|
| **Seniors** | `cawgcap.org` | `seniors` (default) | CAWG senior members; also the **cross-tenant driver** |
| **Cadets** | `cawgcadets.org` | `cadets` | CAWG cadets (cadet-lite accounts, smaller group set) |
| **Pacific Region** | `pcr.cap.gov` | `region` | Region-level features (mission webhook, unit-visit report, region chats) |

The **same code** runs on all three. Each project carries its own `TENANT_*` identity and a
`TENANT_PROFILE` selector in Script Properties (which `clasp push` never touches), so a deploy can
never repoint one tenant at another's domain. Canonical non-secret values are version-controlled in
[`config-tenants/`](config-tenants/README.md); secrets live only in each project's Script Properties.

> ⚠️ **`src/config.gs` is tenant-neutral by design and is overwritten on every push.** Never
> hand-edit a domain, ORGID, or folder ID into it — set a Script Property instead. See
> [config-tenants/](config-tenants/README.md) and [Admin Guide §5](docs/ADMIN_GUIDE.md#5-the-three-tenants-and-how-code-gets-deployed).

### What it does

- **Account management** — creates, updates, suspends, reactivates, and deletes Workspace accounts from CAPWATCH member status, placing each in the correct Organizational Unit.
- **Email group synchronization** — wing / duty-position / specialty distribution lists, plus per-squadron all-hands / cadets / seniors / parents lists (the exact set is profile-driven).
- **Cross-tenant integration** — the seniors tenant nests cadet groups (living on `cawgcadets.org`) into matching senior groups, and seniors ⇄ cadets publish each other's members into their Global Address List (see [cross-tenant contacts](docs/CROSS_TENANT_CONTACTS.md)).
- **Calendars & Chat** — shares unit calendars with new/transferred members; syncs squadron and committee Chat spaces.
- **Calendar resources** — aircraft and vehicles as bookable Calendar resources, with squadrons as buildings (seniors only).
- **License lifecycle** — suspends expired members, and deletes long-ineligible accounts to stay under the Workspace-for-Nonprofits seat cap.
- **Cadet → senior transition** — when a cadet ages out or converts, the cadets tenant carries their mail, Drive, and contacts to their new seniors-tenant account before the old account is deleted, then forwards the freed address (cadets tenant only; the final delete stays a manual step).
- **Region features** (Pacific only) — region duty groups + Chat spaces, a region-wide unit-visit report, and an on-demand **mission provisioning webhook** (Group + Chat space + Drive folder per mission).

### Key characteristics

- ✅ **One codebase, many tenants** — behavior varies by Script Property, not by code fork.
- ✅ **Zero manual account updates** — member changes flow from CAPWATCH automatically.
- ✅ **No infrastructure** — runs entirely inside Google Apps Script; time-driven triggers do the work overnight.
- ✅ **Seat-cap aware** — the Nonprofits edition caps users regardless of suspension, so lifecycle logic deletes to reclaim seats.
- ✅ **Auditable** — structured logging, per-tenant execution logs, and error/report emails.

## Table of Contents

- [Architecture](#architecture)
- [Repository layout](#repository-layout)
- [Prerequisites](#prerequisites)
- [Getting started (developers)](#getting-started-developers)
- [Per-tenant configuration](#per-tenant-configuration)
- [What runs, when](#what-runs-when)
- [Modules](#modules)
- [Documentation](#documentation)
- [Contributing](#contributing)
- [License & acknowledgments](#license--acknowledgments)

## Architecture

```
┌──────────────────────────────────────────────────────────────┐
│                      CAPWATCH eServices                       │
│                      (source of truth)                        │
└───────────────────────────┬──────────────────────────────────┘
                            │ daily download (after ~4 AM blackout)
                            ▼
┌──────────────────────────────────────────────────────────────┐
│                   One shared codebase (src/)                  │
│      Apps Script: config-neutral; behavior via Script Props   │
└───────────┬───────────────────┬───────────────────┬──────────┘
   clasp push│         clasp push│          clasp push│
            ▼                   ▼                   ▼
   ┌──────────────┐    ┌──────────────┐    ┌──────────────┐
   │  SENIORS     │    │   CADETS     │    │  PACIFIC     │
   │ cawgcap.org  │◄──►│cawgcadets.org│    │  pcr.cap.gov │
   │ profile:     │xt  │ profile:     │    │ profile:     │
   │  seniors     │    │  cadets      │    │  region      │
   └──────┬───────┘    └──────┬───────┘    └──────┬───────┘
         │ Admin SDK / Calendar / Chat / Groups Settings APIs
         ▼                   ▼                   ▼
   ┌──────────────────────────────────────────────────────┐
   │  Workspace: accounts · OUs · groups · calendars ·     │
   │  chat spaces · resources · shared contacts            │
   └──────────────────────────────────────────────────────┘

  Per-tenant identity + behavior comes from Script Properties
  (TENANT_* + TENANT_PROFILE), never from committed code.
  "xt" = cross-tenant sync (group nesting + shared contacts).
```

Each tenant's project has its own time-driven triggers and its own dedicated Google Cloud **service
account** (used only for the Gmail-settings and Calendar calls that require domain-wide
impersonation). Because the three projects are pushed independently, **they can silently run
different code** — `master` is not proof of what is live on any tenant. See
[Admin Guide §5](docs/ADMIN_GUIDE.md#5-the-three-tenants-and-how-code-gets-deployed).

## Repository layout

```
src/                          # The shared codebase — deployed unchanged to all three tenants
├── config.gs                 # Tenant-NEUTRAL config; reads identity/profile from Script Properties
├── utils.gs                  # parseFile (CSV cache), executeWithRetry, email/validation helpers
├── GetCapwatch.gs            # Download CAPWATCH ZIP → Drive; credential storage
├── SyncOrgPaths.gs           # ORGID→OU map; auto-creates OUs for new squadrons
├── appsscript.json           # Manifest: advanced services, OAuth scopes, web-app config
├── accounts-and-groups/      # Members, groups, licenses, calendars, chat, cross-tenant cadet groups + cadet→senior transition
├── squadron-groups/          # Per-squadron distribution lists (profile-driven set)
├── calendar-resources/       # Aircraft/vehicle resources + squadron buildings (seniors only)
├── cross-tenant-contacts/    # Peer-tenant directory → this tenant's shared contacts (seniors ⇄ cadets)
├── region/                   # Region group chats + unit-visit report (Pacific only)
├── mission-provisioning/     # doPost webhook: Group + Chat space + Drive folder per mission
└── recruiting-and-retention/ # Age-out / expiration / welcome emails + HTML templates

config-tenants/               # Canonical NON-SECRET per-tenant values (repo-only; never pushed)
├── seniors.json  cadets.json  region.json
clasp-targets/                # {scriptId, rootDir:../src} pointers, one per tenant
docs/                         # Admin Guide, cross-tenant, migration, troubleshooting, etc.
```

## Prerequisites

- **Super Admin** on each Workspace tenant you intend to operate.
- **CAP eServices / CAPWATCH** credentials for the relevant ORGID (requires a security assessment and commander approval).
- **[clasp](https://github.com/google/clasp)** and Node, plus the Apps Script API enabled for your Google account ([script.google.com/home/usersettings](https://script.google.com/home/usersettings)).
- A dedicated Google Cloud **service account** per tenant for domain-wide-delegated Gmail/Calendar calls (see [Admin Guide §3](docs/ADMIN_GUIDE.md#3-system-inventory-the-facts-you-cannot-guess)).

## Getting started (developers)

There is **one** `src/` and three clasp targets. All operations go through npm scripts:

```bash
npm install -g @google/clasp
clasp login                       # log in as an owner of the target project(s)

npm run status:seniors            # preview what would change on a tenant
npm run push:seniors              # deploy src/ → seniors project
npm run pull:seniors              # pull project → src/ (to inspect drift)
npm run open:seniors              # open the project in the browser
# ...and the :cadets / :region equivalents
```

**Recommended change flow:** branch → edit `src/` → `push:seniors` → run the relevant `preview…`
function and check **Executions** → push `cadets`, then `region` → update
[PCR_CHANGELOG.md](PCR_CHANGELOG.md) → open a PR to **`CAP-Pacific-Region/MIWG-PCR-gsuite-automation`
master** (not the upstream `cap-miwg` repo). Full details in
[Admin Guide §6](docs/ADMIN_GUIDE.md#6-local-development-with-clasp) and
[DEVELOPMENT.md](docs/DEVELOPMENT.md).

> Push to one tenant, confirm it's healthy, then the next — never push all three blind, and remember
> each push resets that project's `config.gs` to the shared copy.

## Per-tenant configuration

A project's identity and behavior come entirely from **Script Properties**, set once per project
and never overwritten by a push:

- **Identity** — `TENANT_DOMAIN`, `TENANT_EMAIL_DOMAIN`, `TENANT_CAPWATCH_ORGID`, `TENANT_WING`, `TENANT_REGION`, the Drive folder / spreadsheet IDs, and contact addresses. Canonical values live in [`config-tenants/<tenant>.json`](config-tenants/README.md).
- **Behavior** — `TENANT_PROFILE` (`seniors` | `cadets` | `region`) selects member types, cadet-lite mode, the squadron-group set, region-feature flags, and cross-tenant behavior (`PROFILE_` in `config.gs`).
- **Secrets** — `SA_IMPERSONATION_EMAIL` / `SA_PRIVATE_KEY` (per-tenant service account), `XT_PEER_*` (cross-tenant peer SA), `MISSION_WEBHOOK_SECRET`, and the per-user `CAPWATCH_AUTHORIZATION` token. Never committed.

Apply values with `setupTenantConfig()` (or by hand), then run `validateTenantConfig()`. There is
**no fallback** — an unconfigured project fails loudly rather than acting on the wrong domain, so
**set a project's `TENANT_*` properties before pushing to it.** Full inventory:
[Admin Guide §7](docs/ADMIN_GUIDE.md#7-secrets-and-script-properties-the-part-that-breaks-silently).

## What runs, when

Each tenant schedules its own time-driven triggers. Order matters — CAPWATCH has an overnight
blackout, so downloads run after ~4 AM and everything else consumes the data afterward:

| Order | Function | Cadence | Notes |
|------:|----------|---------|-------|
| 1 | `getCapwatch()` | Daily 4–5 AM | Download CAPWATCH ZIP → Drive; refreshes OrgPaths |
| 2 | `updateAllMembers()` | Daily 5–6 AM | Create/update accounts, OU placement, aliases |
| 3 | `suspendExpiredMembers()` | Daily 5–6 AM | Suspend past the 7-day grace window |
| 4 | `updateEmailGroups()` | Daily 5–6 AM | Wing / duty / specialty distribution groups |
| 5 | `updateAllSquadronGroups()` | Daily 6–7 AM | Per-squadron lists; batched via `SQUADRON_BATCH_INDEX` |
| 6 | `updateAdditionalGroupMembers()` | Daily 6–7 AM | Manual additions from the `User Additions` sheet |
| 7 | `syncMemberCalendarsDaily()` | Daily | Share unit calendars with new/transferred members |
| 8 | `updateChatSpaces()` | Daily/weekly | Squadron & committee Chat spaces |
| 9 | `updateResources()` | Weekly (Sun) | Aircraft/vehicle resources + buildings — **seniors only** |
| 10 | `manageLicenseLifecycle()` | Monthly | Reactivate renewed; delete long-ineligible to free seats |

Cross-tenant (`updateCAWGCadetGroups()`, `syncCrossTenantContacts`) and region-only functions
(`updateRegionGroupChats()`, `buildRegionUnitVisitReport()`, `runExternalContactsToDomainSharedContacts()`)
run where their profile flag enables them. Confirm the **actual** triggers per project — the table
is the intended design. See [Admin Guide §8–9](docs/ADMIN_GUIDE.md#8-what-runs-when-the-automation-schedule).

The **cadets** tenant additionally runs the cadet→senior transition lifecycle: `armTransitionTriggers()`
installs five daily triggers staggered 3–7 AM — `detectCadetTransitions` → `resolveTransitionDestinations`
→ `migrateCadetTransitions` → `migrateAllTransitionDrives` → `migrateAllTransitionContacts`. The final
`closeCompletedTransitions()` deletes the old account and is **deliberately not triggered** — it stays a
manual review-then-act step. These belong only on the source (cadets) tenant and must be armed **as the
automation account** ([Accounts & Groups module](src/accounts-and-groups/README.md#4-cadettransitiongs---cadet--senior-account-transition)).

## Modules

| Module | Purpose |
|--------|---------|
| [Accounts & Groups](src/accounts-and-groups/README.md) | Members, email groups, licenses, calendars, Chat spaces, cross-tenant cadet groups, cadet→senior transition |
| [Squadron Groups](src/squadron-groups) | Per-squadron distribution lists (profile-driven set) |
| [Calendar Resources](src/calendar-resources/README.md) | Aircraft/vehicle resources + squadron buildings (seniors only) |
| [Cross-Tenant Contacts](docs/CROSS_TENANT_CONTACTS.md) | Publish the peer tenant's members into this tenant's Global Address List |
| Region features (`src/region`) | Region duty groups + Chat spaces; region-wide unit-visit report (Pacific) |
| Mission Provisioning (`src/mission-provisioning`) | Webhook that provisions a Group + Chat space + Drive folder per mission |
| [Recruiting & Retention](src/recruiting-and-retention/README.md) | Age-out / expiration / welcome emails |
| [Secondary Alias web app](docs/ALIAS_WEB_APP.md) (`webapp/`) | Separate script project: HTML admin UI to add/remove secondary-domain aliases by CAPID |

## Documentation

- **[Administrator & Successor Guide](docs/ADMIN_GUIDE.md)** — the operational runbook (start here for anything live).
- **[New Tenant / New Wing Setup](docs/NEW_TENANT_SETUP.md)** — bare-metal, end-to-end runbook for provisioning a brand-new tenant from nothing (Hawaii-Wing worked example).
- **[Cross-Tenant Contacts](docs/CROSS_TENANT_CONTACTS.md)** — seniors ⇄ cadets shared-contact sync.
- **[Secondary Alias Web App](docs/ALIAS_WEB_APP.md)** — the `webapp/` admin UI: setup, the access model, and the sheet contract it shares with the nightly run.
- **[Pacific Diff](docs/REGION_DIFF.md)** / **[GCP Project Migration](docs/GCP_PROJECT_MIGRATION.md)** — Pacific-specific setup and the standard-GCP-project migration.
- **[Spreadsheet Setup](docs/SPREADSHEET_SETUP.md)** — the automation config spreadsheet tabs.
- **[API Reference](docs/API_REFERENCE.md)** · **[Utilities](docs/UTILITIES.md)** · **[Development Guide](docs/DEVELOPMENT.md)** · **[Troubleshooting](docs/TROUBLESHOOTING.md)** — internals, inherited from the upstream single-wing project and reconciled for the multi-tenant model (each carries a note on what changed).
- **[SECURITY.md](SECURITY.md)** · **[PCR_CHANGELOG.md](PCR_CHANGELOG.md)** / **[CHANGELOG.md](CHANGELOG.md)** · **[Versioning](docs/VERSIONING.md)**.

## Contributing

Contributions are welcome. Branch from `master`, follow the existing code style, add JSDoc and a
`preview…`/`test…` variant for anything that mutates Workspace state, and open a PR to
**`CAP-Pacific-Region/MIWG-PCR-gsuite-automation`** — **not** the upstream `cap-miwg` repo. See
[CONTRIBUTING.md](CONTRIBUTING.md) and [DEVELOPMENT.md](docs/DEVELOPMENT.md).

## License & acknowledgments

MIT — see [LICENSE](LICENSE).

**Pacific Region multi-tenant work:** Noel Luneau (PCR IT) and Isaac Wilson IV (CAWG IT).
**Upstream single-wing automation:** Luke Bunge (lead, v1.0→v2.0) and Jeremy Ginnard (original), Michigan Wing IT; Calendar resources by California Wing IT.
Thanks to CAP National IT for CAPWATCH access, and to the Michigan and California Wing IT teams for testing and feedback.

---

**Built for the Civil Air Patrol Pacific Region IT team.**
*Unofficial project — not endorsed by CAP National.*
