# Administrator & Successor Guide (PCR Deployment)

> **Read this first if you have just inherited this system.** It is the "hit by a bus"
> runbook for the Pacific Region (PCR) deployment of the CAPWATCH → Google Workspace
> automation. It explains what is actually running, where it lives, how to get into it,
> how to keep it alive, and how to rebuild it from scratch if you have to.
>
> This guide is **specific to the PCR fork** (`CAP-Pacific-Region/MIWG-PCR-gsuite-automation`).
> The top-level [README](../README.md) and [DEVELOPMENT.md](DEVELOPMENT.md) are inherited
> from the upstream Michigan Wing project and describe a single-wing, single-domain setup.
> Where they disagree with this document, **this document wins for our deployment.**

---

## Table of Contents

1. [The 60-second overview](#1-the-60-second-overview)
2. [What you need to take over (access checklist)](#2-what-you-need-to-take-over-access-checklist)
3. [System inventory (the facts you cannot guess)](#3-system-inventory-the-facts-you-cannot-guess)
4. [Google Apps Script crash course](#4-google-apps-script-crash-course)
5. [The three tenants and how code gets deployed](#5-the-three-tenants-and-how-code-gets-deployed)
6. [Local development with clasp](#6-local-development-with-clasp)
7. [Secrets and Script Properties (the part that breaks silently)](#7-secrets-and-script-properties-the-part-that-breaks-silently)
8. [What runs, when: the automation schedule](#8-what-runs-when-the-automation-schedule)
9. [Entry-point function reference](#9-entry-point-function-reference)
10. [The code, module by module](#10-the-code-module-by-module)
11. [Mission provisioning webhook](#11-mission-provisioning-webhook)
12. [Known constraints and gotchas](#12-known-constraints-and-gotchas)
13. [Routine maintenance](#13-routine-maintenance)
14. [Troubleshooting playbook](#14-troubleshooting-playbook)
15. [Disaster recovery: rebuild from zero](#15-disaster-recovery-rebuild-from-zero)
16. [Glossary](#16-glossary)

---

## 1. The 60-second overview

Civil Air Patrol membership lives in **CAPWATCH** (the national eServices database).
This system keeps **Google Workspace** in sync with CAPWATCH automatically, so that:

- Every active member has a Workspace account in the correct Organizational Unit (OU).
- Email distribution groups (squadron, wing, duty-position, specialty) stay accurate.
- Expired members are suspended, and eventually deleted, to stay under the license cap.
- Aircraft and vehicles show up as bookable Calendar resources.
- Squadrons get Chat spaces; new members get shared calendars.
- Missions get a Group + Chat space + Drive folder provisioned on demand (via a webhook).

It runs **entirely inside Google Apps Script** — no servers, no cloud VMs. Time-driven
triggers fire the sync functions overnight. The only moving parts are: the Apps Script
projects, a few Google Drive folders/Sheets, and a per-tenant Google Cloud **service account**
used only for the calls that need domain-wide user impersonation (Gmail settings and Calendar).

**The single most important thing to understand:** this is deployed to **three completely
separate Workspace tenants** (seniors, cadets, and Pacific Region), each with its **own**
Apps Script project. The same `src/` code is pushed to all three, but each can drift. See
[Section 5](#5-the-three-tenants-and-how-code-gets-deployed).

---

## 2. What you need to take over (access checklist)

Get every item below **before** you need it. Missing any one of these can lock you out.

- [ ] **Super Admin** on the **seniors** Workspace tenant (`cawgcap.org`).
- [ ] **Super Admin** on the **cadets** Workspace tenant (`cawgcadets.org`).
- [ ] **Super Admin** on the **Pacific Region** tenant (see [inventory](#3-system-inventory-the-facts-you-cannot-guess); verify the domain live).
- [ ] Access to the Google account that **owns each Apps Script project** (they may be owned
      by a shared/role account, not a person — find out which).
- [ ] Access to the **Google Cloud project(s)** that hold each tenant's dedicated **service
      account** (one per tenant; find each via `SA_IMPERSONATION_EMAIL` in that project's Script Properties).
- [ ] The **CAPWATCH eServices credentials** used to download data (a member account with
      CAPWATCH permission for the region ORGID).
- [ ] Owner/editor on the **Drive folders and config Sheets** listed in the inventory.
- [ ] **GitHub** write access to `CAP-Pacific-Region/MIWG-PCR-gsuite-automation`.
- [ ] The **mission webhook shared secret** (or the ability to rotate it) if the FileMaker
      mission integration is in use.

> **If you can only get one thing:** get Super Admin on the seniors tenant and access to
> that tenant's Apps Script project. Everything else can be rebuilt from there.

---

## 3. System inventory (the facts you cannot guess)

These IDs are the load-bearing facts. Keep this table current — it is the map to everything.

### Tenants and Apps Script projects

| Tenant | Domain | Project owner | clasp target | Apps Script `scriptId` |
|--------|--------|---------------|--------------|------------------------|
| Seniors (the driver) | `cawgcap.org` | `automation@cawgcap.org` | `clasp-targets/seniors.clasp.json` | `1ZjkCGQ2Dt-goAYO6n9y6cDwUnvm3Jor6DV0sLIsdCu4iB5zSzS9gmjAi` |
| Cadets | `cawgcadets.org` | `automation@cawgcadets.org` | `clasp-targets/cadets.clasp.json` | `15LWpFVw0qis2XOZBZOo0YL4hMN-eNRGK6EC6yQAeirzrUl-iDcbzjUHc` |
| Pacific Region | `pcr.cap.gov` | `automation@pcr.cap.gov` (project in a `pcr.cap.gov` Shared Drive) | `clasp-targets/region.clasp.json` | `1s2Fmdo0sxWjuPawYBU_dCGYa5qA0h8LuGbQkIGzptzhBlTqL14JqW-T0` |

Open any project in the browser from the repo with e.g. `npm run open:seniors`.

> **Pacific is deployed and verified (2026-07-09).** It runs the shared `src/` with
> `TENANT_PROFILE=region`. Two things to know for future pushes:
> - The project lives in a **`pcr.cap.gov` Shared Drive**. An external-org account (e.g.
>   `automation@cawgcap.org`) can `clasp pull` but gets a 403 *"User has not enabled the Apps
>   Script API"* on **push** — that message is misleading; it's a **cross-org write block**, not
>   an API-toggle problem (reads work, `canEdit` is true). **Push from an account internal to
>   `pcr.cap.gov` — `automation@pcr.cap.gov` (the operator/owner).**
> - `CAPWATCH_AUTHORIZATION` is a per-user **User Property** (not shown in the Script Properties
>   UI) belonging to `automation@pcr.cap.gov`, so that account must also **own the triggers**.
>
> The earlier "owned by `it@pcr.cap.gov`, pull blocked" story was a misdiagnosis — the real cause
> of `Requested entity was not found` was a stale `scriptId` in the clasp target (now corrected).

### Drive / Sheets IDs (seniors tenant — canonical copy in [`config-tenants/seniors.json`](../config-tenants/seniors.json))

| Purpose | Config key | ID |
|---------|-----------|-----|
| CAPWATCH data folder (downloaded `.txt` files land here) | `CAPWATCH_DATA_FOLDER_ID` | `10T0wBubqzUzHa_7nx__eNfuzhTpFRDs3` |
| Automation folder | `AUTOMATION_FOLDER_ID` | `1lLUs0RsTQXNgRnt_fURsw8B3E8DpsgE2` |
| Automation config spreadsheet (`Groups`, `User Additions`, `Error Emails`, `Mission Provisioning`, …) | `AUTOMATION_SPREADSHEET_ID` | `1UqCc6aRMEYw-Y_bTcTDKXuaYLsQ6bQzkdoVG7rRsV9Q` |
| Retention log spreadsheet | `RETENTION_LOG_SPREADSHEET_ID` | `1ouL6YHtTfpJs32YQ2NyfYxjHSDg39RydHMamGHXM7yA` |

> These IDs are the **seniors** tenant's. After the Script-Properties refactor, `config.gs` no
> longer hard-codes any tenant's identity — `getTenantConfig_()` reads it from each project's
> **Script Properties** (`TENANT_*`), so the values above now live in
> [`config-tenants/seniors.json`](../config-tenants/seniors.json), not in `config.gs`. Each tenant
> has its own set: cadets in [`config-tenants/cadets.json`](../config-tenants/cadets.json)
> (ORGID `188`, `cawgcadets.org`), Pacific in
> [`config-tenants/region.json`](../config-tenants/region.json) (ORGID `434`, `pcr.cap.gov`). The
> cadet cross-tenant nesting is **not** addressed via `CONFIG.DOMAIN`; see
> [Section 5](#5-the-three-tenants-and-how-code-gets-deployed). To trust a value for a tenant, read
> that tenant's `config-tenants/*.json` (or its live Script Properties), not the shared `config.gs`.

### Service accounts (Google Cloud)

- **Each tenant has its own dedicated service account.** There is no single shared account,
  and `config.gs` no longer contains any service-account JSON, project id, or client id
  (that builder was removed in the security-hardening pass — commit `848f724`).
- The service account is used only for **per-user impersonation** where an API requires
  domain-wide delegation — specifically **Gmail settings** (signatures/aliases) and
  **Calendar** operations. Its email and private key live in that project's **Script
  Properties** (`SA_IMPERSONATION_EMAIL`, `SA_PRIVATE_KEY`), never in code, and are read at
  runtime by `getImpersonatedToken_()` in `UpdateMembers.gs`, which signs the delegation JWT.
- Domain-wide delegation for each SA should be scoped to **only** the Gmail-settings and
  Calendar scopes the automation needs — not the full scope list.
- **To find a tenant's SA:** read `SA_IMPERSONATION_EMAIL` in that project's Script Properties;
  the domain of that address tells you which Google Cloud project owns it. (Older code
  referenced a `pcr-capwatch` project and client id `117582666328715304137`; that is historical
  and no longer authoritative — verify live.)

### People / contacts (update as roles change)

- Automated mail sender: `automation@cawgcap.org` (display name "CAWG Information Technology").
- IT support mailbox: `it@cawgcap.org`.
- Recruiting/retention: `recruiting@cawgcap.org`, Director `adam.staley@cawgcap.org`.
- Upstream code authors (historical): Luke Bunge, Jeremy Ginnard (Michigan Wing);
  PCR config by Noel Luneau.

---

## 4. Google Apps Script crash course

If you have never used Apps Script, read this section once. It removes 90% of the confusion.

**What it is.** Apps Script is Google's serverless JavaScript platform. A "project" is a
bundle of `.gs` files (server-side JavaScript) and `.html` files, attached to your Google
account, that can call Google APIs (Gmail, Drive, Admin SDK, Calendar, Chat, …) as *you*.
There is no build step and nothing to deploy to run it interactively — you press **Run**.

**Where it lives.** <https://script.google.com>. Each of our three tenants has its own
project (see inventory). Open the editor to see the files, run functions, view logs, and
manage triggers and properties.

**Key concepts you must know:**

- **Functions are the unit of execution.** You pick a function from the toolbar dropdown and
  click **Run**. A function with no parameters can be run directly or attached to a trigger.
  Functions ending in `_` (e.g. `ensureMissionGroup_`) are private helpers by convention —
  you generally don't run them directly.
- **Triggers** are how things run unattended. **Editor → Triggers (clock icon) → Add Trigger.**
  A *time-driven* trigger runs a chosen function on a schedule (hourly, daily in a 1-hour
  window, weekly, monthly). Triggers belong to the **user who created them** and run as that
  user — if that person leaves and their account is deleted, **their triggers stop.** This is
  a classic bus-factor failure; see [Section 13](#13-routine-maintenance).
- **Authorization / OAuth scopes.** The first time a function touches a Google service, Apps
  Script prompts *you* to grant permission. The scopes it can request are declared in
  [`src/appsscript.json`](../src/appsscript.json) (`oauthScopes`). Because we do admin-level
  things, the running user must be a **Super Admin**. If you see "insufficient permission",
  it is almost always a scope-authorization or admin-role problem.
- **Advanced Services.** APIs like Admin SDK Directory, Calendar, Chat, Gmail, Drive v3, and
  Groups Settings are enabled as "advanced services" in `appsscript.json`
  (`enabledAdvancedServices`). When you push code with clasp, these come along in the manifest.
  In the editor they appear under **Services (+)**. Symbols like `AdminDirectory`, `Chat`,
  `Drive`, `Calendar`, `AdminGroupsSettings` come from here.
- **Properties = key/value storage.** `PropertiesService` has two stores we use:
  - **Script Properties** — shared by everyone using the project. This is where **secrets and
    per-tenant flags** live (service account key, mission secret, `MANAGE_RESOURCES`). Manage
    them in **Editor → Project Settings (gear) → Script Properties.**
  - **User Properties** — private to the running user. The **CAPWATCH login token** is stored
    here (per `setAuthorization()`), which means the download only works for the user who set it.
- **Execution limits.** A single execution may run **up to 6 minutes** (Workspace accounts).
  This is why the code processes members/squadrons in **batches** and tracks progress in Script
  Properties (e.g. `SQUADRON_BATCH_INDEX`) so the next run resumes where it left off.
- **Logs.** `console.log`/`Logger.info` output goes to **Executions** (list icon) and to Cloud
  Logging (Stackdriver — `exceptionLogging: "STACKDRIVER"`). Every automated run shows up in
  **Executions** with status Completed/Failed — this is your first stop when something breaks.
- **Web apps.** A project can expose an HTTP endpoint by implementing `doGet(e)`/`doPost(e)`
  and creating a **Deployment → Web app**. We use this for the [mission webhook](#11-mission-provisioning-webhook).
  A web-app deployment has a stable `/exec` URL and runs as a fixed identity — editing the code
  does **not** change the URL, but you must **create a new version/deployment** for changes to
  take effect on that URL (unlike triggers, which run the latest saved code).

**Golden rule:** a plain `clasp push` (or saving in the editor) updates the code that triggers
run **immediately**, but does **not** create a numbered "version" and does **not** update a
web-app deployment. Don't assume "the versions list didn't change" means "nothing changed."

---

## 5. The three tenants and how code gets deployed

There is **one** `src/` directory in this repo. It is deployed, unchanged, to **three**
Apps Script projects via three clasp targets. Each target is just a `{ scriptId, rootDir }` pointer:

```
clasp-targets/
├── seniors.clasp.json   → cawgcap.org project      (seniors tenant + cross-tenant driver)
├── cadets.clasp.json    → cawgcadets.org project   (cadets tenant)
└── region.clasp.json   → Pacific Region project   (pcr.cap.gov Shared Drive; deployed)
```

### How the cadet split *actually* works

The seniors and cadets projects are two **separate, active** tenants that run the **same `src/`**,
differentiated only by Script Properties (`TENANT_*` identity + `TENANT_PROFILE`):

- **Seniors project → `cawgcap.org`.** `CONFIG.DOMAIN` is `cawgcap.org` (from `TENANT_DOMAIN`). It
  manages the seniors tenant and is also the **cross-tenant driver**: `updateCAWGCadetGroups()` runs
  here and writes "User Additions" rows that nest the cadets' groups (which live on `cawgcadets.org`)
  into the matching `cawgcap.org` groups. The cadet domain it nests against is resolved by
  `getCAWGCadetsTenantDomain_()` in `UpdateCAWGCadetGroups.gs` — derived from `WING` (`"CA"` →
  `cawgcadets.org`) unless a `CADETS_TENANT_DOMAIN` Script Property overrides it.
- **Cadets project → `cawgcadets.org`.** `CONFIG.DOMAIN` is `cawgcadets.org` (from `TENANT_DOMAIN`)
  and `TENANT_PROFILE=cadets`, so it processes **CADET** members only, in Cadet-Lite mode, and
  creates the cadet list set (all-hands + cadets + parents). It provisions cadet accounts and groups
  on `cawgcadets.org` directly. **Re-enabled 2026-07-09**; its time-driven triggers run under the
  role account **`automation@cawgcadets.org`**.

**Crossing the split (aging out).** When a cadet turns 21 or converts, they leave the cadets tenant
for the seniors tenant, and their `cawgcadets.org` mailbox would otherwise be deleted with no archive
behind it. The **cadet→senior transition** subsystem (`CadetTransition*.gs`) handles this: the cadets
tenant (`TRANSITION_CONFIG.ROLE=source`) migrates their mail/Drive/contacts to the new
`cawgcap.org` account, then deletes the old account and forwards its address; the seniors tenant
(`ROLE=destination`) just exempts them from the Level I gate so the receiving account exists. It runs
on triggers **except** the final delete, which stays manual. Full detail:
[Section 9](#9-entry-point-function-reference) and the
[module README](../src/accounts-and-groups/README.md#4-cadettransitiongs---cadet--senior-account-transition).

> **History (resolved).** This section used to describe the cadets project as a dormant
> "byte-for-byte clone" pointed at `cawgcap.org` with no triggers. That was a symptom of the
> shared-`config.gs` clobber (below), since fixed: the cadets project now carries its own `TENANT_*`
> Script Properties, so a push no longer repoints it at the seniors domain, and its automation runs
> on schedule again.

> Because the three projects are pushed **independently**, they can **silently run different code**
> from each other and from git. **Do not assume HEAD reflects what is live on any tenant.**
> **Deployment status (2026-07-09):** all three tenants — seniors, cadets, and Pacific — have been
> pushed to the reconciled `master` and run **identical `src/`**, differentiated only by Script
> Properties + `TENANT_PROFILE`. The earlier gap (live tenants behind HEAD on the security-hardening
> pass) is closed. Still, treat this as a point-in-time fact: always confirm before assuming.

### ⚠️ `config.gs` is shared and overwritten on every push — keep per-tenant config in Script Properties

`.claspignore` ships **everything under `src/`**, and `config.gs` lives there; all three targets use
`rootDir: ../src`. So **every `clasp push` overwrites that project's `config.gs` with the shared
copy.** The design accounts for this: `config.gs` is **tenant-neutral** — it hard-codes no domain,
ORGID, folder, or sheet values. `getTenantConfig_()` reads the per-tenant identity (`DOMAIN`,
`EMAIL_DOMAIN`, `CAPWATCH_ORGID`, `WING`, `REGION`, the two folder IDs, the automation + retention
sheet IDs, and the contact emails) from `TENANT_*` **Script Properties**, which pushes never touch,
and `TENANT_PROFILE` selects per-tenant behavior. Canonical non-secret values are version-controlled
in [`config-tenants/`](../config-tenants/).

- **Never hand-edit per-tenant values into `config.gs`** — the next push wipes them. Set a Script
  Property instead (and, if it's a new field, read it via `getTenantConfig_()`). Historically the
  cadets project's `config.gs` went byte-identical to seniors precisely because a push clobbered a
  hand-edited copy; the Script-Property layer is what removes that failure mode.
- **The only durable per-tenant layer is Script Properties** — identity (`TENANT_*`), behavior
  (`TENANT_PROFILE`), plus `MANAGE_RESOURCES`, `SA_*`, mission secrets, and CAPWATCH credentials.

Set them per project with `setupTenantConfig()` (or by hand), verify with `validateTenantConfig()`,
and see [Section 7](#7-secrets-and-script-properties-the-part-that-breaks-silently).

> **⚠️ Migration order — do this before pushing the refactor to any live tenant.** There is
> deliberately **no fallback**: an unconfigured project yields an empty `DOMAIN` and fails loudly
> rather than acting on the wrong tenant. So on each existing project, **set its `TENANT_*`
> properties first** (values in `config-tenants/`), run `validateTenantConfig()`, and **only then**
> `clasp push`. Push to a project with unset properties and its automation will halt until you set them.

**Deploy workflow (normal change):**

```bash
npm run push:seniors    # push src/ to the seniors project
npm run push:cadets     # then cadets
npm run push:region    # then pacific
```

Push to one, confirm it's healthy (run a preview function, check Executions), then the next.
Never push all three blind — and remember each push resets the target's `config.gs` to the shared values.

---

## 6. Local development with clasp

`clasp` is Google's CLI for Apps Script. It lets you edit `.gs`/`.html` in git and push/pull
to a project. Everything is wired through npm scripts in [`package.json`](../package.json).

### One-time setup

```bash
npm install -g @google/clasp     # install clasp globally
clasp login                      # opens a browser; log in as an owner of the projects
```

Also enable the Apps Script API for your account (one-time, per Google account):
<https://script.google.com/home/usersettings> → **Google Apps Script API: ON**.
Without this, `clasp push`/`pull` fail with an API-disabled error.

### Everyday commands

| Action | Seniors | Cadets | Pacific |
|--------|---------|--------|---------|
| Status (what would change) | `npm run status:seniors` | `npm run status:cadets` | `npm run status:region` |
| Push (deploy `src/` → project) | `npm run push:seniors` | `npm run push:cadets` | `npm run push:region` |
| Pull (project → `src/`) | `npm run pull:seniors` | `npm run pull:cadets` | `npm run pull:region` |
| Open in browser | `npm run open:seniors` | `npm run open:cadets` | `npm run open:region` |

Under the hood each is `clasp <cmd> -P clasp-targets/<tenant>.clasp.json`.

> **Filenames in Apps Script are not cosmetic.** A file at
> `src/recruiting-and-retention/WelcomeEmail.html` deploys with the literal name
> `recruiting-and-retention/WelcomeEmail` (slash included). Any
> `HtmlService.createTemplateFromFile('...')` / `createHtmlOutputFromFile('...')` call must use
> that **exact full slash-prefixed name**, or it throws "file not found." If you rename or move
> HTML files, fix every reference.

### Recommended change process

1. Branch: `git checkout -b fix/thing`.
2. Edit `src/` locally.
3. `npm run status:seniors` to preview the diff, then `npm run push:seniors`.
4. In the editor, run the relevant **preview** function (see below) and check **Executions**.
5. Push to cadets and pacific once seniors looks good.
6. Commit and open a PR to **`CAP-Pacific-Region/MIWG-PCR-gsuite-automation` master**
   (do **not** PR to the upstream `cap-miwg` repo).

---

## 7. Secrets and Script Properties (the part that breaks silently)

Secrets are **never** committed. They live in each project's **Script Properties** and
**User Properties**. If these are missing or wrong, functions fail in confusing ways
(missing-credential errors, 401s, "not configured" errors). Here is the complete inventory.

### Script Properties (per project — set in Project Settings → Script Properties)

| Key | Used by | Purpose |
|-----|---------|---------|
| `TENANT_PROFILE` | `config.gs` (`PROFILE_`) | **Per-tenant behavior selector:** `seniors` (default if unset) or `cadets`. Picks `MEMBER_TYPES.ACTIVE`, `CADET_LITE`, and the squadron-group set (access/public auto-create + distribution-list types). Set to `cadets` on the cadets project. |
| `TENANT_DOMAIN`, `TENANT_EMAIL_DOMAIN`, `TENANT_CAPWATCH_ORGID`, `TENANT_WING`, `TENANT_REGION` | `getTenantConfig_()` → `CONFIG` | **Per-tenant identity.** The Workspace domain, email domain, CAPWATCH ORGID, wing code, and region. First seven are required; `validateTenantConfig()` checks them. |
| `TENANT_CAPWATCH_DATA_FOLDER_ID`, `TENANT_AUTOMATION_FOLDER_ID`, `TENANT_AUTOMATION_SPREADSHEET_ID` | same | **Per-tenant** Drive data folder, automation folder, and automation spreadsheet IDs. Required. |
| `TENANT_RETENTION_LOG_SPREADSHEET_ID`, `TENANT_RETENTION_EMAIL`, `TENANT_DIRECTOR_RECRUITING_EMAIL`, `TENANT_AUTOMATION_SENDER_EMAIL`, `TENANT_SENDER_NAME`, `TENANT_TEST_EMAIL`, `TENANT_ITSUPPORT_EMAIL` | same | **Per-tenant** retention log + contact/sender addresses. `TENANT_SENDER_NAME` defaults to "CAP Information Technology" if unset. |
| `TENANT_WING_ABBREVIATION`, `TENANT_WING_NAME`, `TENANT_CADETS_TENANT_DOMAIN` | `getTenantConfig_()` → `CONFIG` | **Optional, all blank-derive.** Display forms of the wing used in automation emails and the member-facing transition email: abbreviation (`CA` → `CAWG`, `HI` → `HIWG`) and proper name (`California Wing` / `Hawaii Wing`, via `WING_NAMES_`). `TENANT_CADETS_TENANT_DOMAIN` is the peer cadet domain for `updateCAWGCadetGroups()` (blank derives `<wing>wgcadets.org`). Set any of these only to override the derived default. |
| `SA_IMPERSONATION_EMAIL` | `getImpersonatedToken_()` in `UpdateMembers.gs` | Dedicated service account's `client_email` for this tenant. |
| `SA_PRIVATE_KEY` | same | Service account private key (PEM). Literal `\n` sequences are converted to real newlines at runtime, so pasting the one-line form works. |
| `SA_PRIVATE_KEY_ID` | same | Optional key id. |
| `MANAGE_RESOURCES` | `UpdateResources.gs` | Set to `false` to disable Calendar resource management on this tenant. **Cadets = `false`** (resources are managed only by seniors). Absent/anything-else = enabled. |
| `MISSION_WEBHOOK_SECRET` | `MissionProvisioning.gs` | Shared secret the FileMaker webhook must send in the JSON body. Required for the webhook to accept requests. |
| `MISSION_PARENT_FOLDER_ID` | `MissionProvisioning.gs` | Default Drive parent folder for provisioned mission folders (can be overridden per request). |
| `SQUADRON_BATCH_INDEX` | `SquadronGroups.gs` | **Managed automatically** — batch progress cursor. Delete it (or run `resetBatchProgress()`) to restart squadron-group processing from the beginning. |

### User Properties (private to the running user)

| Key | Set by | Purpose |
|-----|--------|---------|
| `CAPWATCH_AUTHORIZATION` | `setAuthorization()` | Base64 `username:password` for the eServices CAPWATCH API. **Only the user who set it can download.** If that user's account is gone, `getCapwatch()` fails until a current admin re-runs `setAuthorization()`. |

### How to set the CAPWATCH credential (do this on each tenant that downloads)

1. Open the project, open `src/GetCapwatch.gs`.
2. In `setAuthorization()`, temporarily set `username` and `password` to valid eServices creds.
3. Run `setAuthorization()` once.
4. **Immediately delete the username/password from the code** and save. The token is now in
   your User Properties; the plaintext must not remain in source.
5. Verify with `testGetCapwatch()`.

### ⚠️ Standing security item: leaked service-account key

A service-account private key was previously committed to **public git history** in this repo.
Removing it from current files (the security-hardening pass deleted the in-`config.gs` builder
that carried it) does **not** remove it from history. **Action required:** rotate the exposed key
in whichever Google Cloud project owns that service account (create a new key, update
`SA_PRIVATE_KEY` in the affected project's Script Properties, delete the old key). Until rotated,
treat the exposed key as compromised. See [SECURITY.md](../SECURITY.md).

---

## 8. What runs, when: the automation schedule

Each tenant's project has its **own** time-driven triggers (Editor → Triggers). Timing matters:
CAPWATCH has an overnight blackout, so downloads run after ~4 AM, and everything that consumes
the data runs after the download. Apps Script schedules within a 1-hour window, not to the minute.

| Order | Function | Suggested cadence | Why |
|------:|----------|-------------------|-----|
| 1 | `getCapwatch()` | Daily, 4–5 AM | Download CAPWATCH ZIP → Drive; also calls `syncOrgPaths()`. |
| 2 | `updateAllMembers()` | Daily, 5–6 AM | Create/update Workspace accounts from member data. |
| 3 | `suspendExpiredMembers()` | Daily, 5–6 AM | Suspend members expired past the 7-day grace window. |
| 4 | `updateEmailGroups()` | Daily, 5–6 AM | Sync wing/duty/specialty distribution groups. |
| 5 | `updateAllSquadronGroups()` | Daily, 6–7 AM | Squadron all-hands/cadets/seniors/parents + public-contact + access groups. Batches via `SQUADRON_BATCH_INDEX`. |
| 6 | `updateAdditionalGroupMembers()` | Daily, 6–7 AM | Merge manual additions from the `User Additions` sheet. |
| 7 | `addSecondaryDomainAliases()` | Daily, 7–8 AM | Give accounts on the `Secondary Aliases` tab a matching address on the tenant's secondary domain. **Seniors only** (`TENANT_SECONDARY_EMAIL_DOMAIN`). Must run *after* `updateAllMembers()` so a same-morning account already exists. |
| 8 | `updateAllSendAsNames()` | Daily, 8–9 AM | Sync Directory displayName **and** Gmail Send-As names (primary + org-owned aliases) from CAPWATCH, so promotions propagate. Resumable: time-boxes at 25 min and self-schedules a continuation, so one firing may span several executions. |
| 9 | `syncMemberCalendarsDaily()` | Daily | Share unit calendars with new/transferred members. |
| 10 | `updateChatSpaces()` | Daily or weekly | Sync squadron/committee Chat spaces & membership. |
| 11 | `updateResources()` | Weekly (Sun) | Aircraft/vehicle Calendar resources + squadron buildings. **Seniors only** (`MANAGE_RESOURCES`). |
| 12 | `manageLicenseLifecycle()` | Monthly | Reactivate renewed, delete long-ineligible accounts to free seats. |
| 13 | `notifyLSCodeChanges()` | Weekly (Mon) | Mail unit CCs when a member's FBI background check is granted or lapses. **Seniors only.** The cadence *is* the reported detection window. |
| 14 | `notifyRecoveryEmailCompliance()` | Monthly (1st) | Mail unit CCs + personnel officers + deputies about members whose email record blocks a password reset. **Seniors + cadets.** |

> ⚠️ **Both notification triggers must be created while signed in as the automation
> account.** A time-driven trigger runs as whoever created it, and the digests send as
> `AUTOMATION_SENDER_EMAIL` — which Gmail permits only for an account owning that verified
> Send-As alias. Created under any other identity (e.g. `it@`), *every* digest fails with
> "Invalid argument". This happened live on 2026-07-16. Both modules' IT failure-summaries
> deliberately send *without* a `from` override so the alarm still gets through when the
> identity is wrong — an "attention needed" email listing every unit as failed is the
> signature of exactly this misconfiguration.

`updateCAWGCadetGroups()` (cross-tenant cadet groups) and the recruiting/retention emails run on
their own cadence where enabled. **Confirm the actual triggers in each project** — this table is
the intended design, not a guarantee of what is currently scheduled.

> The **cadets** tenant runs largely the same schedule **except** resource management
> (`MANAGE_RESOURCES=false`), **plus** the cadet→senior transition lifecycle: six daily
> triggers staggered 3–8 AM (`detectCadetTransitions` → `resolveTransitionDestinations`
> → `migrateCadetTransitions` → `migrateAllTransitionDrives` → `migrateAllTransitionContacts`
> → `remindPendingTransitionCloses`), installed by `armTransitionTriggers()`. The
> account-deleting `closeCompletedTransitions()` is **deliberately not scheduled** — it stays
> a manual review-then-act step, and `remindPendingTransitionCloses` just emails IT when
> accounts are past grace and due for that manual close (see [Section 9](#9-entry-point-function-reference)).
> The **Pacific** tenant runs a **leaner core** (single unit
> PCR-PCR-001 → no squadron groups, no org-path sync, no resources) **plus its own region
> features**, enabled by profile flags that are `false` on the wing tenants:
> `updateRegionGroupChats()` (`RUN_REGION_GROUP_CHATS`), `buildRegionUnitVisitReport()`
> (`RUN_UNIT_VISIT_REPORT`), and `runExternalContactsToDomainSharedContacts()`
> (`RUN_SHARED_CONTACTS`). The code ships everywhere; only Pacific schedules and runs them.

---

## 9. Entry-point function reference

These are the functions worth running by hand or attaching to a trigger. Anything not listed is
an internal helper. **Preview/test functions never modify Workspace** — use them freely.

### Data download
- `getCapwatch()` — download CAPWATCH ZIP into the data folder; refreshes `OrgPaths` too.
- `testGetCapwatch()` — same, wrapped in error logging for a manual smoke test.
- `setAuthorization()` — store eServices credentials (run once; see [Section 7](#7-secrets-and-script-properties-the-part-that-breaks-silently)).

### Members & accounts (`UpdateMembers.gs`)
- `updateAllMembers()` — main account sync (create/update, OU placement, aliases, custom schema).
- `suspendExpiredMembers()` — suspend members past the grace period.
- `reactivateRenewedMembers()` — unsuspend members who became active again.
- `addAliasesFromSheet()` / `updateMissingAliases`-style helpers — alias repair.
- `getMembers()`, `getSquadrons()`, `getAEMembers()` — data builders (used by tests/other modules).

### Org units / OU paths (`SyncOrgPaths.gs`)
- `syncOrgPaths()` — rebuild the ORGID→OU mapping and **auto-create OUs** for newly activated
  squadrons; emails a summary of added/needs-manual/deactivated units.

### Email groups (`UpdateGroups.gs`)
- `updateEmailGroups()` — sync configured groups from the `Groups` sheet.
- `updateAdditionalGroupMembers()` — apply manual `User Additions`.
- `getEmailGroupDeltas()` — **preview** adds/removes without applying (inspect before running).

### Squadron groups (`SquadronGroups.gs`)
- `updateAllSquadronGroups()` — full squadron group sync (batched).
- `createAllSquadronGroups()` — create missing squadron groups.
- `updateSquadronGroupsBatch(batchSize)` / `checkBatchStatus()` / `resetBatchProgress()` — batch control.
- `previewSquadronGroups()`, `testPreviewSquadronGroups()`, `listAvailableSquadrons()` — **preview/inspect.**
- `updateSingleSquadronGroups(unitNumber)` — operate on one unit.
- Which list types get created is **tenant-driven** via `PROFILE_.SQUADRON_DISTRIBUTION_TOGGLES`
  (`config.gs`, per `TENANT_PROFILE`): seniors = all-hands / cadets / seniors / parents + command
  staff; cadets = all-hands / cadets / parents (no `.seniors` or command-staff); pacific = none.
  Groups left behind after a toggle is disabled are staged for cleanup by
  `groupAdministration_stageOrphanedSquadronGroups()`.

### Region features (Pacific only — profile-gated)
These ship in the shared `src/` but no-op unless their profile flag is on (`true` only for
`TENANT_PROFILE=region`; `false`/unscheduled on the wing tenants):
- `updateRegionGroupChats()` — region duty groups + duty Chat spaces, built from the region
  CAPWATCH folder (`RUN_REGION_GROUP_CHATS`; `TENANT_REGION_CAPWATCH_DATA_FOLDER_ID`).
- `buildRegionUnitVisitReport()` — region-wide unit-visit spreadsheet, one tab per wing
  (`RUN_UNIT_VISIT_REPORT`; destination sheet/calendar from `TENANT_UNIT_VISIT_*` properties).
- `runExternalContactsToDomainSharedContacts()` — sync the `External Contacts` sheet into Domain
  Shared Contacts (`RUN_SHARED_CONTACTS`; requires the `contacts` OAuth scope).
- `cleanupUnnecessaryDistributionLists(dryRun=true)` — **defaults to dry run**; pass `false` to delete.

### Cadet cross-tenant groups (`UpdateCAWGCadetGroups.gs`)
- `updateCAWGCadetGroups()` — build cadet group structure that spans tenants.
- `previewCAWGCadetGroups()` — **preview.**

### Chat spaces (`UpdateChatSpaces.gs`)
- `updateChatSpaces()` — sync unit/committee Chat spaces and membership.

### Calendars (`UpdateCalendars.gs`)
- `initializeCalendarsInfrastructure()` — one-time/idempotent setup of calendar sharing infra.
- `syncMemberCalendarsDaily()` — share calendars with new/transferred users.
- `runWriterGroups()` — manage writer-access groups.

### Calendar resources (`UpdateResources.gs`) — seniors only
- `updateResources()` — sync aircraft/vehicles as resources + squadron buildings.
- `previewAircraftResources()`, `previewVehicleResources()`, `testAcCodeResolution()` — **preview/test.**

### License lifecycle (`ManageLicenses.gs`)
- `manageLicenseLifecycle()` — the monthly orchestrator (reactivate + delete ineligible).
- `previewLicenseLifecycle()`, `previewArchival()`, `previewDeletion()`, `previewIneligibleMembers()` — **preview.**
- `getLicenseStatistics()` / `testGetLicenseStats()` — current seat usage.
- `suspendIneligibleMembers()`, `deleteIneligibleSuspendedUsers()` — the enforcement steps.
- `auditWorkspaceUsersForRemoval()`, `deleteIneligibleWorkspaceUsers(dryRun=true)` — audit tooling (**dry run by default**).
- `manualReactivateArchivedUser(email)` — one-off manual reactivation.

> **Arming real deletion — `LICENSE_DELETION_ARMED`.** Every deletion path stays a **dry run
> until the per-tenant Script Property `LICENSE_DELETION_ARMED` is exactly `true`** — even when
> a caller passes `dryRun=false`. `manageLicenseLifecycle()` itself passes `dryRun=!armed`, and
> `deleteIneligibleSuspendedUsers()` / `deleteIneligibleWorkspaceUsers()` re-check the property
> before any `Users.remove()`. It is a **property, not a code flag**, on purpose: `clasp push`
> overwrites all of `src/`, so a committed `false→true` edit would be reverted by the next push
> **and** would arm every tenant at once. A property survives push and is per-project, so each
> tenant is armed deliberately and independently.
>
> **To arm one tenant:** Project Settings → Script Properties → `LICENSE_DELETION_ARMED = true`.
> **Always dry-run first** — `deleteIneligibleSuspendedUsers(true)` — and read the buckets:
> `WOULD DELETE`, `SPARED within grace`, `SPARED current-member-ineligible-by-type` (PATRON etc.,
> a human call), and `SPARED left-the-wing` (a member absent from CAPWATCH but recently active is
> treated as a **wing transfer**, not a lapse, and put on a departure timer — this guard is what
> stops a transferred member being wrongly deleted). Grace is 30 days from the CAPWATCH
> Expiration date; deletion is permanent (this edition has no Archived-User fallback).
> **Deploy status:** armed on **seniors** (2026-07-17); **cadets** and **region** remain dry-run.

### Cadet → senior transition (`CadetTransition*.gs`) — **cadets tenant only**
Carries a converting cadet's mail, Drive, and contacts to their new seniors-tenant
account before the cadet account is deleted (`TRANSITION_CONFIG.ROLE === 'source'`;
no-ops elsewhere). State lives in the `Transitions` sheet. See the
[module README](../src/accounts-and-groups/README.md#4-cadettransitiongs---cadet--senior-account-transition).
- `armTransitionTriggers()` — install the six daily lifecycle triggers (detect →
  migrate → remind, 3–8 AM). **Run as `automation@cawgcadets.org`** — triggers are owned by
  their creator and the completion email's Send-As is that account.
  `disarmTransitionTriggers()` / `listTransitionTriggers()` manage/inspect them.
- `detectCadetTransitions()`, `resolveTransitionDestinations()` — open rows, resolve destinations.
- `migrateCadetTransitions(notify)`, `migrateAllTransitionDrives()`, `migrateAllTransitionContacts()` — the three migration phases (resumable).
- `previewCadetTransitions()`, `previewCadetTransitionMigration()`, `previewSingleTransitionDrive(capid)`, `previewSingleTransitionContacts(capid)` — **preview** the queue and prove credentials without copying.
- `remindPendingTransitionCloses()` — emails IT (`TENANT_ITSUPPORT_EMAIL`) when accounts have
  passed grace and are ready for the manual close (or are stuck past grace on a `DO NOT DELETE`
  hold). Read-only; silent when nothing is due. Runs daily at 08:00.
- `closeCompletedTransitions(dryRun=true)` — **the only destructive step; not triggered.**
  Deletes the cadet account and forwards its old address (12-month Group) once migration is
  COMPLETE and the hold has expired — **14 days after a verified migration**, or the full
  90-day `HOLD_DAYS` for rows never migrated (e.g. a member who stayed PATRON). Review with
  `(true)`, act with `(false)`.

### Group administration utilities (`groupAdministration.gs`)
- Reporting: `groupAdministration_writeAllGroupsReport()`, `..._writeNoMemberGroupsReport()`,
  `..._previewStaleGroups()`, `..._previewConfiguredGroups()`.
- Audit / cleanup prep (read-only, or writes a review sheet only): `..._auditReceiveListPosting()`
  — flags managed `.cadets`/`.parents`/`.all` receive lists whose `whoCanPostMessage` would reject
  cross-tenant fan-out from a wing `.all` list; run it on the tenant that owns them (e.g. cadets).
  `..._stageOrphanedSquadronGroups()` — writes squadron groups whose list type is disabled for this
  tenant (via `SQUADRON_DISTRIBUTION_TOGGLES`) to the "Delete Groups" tab for review before deletion.
- Destructive (double-check before running): `..._deleteGroup()`, `..._bulkDeleteGroupsFromSheet()`,
  `..._clearGroup()`, `..._deleteConfiguredGroups()`, `..._hideGroupFromGal()`.

### Mission provisioning (`MissionProvisioning.gs`)
- `doPost(e)` / `doGet(e)` — the web-app webhook (not run manually).
- `testMissionProvisioningPayload_()` — provision a `TEST-001` mission locally to verify wiring.

### Notifications to command staff (`notifications/`)
Both modules mail unit command staff, keep their **own** Drive state file, and run on their
**own** trigger — a failure in either cannot affect account provisioning. Both trigger must be
owned by the automation account (see the warning in [Section 8](#8-what-runs-when-the-automation-schedule)).

- `previewLSCodeChanges()` / `notifyLSCodeChanges()` — FBI background-check changes.
  **Seniors only** (`RUN_LSCODE_NOTIFICATIONS`). First run is **silent by design** (it lays a
  baseline); only a *change* against recorded state notifies. `installLSCodeWeeklyTrigger()`.
- `previewRecoveryEmailCompliance()` / `notifyRecoveryEmailCompliance()` — members whose
  CAPWATCH email record would stop them resetting their password: no CAP address in the
  PRIMARY slot, and/or no personal (non-CAP) address on file at all. **Seniors + cadets**
  (`RUN_RECOVERY_EMAIL_NOTIFICATIONS`). Sent to the unit commander, copying its personnel
  officers (primary *and* assistant) and deputy commanders. `installRecoveryComplianceMonthlyTrigger()`.
  - ⚠️ Unlike LSCode, the **first run is deliberately loud** — it reports a standing condition,
    so it surfaces the entire existing backlog at once. **Run the preview first and look at the
    volume** before scheduling it.
  - A member reported once is not reported again for **3 months**, even if uncorrected. A member
    who fixes the issue drops out of state, so a later relapse is reported immediately.
  - Cadet-lite members are excluded automatically (no account ⇒ no password). A cadet whose own
    secondary email is empty but who has a **parent/guardian** address on file is *not* flagged
    for recovery — that address is a working reset path.
  - On the **cadets** tenant this requires `TENANT_COMMAND_EMAIL_DOMAIN` = the senior domain
    (`@cawgcap.org`): a cadet unit's commander and personnel officer are seniors and hold no
    cadet-domain account, so without it every digest is addressed to an account that
    does not exist.
- `testRecoveryDigestToTestEmail()` / `testLSCodeDigestToTestEmail()` — render a digest from
  fabricated data to `TENANT_TEST_EMAIL`, to review wording without waiting for a real run.

**Rule of thumb:** for anything that changes or deletes Workspace state, there is a `preview…`,
`test…`, or `dryRun=true` variant. **Always run the preview first** and sanity-check the counts.

---

## 10. The code, module by module

```
src/
├── config.gs                 # All configuration constants (domain, ORGID, folder/sheet IDs). Start here.
├── utils.gs                  # parseFile (CSV cache), executeWithRetry, email/validation helpers.
├── GetCapwatch.gs            # Download CAPWATCH ZIP → Drive; credential storage.
├── SyncOrgPaths.gs           # ORGID→OU map; auto-creates OUs for new squadrons.
├── appsscript.json           # Manifest: timezone, advanced services, OAuth scopes, web-app config.
│
├── accounts-and-groups/
│   ├── UpdateMembers.gs           # Account create/update/suspend/reactivate; aliases; custom schema.
│   │                              # Also getImpersonatedToken_(): signs the SA delegation JWT
│   │                              # (reads SA_* from Script Properties) for Gmail-settings/Calendar.
│   ├── UpdateGroups.gs            # Configured email groups (Groups sheet) + manual additions.
│   ├── UpdateCAWGCadetGroups.gs   # Cross-tenant cadet group structure.
│   ├── UpdateChatSpaces.gs        # Unit/committee Chat spaces (Chat advanced service, runs as executing user).
│   ├── UpdateCalendars.gs         # Share unit calendars with members; writer groups.
│   ├── ManageLicenses.gs          # Lifecycle: reactivate/suspend/delete to manage seats.
│   ├── groupAdministration.gs     # Ad-hoc group reporting & cleanup utilities.
│   └── CadetTransition*.gs         # Cadet→senior cross-tenant transition (cadets tenant only):
│                                  #   CadetTransition.gs (detect/state/triggers),
│                                  #   ...Migrate.gs (Gmail), ...Drive.gs, ...Contacts.gs,
│                                  #   ...Cleanup.gs (manual delete + forward).
│
├── squadron-groups/
│   └── SquadronGroups.gs          # Per-squadron all-hands/cadets/seniors/parents + public-contact
│                                  # + access groups. Batched via SQUADRON_BATCH_INDEX.
│
├── calendar-resources/
│   └── UpdateResources.gs         # Aircraft/vehicle resources + squadron buildings (seniors only).
│
├── mission-provisioning/
│   └── MissionProvisioning.gs     # doPost webhook: Group + Chat space + Drive folder per mission.
│
├── notifications/                 # Push signals to unit command staff. Each keeps its OWN state
│   │                              # file and trigger, and never touches provisioning.
│   ├── LSCodeNotify.gs            # Weekly: FBI background-check (Member.txt LSCode) changes.
│   │                              # Seniors only — cadet records carry no LSCode.
│   └── RecoveryEmailNotify.gs     # Monthly: members whose CAPWATCH email record blocks a
│                                  # password reset. Seniors + cadets. 3-month per-member quiet
│                                  # period; needs TENANT_COMMAND_EMAIL_DOMAIN on cadets.
│
└── recruiting-and-retention/
    ├── SendRetentionEmail.gs      # Age-out / expiration / welcome emails.
    └── *.html                     # Email templates (exact slash-prefixed names matter!).
```

**Data flow:** `getCapwatch()` writes CAPWATCH `.txt` files to Drive → `utils.parseFile()` reads
and caches them → `getMembers()`/`getSquadrons()` turn rows into objects → the update modules push
those to Workspace via the Admin SDK / Calendar / Chat APIs. Change detection uses a saved snapshot
so only changed members are re-written each night.

For deeper reference on data structures, member objects, caching, retry logic, and coding standards,
see [DEVELOPMENT.md](DEVELOPMENT.md) and [UTILITIES.md](UTILITIES.md) (inherited upstream — accurate
on internals, but written for the single-wing setup).

---

## 11. Mission provisioning webhook

A FileMaker system POSTs mission data to the Pacific (or relevant) project's **web app**; the
webhook creates a Google Group, a Chat space, and a Drive folder for that mission, and records the
result in the `Mission Provisioning` sheet of the automation spreadsheet.

- **Endpoint:** the project's web-app `/exec` URL (Deploy → Manage deployments → Web app).
- **Auth:** shared secret in the JSON **body** as `secret`, compared in constant time against the
  `MISSION_WEBHOOK_SECRET` Script Property. The secret must **never** be put in the query string
  (URLs get logged). Requests without a valid secret get `Unauthorized`.
- **Required field:** `missionId`. Recommended: `missionName`, `missionCode`, `ownerEmail`,
  `parentFolderId` (falls back to `MISSION_PARENT_FOLDER_ID`).
- **Idempotent:** a `missionId` already marked `SUCCESS` returns the existing resources rather than
  re-creating them. Failures are recorded with status `ERROR` and the error message.
- **Concurrency:** guarded by a `LockService` script lock so two simultaneous POSTs can't double-create.

**To (re)deploy the webhook after code changes:** editing the code is not enough — create a **new
web-app version/deployment** (Deploy → Manage deployments → edit → new version) so the `/exec` URL
serves the new code. Keeping the same deployment preserves the URL FileMaker already points at.

**To test without FileMaker:** run `testMissionProvisioningPayload_()` in the editor (provisions a
`TEST-001` mission). Clean up the test group/space/folder afterward.

---

## 12. Known constraints and gotchas

These are the non-obvious things that will bite a successor. Read them before you debug.

1. **Three tenants can drift.** `master` is not ground truth for any live tenant. Verify per-project
   file lists / Project History. (See [Section 5](#5-the-three-tenants-and-how-code-gets-deployed).)

2. **`config.gs` in git = seniors.** Cadets/Pacific use different domain, ORGID, and folder/sheet
   IDs held in their own project copies. Read the live config before trusting a value.

3. **Workspace for Nonprofits license cap.** The domain has a **2,000-user cap** that counts **all**
   accounts regardless of suspension. Suspending a member does **not** free a seat — only **deletion**
   does. This is why `manageLicenseLifecycle()` eventually deletes long-ineligible accounts
   (`DAYS_BEFORE_DELETE_INELIGIBLE = 30`).

4. **Archiving is a no-op here.** `AdminDirectory.Users.update({archived:true})` returns **412** on
   this edition — Archived-User licenses aren't provisioned. `archiveLongSuspendedUsers()` and the
   archive→delete path are effectively disabled. Don't rely on archiving to reclaim seats; deletion is
   the only lever.

5. **HTML template filenames include the folder path with a slash.** `WelcomeEmail.html` in a subfolder
   deploys as `recruiting-and-retention/WelcomeEmail`. `HtmlService` calls must use that exact name.

6. **`clasp push` ≠ new version ≠ web-app redeploy.** Push updates trigger-executed code immediately
   but does not bump the Versions list or update a web-app deployment. Redeploy the web app explicitly.

7. **Triggers are owned by a user.** If the creating account is deleted, its triggers silently stop.
   Prefer creating triggers under a durable role/service account and document who owns them.

8. **CAPWATCH auth is per-user.** The download token lives in User Properties. A departed admin's token
   dies with their account; re-run `setAuthorization()` as a current admin.

9. **Leaked SA key in git history** — rotate it. See [Section 7](#7-secrets-and-script-properties-the-part-that-breaks-silently).

10. **PCR/NHQ members** are added via the **Manual Members / User Additions** sheet and are **never**
    auto-suspended — the lifecycle logic explicitly protects manual members.

11. **6-minute execution limit** drives the batching. If a job stops mid-way, it's usually the timeout;
    check `SQUADRON_BATCH_INDEX` / `checkBatchStatus()` and let the next scheduled run continue, or run
    the batch function again.

12. **A few `cawgcap.org` literals remain outside the automated path** (audited when config was moved to
    Script Properties). These are **not** in the nightly sync, so they were intentionally left, but be
    aware on non-seniors tenants: the `GROUP_ADMINISTRATION_STALE_GROUP_EMAILS` list in
    `groupAdministration.gs` (operator input for the manual stale-group cleanup utilities), the
    `automation@pcr.cap.gov` in `testMissionProvisioningPayload_()` (test helper), and the help-desk
    intranet link in `recruiting-and-retention/WelcomeEmail.html` (shown to new members). The daily
    OrgPath sync email was fixed to use this tenant's IT mailbox (`getOrgPathSyncEmail_()` → `ITSUPPORT_EMAIL`).

---

## 13. Routine maintenance

**Daily (2 minutes):**
- Skim **Executions** in each project for red/Failed rows.
- Check the `Error Emails` sheet and the IT mailbox (`it@cawgcap.org`) for reports.
- Confirm CAPWATCH files in the data folder have today's date (proves `getCapwatch()` ran).

**Monthly:**
- Review the license-management report email; confirm deletions look sane before trusting them.
- Run `getLicenseStatistics()` and watch headroom under the 2,000 cap.
- Verify triggers still exist and are owned by a durable account (not a person about to leave).

**Quarterly / on personnel change:**
- Re-run `setAuthorization()` if the CAPWATCH credential owner changed.
- Re-verify Super Admin access across all three tenants and Apps Script project ownership.
- Rotate the service-account key and the mission webhook secret if anyone with access has left.
- Reconcile the three tenants: `clasp pull` each and diff against `master`; re-push to align.

**When you change code:** push seniors → verify → cadets → pacific; update
[CHANGELOG.md](../CHANGELOG.md) / [PCR_CHANGELOG.md](../PCR_CHANGELOG.md); PR to the PCR repo.

---

## 14. Troubleshooting playbook

| Symptom | Likely cause | Fix |
|---------|--------------|-----|
| `getCapwatch()` fails / no fresh files | `CAPWATCH_AUTHORIZATION` missing or owner's account gone; eServices blackout | Re-run `setAuthorization()` as a current admin; retry after ~4 AM. |
| "insufficient permission" / 403 on user/group ops | Running user isn't Super Admin, or scope not authorized | Ensure Super Admin; re-run the function to re-trigger the OAuth consent for `appsscript.json` scopes. |
| Gmail-settings / Calendar impersonation fails ("Missing service account credentials…") | `SA_IMPERSONATION_EMAIL` / `SA_PRIVATE_KEY` unset or malformed for this tenant | Set them in that project's Script Properties; ensure `\n` in the key is preserved (`getImpersonatedToken_()` un-escapes `\n`). |
| Users created in wrong/no OU | `OrgPaths` mapping stale or OU missing | Run `syncOrgPaths()`; check its email for "needs manual" units; create OUs in Admin Console. |
| Squadron groups stop partway | 6-minute timeout mid-batch | `checkBatchStatus()`; let the next run resume, or `resetBatchProgress()` to restart. |
| Aircraft/vehicles missing from Calendar | `MANAGE_RESOURCES=false`, or Directory Resource API not enabled | Enable on seniors only; verify advanced services; delete `AirportCache.json` and re-run to refresh airport names. |
| Mission webhook returns `Unauthorized` | Secret mismatch or sent in query string | Confirm FileMaker sends `secret` in the JSON body matching `MISSION_WEBHOOK_SECRET`. |
| Mission webhook code changes not taking effect | Web app not redeployed | Create a new web-app version/deployment. |
| Seats exhausted despite suspensions | Suspension doesn't free seats; archiving is a no-op | Rely on deletion of ineligible accounts; run `previewLicenseLifecycle()` then `manageLicenseLifecycle()`. |
| Retention/HTML email "file not found" | Wrong HTML template name | Use the full slash-prefixed filename in `HtmlService` calls. |

Also see the fuller [TROUBLESHOOTING.md](TROUBLESHOOTING.md).

---

## 15. Disaster recovery: rebuild from zero

If a project is lost/corrupted but you still have Super Admin and the Drive data:

1. **Recover access.** Confirm Super Admin on the tenant and Apps Script project ownership.
   If the project itself is gone, create a new Apps Script project (standalone, in the
   Automation folder).
2. **Get the code.** `git clone` this repo. Point the tenant's clasp target `scriptId` at the
   (new or existing) project. `npm run push:<tenant>`.
3. **Restore the manifest.** The push includes `appsscript.json` (advanced services + scopes).
   In the editor, confirm the advanced services are enabled and re-authorize when prompted.
4. **Restore secrets** (Script Properties): `SA_IMPERSONATION_EMAIL`, `SA_PRIVATE_KEY`
   (rotate — see security note), `MISSION_WEBHOOK_SECRET`, `MISSION_PARENT_FOLDER_ID`,
   and `MANAGE_RESOURCES=false` on cadets. Run `setAuthorization()` for the CAPWATCH token.
5. **Restore config.** Ensure the tenant's `config.gs` has that tenant's domain/ORGID/folder/sheet
   IDs. Verify the Drive folders and the automation spreadsheet (with its `Groups`,
   `User Additions`, `Error Emails`, `Mission Provisioning` tabs) still exist; recreate empty tabs
   if needed.
6. **Smoke test in order:** `testGetCapwatch()` → `previewLicenseLifecycle()` →
   `previewSquadronGroups()` → `getEmailGroupDeltas()`. All are read-only/preview. Only proceed to
   real runs once previews look right.
7. **Recreate triggers** (Editor → Triggers) per the [schedule](#8-what-runs-when-the-automation-schedule),
   owned by a durable account.
8. **Redeploy the web app** if this tenant hosts the mission webhook, and re-point FileMaker at the URL.
9. **Repeat per tenant.**

If a tenant's **service account** is lost: create a new dedicated service account in a Cloud
project, grant it domain-wide delegation scoped to **only** the Gmail-settings and Calendar scopes
the automation impersonates with, download a key, and put its email/key into **that tenant's**
Script Properties (`SA_IMPERSONATION_EMAIL` / `SA_PRIVATE_KEY`). Each tenant has its own SA — do
not reuse one across tenants.

---

## 16. Glossary

- **CAPWATCH** — CAP's national membership database; the source of truth. Data pulled as a ZIP of
  `.txt` (CSV) files from the eServices API.
- **eServices** — CAP's member portal; provides CAPWATCH API credentials.
- **ORGID** — CAPWATCH numeric organization id. PCR seniors config uses `188`.
- **OU (Organizational Unit)** — Workspace's account hierarchy (`/CA-001/…`); members are placed by OU.
- **Tenant** — one Workspace domain + its own Apps Script project (we have three).
- **clasp** — Google's CLI to push/pull Apps Script code from git.
- **Advanced Service** — a Google API (Admin SDK, Chat, Calendar, …) enabled in `appsscript.json`.
- **Script / User Property** — Apps Script key-value storage; Script = shared, User = per-user.
- **Trigger** — scheduled/event execution of a function; owned by the user who created it.
- **Service account** — a non-human Google identity; here, one dedicated per tenant, used to
  impersonate users for Gmail-settings and Calendar API calls that require domain-wide delegation.
- **Manual member** — a person added via the `User Additions` sheet (e.g. PCR/NHQ staff); never auto-suspended.

---

*Maintained for the CAP Pacific Region IT team. When any fact here changes — an ID, an owner, a
schedule, a secret location — update this file in the same PR. This document is only useful if it
stays true.*
