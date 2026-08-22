# New Tenant / New Wing Setup (bare-metal runbook)

How to stand up the automation on a **brand-new Workspace tenant from nothing** — a new
wing adopting this codebase, or an additional tenant for an existing wing.

The [Admin Guide](ADMIN_GUIDE.md) is the *operational* runbook (for a system already
running); this is the *provisioning* runbook that gets you to that starting line. It does
not duplicate the detailed docs — it puts them in the right **order** and flags the
ordering traps. Each step links to the authoritative doc for specifics.

> **Worked example — Hawaii Wing (HIWG).** HIWG runs the same split structure as
> California: **two tenants**, seniors + cadets. The example callouts throughout show the
> HIWG-specific choices. Its templates already exist:
> [`config-tenants/setup-hiwg.gs`](../config-tenants/setup-hiwg.gs),
> [`hiwg-seniors.json`](../config-tenants/hiwg-seniors.json),
> [`hiwg-cadets.json`](../config-tenants/hiwg-cadets.json).

## Mental model (read once)

- There is **one** `src/` codebase, deployed **unchanged** to every tenant via a clasp
  target in [`clasp-targets/`](../clasp-targets). Nothing wing-specific lives in source.
- A tenant's entire identity and behavior come from its **Script Properties**, which a
  push never touches. Wing display labels (`HIWG`, `Hawaii Wing`) are **derived** from
  `TENANT_WING` — no code edits to adopt a new wing. See
  [config-tenants/README.md](../config-tenants/README.md).
- **`config.gs` is overwritten on every push.** Per-tenant values must be Script
  Properties, never literals in source.

## Prerequisites

- **Super Admin** on each new Workspace tenant.
- **CAP eServices / CAPWATCH** credentials authorized for the wing's ORGID (needs a
  security assessment + commander approval — start this early; it gates everything).
- **Node + [clasp](https://github.com/google/clasp)**, with the Apps Script API enabled
  for your Google account at [script.google.com/home/usersettings](https://script.google.com/home/usersettings).
- Rights to **create a Google Cloud project** in each tenant's Workspace org.

---

## Phase 1 — Decide the tenant shape

Per tenant, pick the **profile** (`TENANT_PROFILE`) — it selects member types, cadet-lite
mode, the squadron-group set, region flags, and cross-tenant behavior:

| Deployment | Tenants | Profiles |
|---|---|---|
| Split wing (CAWG, **HIWG**) | seniors + cadets | `seniors`, `cadets` |
| Single composite wing | one | `seniors` |
| Region HQ | one | `region` (see [REGION_DIFF.md](REGION_DIFF.md)) |

> **HIWG:** two tenants → `seniors` and `cadets`, mirroring CAWG.

## Phase 2 — Workspace prep (per tenant)

1. Confirm each tenant's **primary domain** is verified and you hold **Super Admin**.
2. Note the wing's **CAPWATCH ORGID** and two-letter **wing code** (`HI`).
3. (Optional) a **verified secondary domain** if the tenant will hand out parallel
   addresses (`TENANT_SECONDARY_EMAIL_DOMAIN`; seniors-style only).
4. On a **cadet** tenant, note the **senior** tenant's email domain — it goes in
   `TENANT_COMMAND_EMAIL_DOMAIN` (Phase 6). Unit commanders, personnel officers and
   deputy commanders are senior members, so `notifications/RecoveryEmailNotify.gs`
   addresses them on the senior domain; left blank it would derive addresses on the
   cadet domain, where those people have no account.

## Phase 3 — Standard GCP project + APIs + OAuth (per tenant)

The automation needs a **standard (self-owned) GCP project**, not the Apps Script default
project, because the Contacts API (shared contacts) can only be enabled on a standard
project. Do this **before** attaching it so the script is only unauthorized for a minute.

Follow [GCP_PROJECT_MIGRATION.md](GCP_PROJECT_MIGRATION.md) steps 1–7 — but for a new build
you *create* the project fresh rather than migrating an existing one:

1. Create the project **in that tenant's Workspace org** (this is what makes the
   **Internal** OAuth consent screen available, skipping verification).
2. Enable all **7 APIs**: Admin SDK, Drive, Gmail, Calendar, Chat, Groups Settings,
   **Contacts**. A missing API fails a time-based trigger *silently* — enable them all up
   front.
3. OAuth consent screen → **User type: Internal**.
4. You'll attach this project to the Apps Script project in Phase 5.

> **HIWG:** one GCP project per tenant — seniors project under the HIWG senior org, cadets
> project under the HIWG cadet org.

## Phase 4 — Service account + domain-wide delegation (per tenant)

Each tenant gets its **own dedicated service account** for the calls that need per-user
impersonation. It is **never** in source — its email/key go into Script Properties in
Phase 6. See [Admin Guide §3](ADMIN_GUIDE.md#3-system-inventory-the-facts-you-cannot-guess).

1. Create a service account, download a JSON key.
2. In the tenant's Admin console → **Security → API Controls → Domain-wide Delegation**,
   authorize the SA's **client ID** with **only** the scopes that tenant actually
   impersonates (least privilege — verify the live set against the code, don't over-grant):

   | Tenant role | DWD scopes to grant |
   |---|---|
   | Any tenant (baseline) | `gmail.settings.basic`, `gmail.settings.sharing`, `calendar` |
   | **Cadets** (transition *source*) | above **+** mail migration: `gmail.insert`, `gmail.labels`, `gmail.metadata`, `gmail.readonly`, `drive`, `contacts` |
   | **Cross-tenant peer SA** (read-only, in the *peer* tenant) | `admin.directory.user.readonly`, `admin.directory.group.readonly` |

> **HIWG:** the cadets SA carries the transition-migration scopes; each tenant also needs a
> read-only **peer** SA in the other tenant if you run cross-tenant contacts (Phase 10).

## Phase 5 — Add a clasp target and first push

1. Copy an existing target, e.g. `clasp-targets/hiwg-seniors.clasp.json`, set its
   `scriptId` to the new Apps Script project (`rootDir` stays `../src`). Add matching
   `npm run status/push/pull/open:hiwg-seniors` scripts to `package.json`.
2. Attach the GCP project from Phase 3: Apps Script editor → **Project Settings → GCP
   Project → Change** → the project number → re-run any function once to re-consent
   ([GCP_PROJECT_MIGRATION.md](GCP_PROJECT_MIGRATION.md) steps 4–5).
3. **Set Script Properties before you push** (Phase 6) — an unconfigured project fails
   loudly by design. Then `npm run push:hiwg-seniors`.

The manifest ([`src/appsscript.json`](../src/appsscript.json)) carries the advanced
services + OAuth scopes and deploys with the push; you re-consent once per project.

## Phase 6 — Script Properties (identity + secrets)

1. **Create the Drive assets first** so you have their IDs: the CAPWATCH-data folder, the
   automation folder, and the automation spreadsheet (Phase 7). This avoids a second pass.
2. Paste the tenant's `setup-hiwg.gs` into the editor, fill the `FILL_IN` values, and run
   `setupHiwgSeniorsScriptProperties()` / `setupHiwgCadetsScriptProperties()`. (Wing labels
   `HIWG` / `Hawaii Wing` derive automatically from `TENANT_WING=HI`.)
3. Add **secrets** by hand in Project Settings → Script Properties: `SA_IMPERSONATION_EMAIL`,
   `SA_PRIVATE_KEY` (from Phase 4), and `XT_PEER_SA_KEY` if using cross-tenant.
4. Run **`validateTenantConfig()`** — it lists any missing required key.
5. On a **cadet** tenant, set `TENANT_COMMAND_EMAIL_DOMAIN` to the senior domain
   (e.g. `@cawgcap.org`). It is optional and *not* flagged by `validateTenantConfig()` —
   blank silently falls back to this tenant's own `TENANT_EMAIL_DOMAIN`, which sends the
   monthly account-compliance digests to cadet-domain addresses that do not exist. See
   [`config-tenants/cadets.json`](../config-tenants/cadets.json).

Full key inventory: [Admin Guide §7](ADMIN_GUIDE.md#7-secrets-and-script-properties-the-part-that-breaks-silently)
and [config-tenants/README.md](../config-tenants/README.md).

## Phase 7 — Automation spreadsheet

Build the automation config spreadsheet per
[SPREADSHEET_SETUP.md](SPREADSHEET_SETUP.md) (§"Minimum viable setup" + CSV templates), and
put its ID into `TENANT_AUTOMATION_SPREADSHEET_ID` (Phase 6). At minimum: `Groups`,
`User Additions`, `Manual Members`, plus the auto-managed `Error Emails` / `Log` tabs.

## Phase 8 — CAPWATCH credential (per downloading tenant)

Run `setAuthorization()` once (temporarily inline the eServices creds, run, then delete
them from source) — it stores a per-user `CAPWATCH_AUTHORIZATION` **User Property**. Verify
with `testGetCapwatch()`. See [Admin Guide §7](ADMIN_GUIDE.md#7-secrets-and-script-properties-the-part-that-breaks-silently).

> ⚠️ It's a **User** Property — only the user who set it can download. Set it as the account
> that will own the `getCapwatch` trigger.

## Phase 9 — Dry-run before arming anything

Confirm the plumbing on real data with the read-only/preview paths before any trigger
mutates the directory:

1. `getCapwatch()` → data lands in Drive.
2. A member-sync **preview** and a groups **preview** — check **Executions** for
   `PERMISSION_DENIED` / `API disabled` (missing DWD scope or un-enabled API).
3. `validateTenantConfig()` clean.

## Phase 10 — Cross-tenant contacts (optional; split tenants)

If seniors ⇄ cadets should see each other in the GAL, configure the peer SA and run
`setupCrossTenantConfig()` + `validateCrossTenantConfig()` per
[CROSS_TENANT_CONTACTS.md](CROSS_TENANT_CONTACTS.md) and
[config-tenants/README.md](../config-tenants/README.md#cross-tenant-contacts-xt_peer_).

## Phase 10b — Member self-service signature app (optional)

A separate Apps Script project (`signature-webapp/`) letting members set their own CAP signature.
Independent of everything above and safe to add later — but note it needs **its own copy** of
`SA_IMPERSONATION_EMAIL` / `SA_PRIVATE_KEY`, so add it to the key-rotation checklist when you
deploy it. Setup: [SIGNATURE_WEB_APP.md](SIGNATURE_WEB_APP.md).

## Phase 10c — Domain admin help desk app (optional)

A separate Apps Script project (`admin-webapp/`) for wing IT staff holding Google's Help Desk
Administrator role: password resets, welcome-email resends, 2SV setup group. One project per
tenant — add a clasp target per tenant (e.g. `admin-webapp-<tenant>.clasp.json`), run its
`admin-webapp/Setup.gs` values (sourced from `config-tenants/<tenant>.json`), and deploy by
hand from the editor, never from clasp. Setup: [ADMIN_WEB_APP.md](ADMIN_WEB_APP.md).

## Phase 10d — Secondary alias web app (optional; seniors only)

A separate Apps Script project (`secondary-alias-webapp/`) letting a CAPID-scoped caller add or
remove a member's secondary-domain alias without spreadsheet or admin access; on Add it
configures Gmail Send-As and emails the member immediately. Only meaningful on a tenant with
`TENANT_SECONDARY_EMAIL_DOMAIN` set. Setup: [secondary-alias-webapp/README.md](../secondary-alias-webapp/README.md).

## Phase 11 — Arm triggers

Install the time-driven triggers in dependency order (CAPWATCH download first, everything
else after). The full schedule is in [README "What runs, when"](../README.md#what-runs-when)
and [Admin Guide §8](ADMIN_GUIDE.md#8-what-runs-when-the-automation-schedule). Push one
tenant, confirm it's healthy, then the next.

The **cadets** tenant also arms the cadet→senior transition pipeline
(`armTransitionPipelineTrigger()`, one daily trigger running all phases in order). This must
run **as the automation account**, and the final `closeCompletedTransitions()` stays
**manual** — never triggered (it only parks the account; `expireParkedAccounts()` is what
later deletes it, automatically, a year after parking).

---

## What the new wing must provide

These are exactly the `FILL_IN` fields in the `setup-hiwg.gs` / `hiwg-*.json` templates:

| Value | Property | Notes |
|---|---|---|
| Senior + cadet Workspace domains | `TENANT_DOMAIN`, `TENANT_EMAIL_DOMAIN` | one per tenant |
| Peer cadet domain | `TENANT_CADETS_TENANT_DOMAIN` | blank derives `<wing>wgcadets.org` |
| CAPWATCH ORGID | `TENANT_CAPWATCH_ORGID` | the wing's ORGID |
| Drive folder / spreadsheet IDs | `TENANT_*_FOLDER_ID`, `TENANT_AUTOMATION_SPREADSHEET_ID` | created in Phases 6–7 |
| Contact / sender addresses | `TENANT_RETENTION_EMAIL`, `TENANT_DIRECTOR_RECRUITING_EMAIL`, `TENANT_DIRECTOR_RECRUITING_NAME`, `TENANT_AUTOMATION_SENDER_EMAIL`, `TENANT_TEST_EMAIL`, `TENANT_ITSUPPORT_EMAIL` | `TENANT_DIRECTOR_RECRUITING_NAME` signs member-facing retention email; blank is valid |
| Service accounts + keys | `SA_IMPERSONATION_EMAIL`, `SA_PRIVATE_KEY`, `XT_PEER_SA_*` | Phase 4; secrets |
| CAPWATCH credential | `CAPWATCH_AUTHORIZATION` (User Property) | Phase 8 |

Wing labels (`HIWG`, `Hawaii Wing`) and the standard cadet-domain pattern are **derived**,
so they are intentionally **not** on this list.

## Cross-reference index

| For… | See |
|---|---|
| Script Properties + push order | [config-tenants/README.md](../config-tenants/README.md), [Admin Guide §7](ADMIN_GUIDE.md#7-secrets-and-script-properties-the-part-that-breaks-silently) |
| GCP project, APIs, OAuth, scopes | [GCP_PROJECT_MIGRATION.md](GCP_PROJECT_MIGRATION.md) |
| Service accounts + DWD | [Admin Guide §3](ADMIN_GUIDE.md#3-system-inventory-the-facts-you-cannot-guess) |
| Automation spreadsheet | [SPREADSHEET_SETUP.md](SPREADSHEET_SETUP.md) |
| Trigger schedule | [README](../README.md#what-runs-when), [Admin Guide §8](ADMIN_GUIDE.md#8-what-runs-when-the-automation-schedule) |
| Cross-tenant contacts | [CROSS_TENANT_CONTACTS.md](CROSS_TENANT_CONTACTS.md) |
| Self-service signature web app | [SIGNATURE_WEB_APP.md](SIGNATURE_WEB_APP.md) |
| Region-tenant differences | [REGION_DIFF.md](REGION_DIFF.md) |
| Local clasp workflow | [DEVELOPMENT.md](DEVELOPMENT.md), [Admin Guide §6](ADMIN_GUIDE.md#6-local-development-with-clasp) |
