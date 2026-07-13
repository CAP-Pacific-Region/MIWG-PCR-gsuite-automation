# Per-tenant configuration (`config-tenants/`)

`src/config.gs` is **shared by all three Apps Script projects** and is **overwritten on every
`clasp push`** (all targets use `rootDir: ../src`, and `.claspignore` ships everything under
`src/`). Tenant-specific values therefore **cannot** live as literals in `config.gs` — a push
would clobber them. This is what wiped the cadets project's config once already.

Instead, each project's identity lives in that project's **Script Properties**, which `clasp push`
never touches. `getTenantConfig_()` in `src/config.gs` reads them at runtime into the `TENANT`
object, which `CONFIG` and the recruiting/retention constants consume.

## What's in this folder

One JSON file per tenant holding the **canonical, non-secret** values for that project's
`TENANT_*` Script Properties. These files are **repo-only** — `config-tenants/` is outside `src/`,
so clasp never uploads them. They exist for version control and disaster recovery.

- `seniors.json` — cawgcap.org (complete)
- `cadets.json` — cawgcadets.org (complete)
- `pacific.json` — Pacific Region (complete; verified from the live "PCR Automation" `config.gs`)
- `setup-pacific.gs` — paste-once helper (`setupPacificScriptProperties()`) that writes
  Pacific's `TENANT_*` values from `pacific.json` to the project's Script Properties. Repo-only,
  never shipped by clasp. See `docs/PACIFIC_DIFF.md`.

**Secrets are NOT stored here** — `SA_PRIVATE_KEY`, `SA_IMPERSONATION_EMAIL`, `SA_PRIVATE_KEY_ID`,
`MISSION_WEBHOOK_SECRET`, `MISSION_PARENT_FOLDER_ID`, the CAPWATCH `CAPWATCH_AUTHORIZATION`
token, and the cross-tenant peer service-account key `XT_PEER_SA_KEY` live only in each project's
Script/User Properties. See `docs/ADMIN_GUIDE.md`.

## Cross-tenant contacts (`XT_PEER_*`)

`src/cross-tenant-contacts/CrossTenantContacts.gs` publishes the **peer** Workspace tenant's
members into this tenant's Global Address List (seniors ⇄ cadets). It reads live data only — one
wing CAPWATCH pull (roster, incl. cadet-lite members) plus the peer tenant's directory — so **no
export spreadsheet is involved**. Behavior is profile-driven (`PROFILE_.CROSS_TENANT` in
`config.gs`, on for seniors/cadets, off for pacific); identity + credentials are Script Properties:

| Key | Secret? | Set via | Seniors project | Cadets project |
| --- | --- | --- | --- | --- |
| `XT_PEER_DOMAIN` | no | `setupCrossTenantConfig()` / JSON | `cawgcadets.org` | `cawgcap.org` |
| `XT_PEER_SA_EMAIL` | no | `setupCrossTenantConfig()` | cadets read-only SA `client_email` | seniors read-only SA `client_email` |
| `XT_PEER_SA_SUBJECT` | no | `setupCrossTenantConfig()` | a **cadets** super admin to impersonate | a **seniors** super admin to impersonate |
| `XT_PEER_SA_KEY` | **YES** | by hand in Script Properties | cadets SA private key (PEM) | seniors SA private key (PEM) |

The peer service account is a **dedicated, read-only** SA in the *peer* tenant with domain-wide
delegation limited to `https://www.googleapis.com/auth/admin.directory.user.readonly` (member sync)
and `https://www.googleapis.com/auth/admin.directory.group.readonly` (parent-group sync, if
`RUN_PARENTS`). It only ever reads the peer directory; writes to this tenant's shared contacts use
the script's own token.

**Setup per wing project:**
1. In the peer tenant, create a service account + grant DWD for `admin.directory.user.readonly`
   (add `admin.directory.group.readonly` too if the parent-group sync is enabled).
2. On this project, run `setupCrossTenantConfig()` (domain, SA email, subject), then add
   `XT_PEER_SA_KEY` by hand in **Project Settings → Script Properties**.
3. Run `validateCrossTenantConfig()` to confirm, then add a trigger for `syncCrossTenantContacts`.

This also requires the legacy `https://www.google.com/m8/feeds` scope (already in
`src/appsscript.json`) — each project re-consents once after the shared `src/` is pushed. Pacific
needs none of the above (`RUN_INBOUND=false`).

## Applying values to a project

1. Open the target project (`npm run open:<tenant>`).
2. Either enter the `TENANT_*` keys by hand in **Project Settings → Script Properties**, or paste
   the values into `setupTenantConfig()` in `config.gs` and **Run** it once. For **Pacific**, the
   values are already filled in — paste `setup-pacific.gs` and run `setupPacificScriptProperties()`.
3. Run `validateTenantConfig()` — it lists any missing required keys.

## ⚠️ Migration / push order (important)

Because `getTenantConfig_()` has **no fallback** (an unconfigured project yields empty identity
fields and fails loudly rather than acting on the wrong domain), you must **set a project's
`TENANT_*` properties BEFORE pushing the shared `config.gs` to it.** For the already-live tenants,
set the properties first (values above), confirm with `validateTenantConfig()`, then push.
