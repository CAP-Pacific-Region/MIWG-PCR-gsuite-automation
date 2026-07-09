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
- `cadets.json` — cawgcadets.org (fill the `TODO` fields from the cadet-tenant restore)
- `pacific.json` — Pacific Region (unverified; fill once `it@pcr.cap.gov` access is restored)

**Secrets are NOT stored here** — `SA_PRIVATE_KEY`, `SA_IMPERSONATION_EMAIL`, `SA_PRIVATE_KEY_ID`,
`MISSION_WEBHOOK_SECRET`, `MISSION_PARENT_FOLDER_ID`, and the CAPWATCH `CAPWATCH_AUTHORIZATION`
token live only in each project's Script/User Properties. See `docs/ADMIN_GUIDE.md`.

## Applying values to a project

1. Open the target project (`npm run open:<tenant>`).
2. Either enter the `TENANT_*` keys by hand in **Project Settings → Script Properties**, or paste
   the values into `setupTenantConfig()` in `config.gs` and **Run** it once.
3. Run `validateTenantConfig()` — it lists any missing required keys.

## ⚠️ Migration / push order (important)

Because `getTenantConfig_()` has **no fallback** (an unconfigured project yields empty identity
fields and fails loudly rather than acting on the wrong domain), you must **set a project's
`TENANT_*` properties BEFORE pushing the shared `config.gs` to it.** For the already-live tenants,
set the properties first (values above), confirm with `validateTenantConfig()`, then push.
