# Secondary Alias web app

A small, domain-restricted web app that lets an authorized person add a member to — or
remove them from — the `Secondary Aliases` tab **by CAPID**, without opening the
spreadsheet or holding Workspace admin rights.

Source: [`webapp/`](../webapp). Nightly counterpart:
[`src/accounts-and-groups/SecondaryDomainAliases.gs`](../src/accounts-and-groups/SecondaryDomainAliases.gs).

---

## Why it is a separate script project

The main automation project already owns `doGet()`/`doPost()` for the FileMaker
mission-provisioning webhook, deployed **`ANYONE_ANONYMOUS` + `USER_DEPLOYING`**. A script
project has exactly one `doGet`, and an admin UI behind an anonymous endpoint that runs as
the deployer would hand alias-management powers to anyone who learned the URL.

So `webapp/` is its own project, with its own manifest (`access: DOMAIN`), its own clasp
target, and a much smaller scope list. The FileMaker webhook is untouched.

The cost is that a handful of pure helpers are **duplicated** from `src/` — address
derivation and CAPID resolution. `test/AliasWebApp.test.js` runs both copies over the same
inputs and fails if they disagree, because a divergence would show up in production only as
a sheet row whose status flips every night.

## How access works

| | |
|---|---|
| **Execute as** | the deploying account (you) |
| **Who has access** | anyone in the domain |
| **Who can actually use it** | members of the group in `WEBAPP_ALIAS_ADMIN_GROUP` |

`executeAs: USER_DEPLOYING` is what lets a unit IT officer who is *not* a Workspace admin
create an alias — the Directory calls run with your rights, not theirs. That also means the
deployment grants nothing on its own and [`webapp/Auth.gs`](../webapp/Auth.gs) is the entire
access control. It fails **closed**: no group configured, no identity, or a Directory API
error all deny.

`access: DOMAIN` is load-bearing, not a preference — under `ANYONE_ANONYMOUS` there would be
no identity to check at all.

**To add or remove an authorized user, edit the group in the Admin console.** No redeploy,
no code change, and membership is auditable there. Nested groups work.

## Setup

Once per tenant.

1. **Create the group** that grants access, e.g. `ca.it-alias-admins@cawgcap.org`, and add
   yourself. Restrict who can join to invited members only.

2. **Create the script project** and record its ID:
   ```bash
   npx clasp create-script --title "CAWG Secondary Alias Admin" --type standalone \
     --rootDir webapp
   ```
   Put the resulting script ID into [`clasp-targets/alias-webapp.clasp.json`](../clasp-targets/alias-webapp.clasp.json)
   (the `rootDir` there is already correct; do not commit a `.clasp.json` clasp wrote into
   `webapp/`).

3. **Push:**
   ```bash
   npm run push:webapp
   ```

4. **Set Script Properties** on the new project (Project Settings › Script Properties).
   The `TENANT_*` names are identical to the main project's, so copy the values from
   `config-tenants/<tenant>.json`:

   | Property | Value |
   |---|---|
   | `TENANT_EMAIL_DOMAIN` | e.g. `@cawgcap.org` |
   | `TENANT_SECONDARY_EMAIL_DOMAIN` | e.g. `@cawg.cap.gov` |
   | `TENANT_AUTOMATION_SPREADSHEET_ID` | the automation spreadsheet |
   | `TENANT_WING_ABBREVIATION` | optional; page heading only |
   | `WEBAPP_ALIAS_ADMIN_GROUP` | the group from step 1 — **the app denies everyone until this is set** |

5. **Deploy:** Deploy › New deployment › Web app. Execute as **Me**, Who has access
   **Anyone in <domain>**. Authorize the scopes when prompted, then share the URL with the
   group.

> The deploying account must be able to create aliases and read group membership — in
> practice a super admin, which is what the existing automation account already is.

## Using it

- **Add:** type a CAPID, press *Look up*. The app shows every account in the tenant carrying
  that CAPID, ranked with the most recently signed-into first, and shows the address each
  would receive. Press *Add and create alias*. The row goes on the tab **and** the alias is
  created immediately — you do not wait for the nightly run.
- **Duplicates:** when a CAPID has more than one account, all of them are listed with a
  warning. This is deliberate. On these tenants the tidy-looking address is frequently the
  dead twin, so the app will not choose for you.
- **Remove:** press *Remove* on a listed row and confirm. This takes the row off the tab
  **and revokes the alias**, so the address stops delivering immediately. Removing the row
  alone would leave a live address behind forever — the nightly module only ever adds.
- **Audit:** every add, removal and failed removal is appended to the `Alias Admin Log` tab
  with timestamp, actor, CAPID, address and result.

### What it refuses to do

[`webappRemoveAliasFromAccount_`](../webapp/Directory.gs) will not remove an address that is
not on the configured secondary domain, is the account's primary address, or is not actually
an alias on that account. A row whose alias was already removed by hand still comes off the
tab, reported as `NOT PRESENT`.

## The sheet contract

`addSecondaryDomainAliases()` reads the whole tab, builds a write-back array, and commits
columns C–D **positionally** in one call. Row position is therefore a shared, unlocked data
structure between two script projects — and `LockService` locks are per-project, so the two
**cannot** interlock.

The web app therefore never shifts a row:

- **Adding** appends, or reuses a blank row. That only makes the tab longer than the array,
  so the extra row is left untouched by a run already in flight.
- **Removing** clears the row in place and leaves it blank. The nightly module skips a blank
  column A and seeds its write-back with each row's current value, so a blank row stays blank.

Columns **E** (`CAPID`) and **F** (`Added By`) belong to the web app. They sit outside the
nightly module's write range, which is why nothing about provisioning may ever depend on
them — they are provenance, not input.

## Tenant scope

The code is tenant-generic and self-disables where `TENANT_SECONDARY_EMAIL_DOMAIN` is blank,
matching how the rest of the codebase handles per-tenant behavior. In practice only
**seniors** has a secondary domain today, so that is the only tenant worth deploying to.

## Troubleshooting

| Symptom | Cause |
|---|---|
| "You are not authorized" for everyone, including you | `WEBAPP_ALIAS_ADMIN_GROUP` unset or wrong. Check the execution log — the group address is named there. |
| Banner: domain not verified | The secondary domain is not verified in this tenant. Aliases cannot be created until it is; the nightly module has the same constraint. |
| Add returns `CONFLICT` | Another account or group already holds that exact address. Run `findSecondaryDomainAliasBlockers()` in the **main** project — it pairs each blocking account with the member it is denying. |
| A row's status never changes after clearing a blocker | The nightly module latches conflicts and only un-latches when the row's alias address *changes*. Nudge column B on that row to force a retry. |
| Blank rows in the middle of the tab | Expected — see [the sheet contract](#the-sheet-contract). |
