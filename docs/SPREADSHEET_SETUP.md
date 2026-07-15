# Spreadsheet Setup (for a new wing/region deployment)

The automation reads its per-tenant configuration and manual overrides from **Google
Sheets**, not from code. A new deployment must create these sheets and point the tenant's
Script Properties at them. This guide lists every tab the code touches, its exact columns,
and which ones are required.

Importable header-row templates live in [`../templates/`](../templates/) — one CSV per
input tab. See [Using the CSV templates](#using-the-csv-templates) at the bottom.

> Column sources are cited as `file.gs` so you can verify against the code if it changes.

---

## The two spreadsheets

| Spreadsheet | Script Property | Contents |
|---|---|---|
| **Automation spreadsheet** | `TENANT_AUTOMATION_SPREADSHEET_ID` | Group definitions, manual members, aliases, error log, etc. (most tabs below) |
| **Retention log** | `TENANT_RETENTION_LOG_SPREADSHEET_ID` | Auto-written log of retention emails sent (one `Log` tab; created automatically) |

Create both as normal Google Sheets, then record their IDs in the tenant's
[`config-tenants/<tenant>.json`](../config-tenants/) and set them as Script Properties (see
[config-tenants/README.md](../config-tenants/README.md)). The retention log only needs to
**exist** — the `Log` tab and its header are created on first send.

---

## Tabs at a glance

All of these live in the **Automation spreadsheet** unless noted. "Auto" = the automation
creates/manages it; you only need the empty tab to exist.

| Tab (canonical name) | Required? | Kind | Used by |
|---|---|---|---|
| `Groups` | **Yes** (for email groups) | input | `updateEmailGroups()`, ChatSpaces |
| `User Additions` | Recommended (may be empty) | input | `updateAdditionalGroupMembers()`, ChatSpaces |
| `Manual Members` | Optional | input | `updateAllMembers()` (members added by hand) |
| `External Contacts` | Region only (`RUN_SHARED_CONTACTS`) | input | `runExternalContactsToDomainSharedContacts()` |
| `Aliases` | Optional | input | send-as alias assignment in `updateAllMembers()` |
| `Secondary Aliases` | Optional | input + Auto | `addSecondaryDomainAliases()` (cols A–B yours, C–D written back) |
| `Error Emails` | **Yes** (must exist) | Auto | error reporting at the end of `updateEmailGroups()` |
| `Log` (in the retention sheet) | Auto | Auto | `logEmailSent()` in retention emails |

> ⚠️ **`Error Emails` must exist even though it's auto-managed.** The code calls
> `getSheetByName('Error Emails').getLastRow()` with no null-check, so a missing tab throws.
> Create it empty; the automation writes its own header row.

---

## Tab-naming consistency (read this)

The code currently accepts **several spellings** of the same tab (historical drift), e.g.
`User Additions` / `UserAdditions` / `USER ADDITIONS`, `Groups` / `GROUPS`, and
`Manual Members` / `ManualMembers`. **Use the canonical names in the "Tabs at a glance" table
above** — they all work today, but standardizing avoids a future cleanup breaking your tenant.
Header names are matched **case-insensitively**.

---

## Tab details

### `Groups` — group / committee / chat-space definitions
Drives which Google Groups (distribution lists) get created and who is in them. **Columns are
positional**, so keep this order (`UpdateGroups.gs`):

| Col | Column | Meaning | Example |
|--:|---|---|---|
| A | `Category` | Definition type | `duty-position` |
| B | `Group Name` | Base name of the DL (becomes `<wing>.<name>`) | `group-commanders` |
| C | `Attribute` | Rule selector | `dutyPositionIdsGroupScope` |
| D | `Values` | Comma-separated values for the rule | `CC` |
| E | `Description` | DL description (also the Chat-space display name) | `Group Commanders` |
| F | `Add Ext` | Allow external members? (`Y`/`Yes`/`X`/`True`) | `N` |

### `User Additions` — manual group membership
People to add to groups beyond what CAPWATCH produces. Header-aware with an A–D positional
fallback (`UpdateChatSpaces.gs`, `UpdateGroups.gs`):

| Col | Column | Meaning |
|--:|---|---|
| A | `Name` | Person's name (label only) |
| B | `Email` | Address to add |
| C | `Role` | `member` / `manager` / `owner` |
| D | `Groups` | Comma-separated group base names (must exist in `Groups`) |

### `Manual Members` — members not in CAPWATCH
Provision/keep accounts for people who aren't in the CAPWATCH pull (region staff attached from
other wings, etc.). **Header-keyed** (order doesn't matter); read by `getMembers()` in
`UpdateMembers.gs`:

| Column | Meaning |
|---|---|
| `CAPID` | Member CAPID (used as the account local-part) |
| `OrgID` | CAPWATCH ORGID for placement |
| `OrgName` | Optional label if the ORGID isn't in CAPWATCH |
| `LastName`, `FirstName`, `MiddleName` | Name |
| `Type` | Member type (defaults `SENIOR`) |
| `Status` | Membership status (defaults `ACTIVE`) |
| `DutyID` | Duty position id (optional) |
| `Email` | Member's personal/recovery email |
| `Secondary Email` | Workspace secondary email (optional) |
| `Primary Phone` | Recovery phone |

### `External Contacts` — Domain Shared Contacts (region feature)
Only used when `RUN_SHARED_CONTACTS` is on (the `pacific` profile). Synced into Google Domain
Shared Contacts by `SharedContacts.gs`. **Header-keyed**:

`CAPID, LastName, FirstName, MiddleName, Suffix, OrgID, Rank, DutyID, Assistant, Type, Status,
Email, Primary Phone`

(`Assistant` = `1` for an assistant duty; `Email` is the contact address matched on.)

### `Aliases` — Gmail send-as aliases
Optional per-account send-as configuration applied in `updateAllMembers()`
(`UpdateMembers.gs`). **Positional**, matched on the account's primary email in column A:

| Col | Column | Meaning |
|--:|---|---|
| A | `Primary Email` | The account to add the alias to |
| B | `Alias Email` | The send-as address |
| C | `Display Name` | Send-as display name |
| D | `Signature` | Optional HTML signature |

### `Secondary Aliases` — parallel addresses on a verified secondary domain
Drives `addSecondaryDomainAliases()` (`SecondaryDomainAliases.gs`). Each listed account
gets a **directory alias** that keeps the local part of its primary address but swaps in
`TENANT_SECONDARY_EMAIL_DOMAIN` — e.g. `jane.doe@cawgcap.org` also receives at
`jane.doe@cawg.cap.gov`. **Positional:**

| Col | Column | Meaning |
|--:|---|---|
| A | `Primary Email` | The existing account to alias |
| B | `Alias Email` | Optional override; leave blank to derive local part + secondary domain |
| C | `Status` | **Written by the script** — `ADDED`, `OK — already present`, `CONFLICT`, or `ERROR` |
| D | `Last Run` | **Written by the script** — timestamp of the last attempt |

This tab is an **opt-in list**, not the roster — only the accounts you list here get an
address on the secondary domain. Nothing is added automatically, so a new member who should
have one needs a row added by hand.

Prerequisites: the secondary domain must be **added and verified** in the tenant
(Admin console > Domains > Manage domains), and `TENANT_SECONDARY_EMAIL_DOMAIN` must be
set as a Script Property (leading `@` included). The script refuses to run otherwise.

Run `previewSecondaryDomainAliases()` first — it fills column C with the addresses it
*would* create without touching any account. The real run is idempotent, so it is safe to
re-run after fixing conflicts. Aliases become sendable in Gmail within roughly 24 hours.

**On a daily trigger** (`addSecondaryDomainAliases()`, see [ADMIN_GUIDE §8](ADMIN_GUIDE.md#8-what-runs-when-the-automation-schedule)):

- A run over a settled list makes **no changes** — rows already carrying their alias are
  reported `OK — already present` and skipped. Column D still updates, which is how you
  confirm the trigger is alive.
- **`CONFLICT` rows report once, then go quiet.** The address belongs to another account or
  group and only a human can free it, so re-reporting nightly would just train everyone to
  ignore the log. The row keeps its original `Status` and `Last Run` (the date it was first
  seen) and is skipped. **Editing column B un-latches it** — a different address means a
  different question, so the next run retries. Clearing column C also forces a retry.
- `ERROR` rows are **not** latched and retry every run, since they are often transient or
  simply pending (a row added before the account exists resolves itself the next morning).
- Never attach `previewSecondaryDomainAliases()` to a trigger — it writes to the same
  `Status` column the real run reads.

### `Error Emails` — auto-managed error log
Create the tab **empty**; the automation writes and maintains it. Header it sets:
`Email, CAPID, Error Count, Groups Affected, Error Codes, Last Error Message, Categories,
First Seen, Last Seen`.

### `Log` — auto-managed retention log
In the **retention** spreadsheet. Created automatically by `logEmailSent()` on the first
retention email; no setup needed beyond the spreadsheet existing.

---

## Minimum viable setup

For a standard wing (no region features):

1. Create the **Automation spreadsheet** with tabs: `Groups`, `User Additions`,
   `Manual Members` (optional), `Aliases` (optional), and an empty `Error Emails`.
2. Create the **Retention log** spreadsheet (empty).
3. Fill `Groups` with your distribution-list definitions (this is the main content).
4. Record both IDs in `config-tenants/<tenant>.json` → set as Script Properties →
   `validateTenantConfig()`.

A **region** deployment adds the `External Contacts` tab and turns on the region `RUN_*`
flags (see [PACIFIC_DIFF.md](PACIFIC_DIFF.md) and the `pacific` profile in `config.gs`).

---

## Using the CSV templates

Each input tab has a header-only CSV in [`../templates/`](../templates/):
`groups.csv`, `user-additions.csv`, `manual-members.csv`, `external-contacts.csv`,
`aliases.csv`.

To create a tab from a template:
1. In your Automation spreadsheet: **File → Import → Upload** the CSV.
2. Choose **Insert new sheet(s)**, then rename the inserted sheet to the canonical tab name
   above (the importer names it after the file).
3. Or simply copy the header row from the CSV into row 1 of a new tab.

The two **Auto** tabs (`Error Emails`, `Log`) have no template — create `Error Emails` empty;
`Log` is created for you.
