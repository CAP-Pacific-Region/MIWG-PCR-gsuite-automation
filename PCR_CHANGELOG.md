# PCR Changelog

Pacific Region (PCR) fork-specific changes to the CAPWATCH / Google Workspace
automation, layered on top of the upstream `cap-miwg/gsuite-automation` project.
Upstream changes are tracked in [CHANGELOG.md](CHANGELOG.md); **this file records
only what the Pacific Region deployment adds or diverges on.**

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).
Individual source files carry their own SemVer version in their header
(see [docs/VERSIONING.md](docs/VERSIONING.md)); the per-file version is noted
next to each entry below.

## [2026-07-18] — Cross-tenant contacts: sort by last name, and carry cadet-lite into the cadet GAL

Two directory issues, both visible in the shared GAL that publishes each tenant's members
into the other.

### Fixed (`CrossTenantContacts.gs` 0.2.0)

- **Shared contacts sorted by first name.** Cadets displayed "Last, First M Grade" but sorted
  by first name, while native directory users sorted by last name — same list, two orders.
  Domain Shared Contacts sort in the GAL by the structured `gd:givenName` (not by `fullName`,
  and not by `familyName` the way directory *users* do). Google exposes no per-contact sort
  override, so `xtBuildContactXml_` now writes the whole **"Last Suffix, First M Grade"**
  display string into `givenName` (and omits a separate `familyName`, so the contact card's
  First/Last split isn't a doubled name). Display is unchanged; the sort key now leads with the
  last name. The display also gains the middle initial + suffix, matching native accounts
  (`xtDisplayName_`).

### Added (`CrossTenantContacts.gs` 0.2.0, `config.gs` 1.8.0, `UpdateMembers.gs` 1.16.0)

- **Self no-account publish.** The cadets tenant now publishes its own **cadet-lite** members
  (grades below C/SSgt, which get no account and so were absent from the cadet GAL) as shared
  contacts off their CAPWATCH personal email — the same way the seniors tenant already carries
  them cross-tenant. Driven by the new `PROFILE_.CROSS_TENANT.SELF_NO_ACCOUNT_TYPES`
  (`['CADET']` on cadets, `[]` elsewhere). Folded into `syncCrossTenantContacts` under the
  existing marker, so **no new trigger**. Adds `xtSelfWorkspaceEmailByCapid_()` (self-directory
  read to skip accountholders) and a `getMembers(types, incDuty, /*includeCadetLite=*/true)`
  bypass of the cadet-lite grade filter. No new OAuth scopes.

### Notes

- Both name changes alter every managed contact's content hash, so the first run after deploy
  updates every existing contact (~2,630 seniors, ~1,714 cadets) plus the new cadet self-lite
  creates. That is within the m8 ~3,000-writes/tenant/day cap but tight on seniors; the sync
  resumes across runs, so expect 1–2 days to converge.

### Fixed follow-ups (`CrossTenantContacts.gs` 0.2.1) — after first live run

- **Sort fix only reached ~2/3 of contacts.** `xtHash_` keyed only on the display fields, so the
  0.2.0 change to the card *shape* (sort value into `givenName`) left the hash unchanged for any
  contact whose display *text* didn't also change (no middle initial / suffix) — the stateless
  reconcile saw "unchanged" and skipped the rewrite, so those kept `givenName=First` and still
  sorted by first name. Verified live on seniors: 1,830 rewritten, 822 left old-format. Added
  `XT_CARD_FORMAT` to the hash so a shape change forces a one-time rewrite of the whole set.
- **Self-publish shadowed real accounts.** The self no-account publish created a duplicate contact
  over an account-holder's own account (observed: a C/2dLt cadet commander appearing twice in the
  cadet directory). Two causes: the account-existence skip keyed only on the org `externalId`, so
  an account carrying its CAPID only in `employeeId` slipped through; and a tenant-domain CAPWATCH
  email was published as if personal. Now the self account map indexes CAPID from `externalId`
  **and** `employeeId`, and a domain guard drops any self-contact whose email is on one of this
  tenant's own domains (`cfg.ownDomains`). Existing shadow contacts self-heal — no longer desired,
  so the reconcile deletes them on the next run.

---

## [2026-07-18] — Recovery email: never blank it, and source it from any personal address

Follow-up to the recovery-phone work below, after a `forceUpdateAllMembers()` run
**removed a live recovery email** (CAPID 123541: `pappasmurf2016@aol.com` → `""`).

### Fixed (`UpdateMembers.gs` 1.14.1)

- **Never blank an existing recoveryEmail/recoveryPhone.** The payload built
  `recoveryEmail: member.secondaryEmail || member.parentEmail || ''`, so a member with
  no usable personal email had their good recovery address **overwritten with `''`** on
  every full re-write — defeating password reset. Latent since the recovery-contact
  feature, but masked by `memberUpdated()` change-gating; the 1.14.0 full re-write
  surfaced it at scale. Recovery fields are now sent **only when non-empty**; omitting a
  field makes `Users.update` preserve the existing value.

### Changed (`UpdateMembers.gs` 1.15.0)

- **Recovery / second-contact email now sources from PRIMARY or SECONDARY**, preferring
  a personal address. `firstPersonalEmail_()` skips any CAPWATCH address on the tenant's
  own domains (`TENANT_DOMAIN` / `TENANT_SECONDARY_EMAIL_DOMAIN`) and takes the next
  candidate — SECONDARY, then PRIMARY, then cadet parent (recovery only). This covers the
  many members who list a personal email as PRIMARY despite the wing recommending
  SECONDARY, who previously got no recovery email at all. `recoveryEmail` is now a
  derived `member.recoveryEmail` field and participates in change detection.

### Notes

- Directory fields (`phones` / `emails`) keep full-replace semantics, so intentional
  cadet directory-phone removal is unaffected.
- Recovery emails already blanked by the force run are being restored separately from the
  "Workspace user update diff" Cloud Logging records, which captured each prior value.

## [2026-07-17] — Recovery phone ignores DoNotContact; cadet phones kept out of the directory

Members whose CAPWATCH cell-phone row is flagged **DoNotContact** were getting no
recovery phone at all, so they couldn't self-serve a password reset. The DoNotContact
flag now applies only to what is *published* — recovery contact info is exempt.

### Changed (`UpdateMembers.gs` 1.14.0)

- **Recovery phone now ignores DoNotContact.** `addContactInfo()` tracks a new
  `member.recoveryPhone` (member cell phone, then cadet parent phone) that is populated
  regardless of the DoNotContact flag, mirroring how `recoveryEmail` already used the
  secondary email. `addOrUpdateUser()` writes it to `recoveryPhone` on create and update.
- **Directory phone (`member.phone`) is unchanged for seniors** — still excludes
  DoNotContact rows — but is **never populated for cadets**. A cadet's number must not
  appear in the global directory; passing `phones: []` on update also **removes any cadet
  number Google had already published**.
- `memberUpdated()` now compares `recoveryPhone`, so a recovery-phone-only change (e.g.
  backfilling a previously-DoNotContact number) triggers a sync.

### Notes

- **No change to email behavior:** `recoveryEmail` already ignored DoNotContact via the
  secondary email, and the directory "other" email already excluded DoNotContact rows.
- Recovery email and recovery phone are correct on **all** tenants (senior and cadet).

## [2026-07-17] — Legacy 'DL-CAWG-*' migration groups: read-only inventory tooling

Groundwork for clearing verbose `DL-CAWG-…` distribution lists left over from the
M365 → Google migration (superseded by the modern `ca###.all` convention). Users
still see them in Gmail autocomplete.

### Added

- **`groupAdministration.gs`** — two **read-only** helpers:
  - `groupAdministration_stageLegacyDlGroups(prefix='dl-cawg', sheetName='Legacy DL Cleanup')`
    inventories live Groups whose **primary address** or an **alias** starts with the
    prefix, to a review tab. Rows are tagged **PRIMARY** (safe to delete the group, via
    `bulkDeleteGroupsFromSheet`, once confirmed an unused duplicate) vs **ALIAS** (remove
    only the alias with `AdminDirectory.Groups.Aliases.remove`, never the group) so the
    two are not conflated.
  - `groupAdministration_resolveLegacyAddress(email)` resolves one address definitively
    via `Groups.get` — group primary / group alias / not-a-group.

### Notes

- **Scope boundary:** these clear neither directory objects (that's a reviewed second
  step) nor per-user Gmail autocomplete. Autocomplete entries live in each user's
  **"Other contacts"** (auto-saved recent recipients) and are **not centrally
  removable** by Apps Script or GAM; deleting a live group/alias only stops it
  re-seeding autocomplete and de-clutters the GAL. Per-user removal is the ✕ on the
  Gmail suggestion (or contacts.google.com → Other contacts).

## [2026-07-17] — Genericize wing labels so the code can deploy to another wing (e.g. Hawaii)

Removes the last hard-coded `CAWG` / `California Wing` literals so a second wing can
adopt the automation by Script Property alone. Prompted by Hawaii Wing (HIWG), which
runs the same split senior/cadet tenant structure.

### Added

- **`config.gs` (v1.7.0)** — programmable wing labels derived from `TENANT_WING`:
  `WING_ABBREVIATION_` (`CA` → `CAWG`, `HI` → `HIWG`), `WING_NAME_` (proper name via a
  `WING_NAMES_` map, e.g. `California Wing` / `Hawaii Wing`), and `ORG_LABEL_` (region
  abbreviation for a region tenant, wing abbreviation otherwise). Exposed as
  `CONFIG.WING_ABBREVIATION` / `WING_NAME` / `ORG_LABEL`. New **optional** Script
  Properties, all blank-derive: `TENANT_WING_ABBREVIATION`, `TENANT_WING_NAME`,
  `TENANT_CADETS_TENANT_DOMAIN`. `setupTenantConfig()` lists them.
- **`config-tenants/setup-hiwg.gs` + `hiwg-seniors.json` / `hiwg-cadets.json`** — a
  Hawaii Wing setup template mirroring `setup-region.gs`: paste-in
  `setupHiwgSeniorsScriptProperties()` / `setupHiwgCadetsScriptProperties()`, with
  `TENANT_WING=HI` and profile pre-set and the tenant-specific IDs/domains marked
  `FILL_IN`. Wing labels ("HIWG", "Hawaii Wing") are derived automatically.
- **`docs/NEW_TENANT_SETUP.md`** — bare-metal, end-to-end provisioning runbook for
  standing up a new tenant/wing from nothing (Workspace → GCP/APIs/scopes → service
  account + DWD → clasp push → Script Properties → CAPWATCH → spreadsheet → dry-run →
  triggers → cross-tenant), cross-linking the existing docs with a Hawaii worked example.
  Fills the gap where no from-scratch guide existed separate from the Admin Guide.

### Changed

- **`SyncOrgPaths.gs` (v1.1.0)** — OrgPath-sync email subject/footer and body now use
  `CONFIG.ORG_LABEL` / `CONFIG.WING` instead of literal `CAWG` / `CA`.
- **`SendRetentionEmail.gs`** — retention report footer uses `CONFIG.ORG_LABEL`.
- **`TransitionCompleteEmail.html` + `CadetTransitionMigrate.gs`** — the member-facing
  masthead and footer now render `{{wingName}}` (`CONFIG.WING_NAME`) instead of a
  literal "CALIFORNIA WING" / "California Wing IT".
- **`config.gs` · `UpdateGroups.gs` · `SquadronGroups.gs`** — the access-group
  description ("… accounts only") and the managed wing-scope display name use
  `WING_ABBREVIATION` instead of literal `CAWG`. (The `UpdateGroups`/`SquadronGroups`
  spots are inside California-specific display branches gated on `WING === 'CA'`; other
  wings fall through to full sentence-cased names, unchanged.)

### Notes

- No behavior change for the existing California or Pacific Region tenants: every new
  label derives to the same value it was before. Hawaii-specific data (domains, ORGIDs,
  Drive/Sheet IDs, service accounts) must still be supplied per tenant.

## [2026-07-17] — Cadet transition: forwarding-group external delivery fix

### Fixed

- **`CadetTransitionCleanup.gs` (v1.2.0)** — the post-deletion forwarding group (the freed
  cadet address that forwards to the new senior mailbox) now sets
  **`allowExternalMembers=true`**. The forward target is on the *peer* (senior) tenant —
  external to the cadets domain — so without this the group could not deliver to it; the
  previous version set only `whoCanPostMessage` (inbound) and would have silently forwarded
  nothing. Settings are now applied **before** the member is added, since a domain that
  restricts external members rejects the insert otherwise — and by that point the cadet
  account is already deleted (a group can't take an address a live user holds), so a failure
  there would strand a half-built forward. Added **`testForwardingGroup(addr, member)`** to
  prove the whole mechanism — including the domain's external-member policy — against a
  throwaway group before a real close depends on it. Never run in production yet (first close
  is ~2026-07-29).

## [2026-07-17] — Cadet transition: close reminder; license reaper armed on seniors

### Added

- **`CadetTransition.gs` · `CadetTransitionCleanup.gs` (v1.1.0)** —
  `remindPendingTransitionCloses()`, a daily **08:00** trigger that emails IT
  (`TENANT_ITSUPPORT_EMAIL`) when transitioned cadet accounts have passed their grace and are
  ready for the manual `closeCompletedTransitions(false)` — or are stuck past grace on a
  `DO NOT DELETE` hold. Read-only, keyed off `whyNotCloseable_` so its "ready" list is exactly
  what a real close would act on; silent when nothing is due. `armTransitionTriggers()` now
  installs **six** daily triggers (was five), adding this after the migration phases. Deletion
  itself stays manual — there is still no close/delete trigger.

### Operational

- **License reaper ARMED on the seniors tenant** — `LICENSE_DELETION_ARMED=true` set on the
  seniors project (2026-07-17). **cadets** and **region** remain dry-run. Preceded by a clean
  dry-run (0 deletable that day; the wing-transfer guard correctly spared a member who had
  moved to another wing; both PATRONs held for a human call). The next monthly
  `manageLicenseLifecycle` run reaps the accounts past their 30-day grace, minus any who renew.
  See [docs/ADMIN_GUIDE.md](docs/ADMIN_GUIDE.md) §9 (“Arming real deletion”).
- **All three tenants synced to this commit** (config **v1.6.0**); the seniors tenant was
  brought current at the same time.

## [2026-07-17] — Cite the signature duty-order requirement

### Documentation

- **`UpdateMembers.gs` (v1.13.2)** — comment-only. The signature already ordered duty
  assignments highest-to-lowest organizational level; this cites the exact CAP brand
  style-guide requirement ("organize assignments from highest to lowest organizational
  level", civilairpatrol.frontify.com/document/449893 → Email Signature) on
  `DUTY_LEVEL_ORDER`, and records that the one-duty-per-echelon selection in
  `getDutyBlock()` is a refinement layered on that ordering — the output stays strictly
  highest-first — not a departure from it. No behavior change.

## [2026-07-17] — Renamed the 'pacific' profile to the generic 'region'

### Changed

- **`config.gs` (v1.6.0)** — the third behavioral profile was named `pacific`; renamed
  to **`region`** so the shared code reads sensibly for any region-level tenant, not just
  the Pacific Region. Renamed with it: `clasp-targets/pacific.clasp.json` →
  `region.clasp.json`, `config-tenants/pacific.json` → `region.json`,
  `config-tenants/setup-pacific.gs` → `setup-region.gs` (+ its
  `setupPacificScriptProperties` → `setupRegionScriptProperties`),
  `docs/PACIFIC_DIFF.md` → `docs/REGION_DIFF.md`, and the `npm run *:pacific` scripts →
  `*:region`.

  **No coordinated property flip is required.** `PROFILE_ALIASES_ = { pacific: 'region' }`
  maps a legacy `TENANT_PROFILE=pacific` to `region` at resolution, so the live region
  project keeps working whether or not its Script Property is updated — without the alias
  it would have fallen through to the **seniors** profile (the `|| TENANT_PROFILES_.seniors`
  default), silently running a region tenant as a wing. Flipping the property to `region`
  (or re-running `setupRegionScriptProperties()`) is optional cleanup.

  "Pacific Region" / "PCR" is preserved wherever it names the actual deploying org or the
  `CAP-Pacific-Region` GitHub repo; only the profile identifier and its tooling changed.

## [2026-07-17] — Housekeeping: dead scope branch + duplicate helper

### Fixed

- **`UpdateChatSpaces.gs` (v2.0.1)** — `buildCommitteeSpaceName()` branched on
  `orgScope === 'SQUADRON'`, a value CAPWATCH never emits — `Organization.txt`'s Scope
  column is `UNIT`/`GROUP`/`WING`/`REGION`. Unit-level committee spaces therefore fell
  through with no prefix. Corrected to `'UNIT'`.
- **`utils.gs` / `UpdateMembers.gs` (v1.13.1)** — `toTitleCase()` was defined in **both**
  files with **different** implementations; in Apps Script's single shared namespace the
  later-loaded copy silently won, so behavior was load-order-dependent. Consolidated to
  one definition in `utils.gs`, keeping the stronger `\b\w+` variant (breaks on hyphens,
  apostrophes and periods: "Auburn-Starr", "O'Brien", "L.A." — the weaker copy produced
  "Auburn-starr", "O'brien", "L.a."). Also fixes the calendar and cross-tenant-contact
  callers that could have inherited the weaker behavior.

## [2026-07-17] — Account deletion gated behind a per-tenant Script Property

### Changed

- **`ManageLicenses.gs` (v1.3.0)** — real account deletion now requires the tenant to
  be explicitly armed via a Script Property, **`LICENSE_DELETION_ARMED`**. Both
  `deleteIneligibleSuspendedUsers()` and `deleteIneligibleWorkspaceUsers()` refuse to
  call `AdminDirectory.Users.remove()` unless that property is exactly `true`
  (case-insensitive, trimmed) — **even when passed `dryRun=false`** — and
  `manageLicenseLifecycle()` passes `dryRun = !armed`.

  Deletion is permanent and this Workspace edition has no Archived-User fallback, so
  it stays off until a property is set on the **specific project**. Previously arming
  meant editing `manageLicenseLifecycle()` to pass `false`: fragile (every `clasp
  push` reverts `src/` to the safe default) and dangerous (committing `false` would
  arm **every** tenant at once). A Script Property survives push and is per-project.

  Truth table (verified): `remove()` fires **iff** `dryRun === false` **and** the
  property is exactly `true`. Property absent → dry run. `'1'`, `'yes'`, or a typo like
  `'ture'` → stays safe. No tenant is armed by this change; the monthly reaper keeps
  reporting what it *would* take until someone sets the property deliberately.

## [2026-07-16] — LSCode weekly-trigger installer

### Added

- **`notifications/LSCodeNotify.gs`** — `installLSCodeWeeklyTrigger()`, a setup
  helper that installs the weekly `notifyLSCodeChanges` time trigger (Mondays
  ~07:00 America/Los_Angeles; edit the day/hour inline). Idempotent — it removes
  any existing `notifyLSCodeChanges` triggers first, leaving other handlers alone,
  so re-running never stacks duplicates. It **must be run signed in as the
  automation account**: a trigger runs as whoever creates it, and only that
  account owns the `AUTOMATION_SENDER_EMAIL` Send-As alias the digests need (see
  the 2026-07-16 identity entry). The trigger was already installed by hand this
  way; this captures the exact setup in version control so it is reproducible for
  future tenants. Covered by a test asserting the dedupe and the weekly schedule.

## [2026-07-16] — LSCode failure-summary survives a bad sender identity

### Fixed

- **`notifications/LSCodeNotify.gs`** — the IT failure-summary email no longer sets
  `from: AUTOMATION_SENDER_EMAIL`; it sends as the executing user (`name` display
  only). The live test on 2026-07-16 showed why: the most likely reason a whole run
  fails is that the executing account lacks the `AUTOMATION_SENDER_EMAIL` Send-As
  alias, which bounces every digest with "Invalid argument". The summary is the
  alarm that reports those failures — and it used the same `from`, so it failed the
  same way and the failure went unreported. Sending it as the executing user means
  the alarm gets through even when the sender identity is misconfigured; if the
  trigger is ever run under the wrong account, IT now gets an "attention needed"
  summary listing every org as failed. The digests themselves still send as
  `AUTOMATION_SENDER_EMAIL` (the sender commanders should see), so the trigger must
  still be owned by the automation account. Covered by a test that reproduces the
  wrong-identity bounce and asserts the alarm is delivered without a `from`.

## [2026-07-16] — Cadet → senior account transition

### Added

- **`CadetTransition.gs` · `CadetTransitionMigrate.gs` · `CadetTransitionDrive.gs`
  · `CadetTransitionContacts.gs` · `CadetTransitionCleanup.gs` (all v1.0.0)** — a
  new subsystem that carries a member across tenants when they age out of the
  cadet program at 21 or convert voluntarily after 18. Previously the cadets
  tenant just suspended and deleted the account ~30 days later, destroying the
  mailbox: this edition provisions **no** Archived User licenses, so deletion is
  the only way to reclaim a seat and there is no archive behind it.

  Runs on the **cadets** tenant only (`TRANSITION_CONFIG.ROLE === 'source'`);
  every entry point no-ops elsewhere. The seniors tenant is `destination` and its
  sole involvement is exempting these members from the Level I gate in
  `updateAllMembers()` so the receiving mailbox exists. The cadets tenant owns the
  whole lifecycle and **polls** the peer directory for the destination account
  rather than the two tenants signalling each other, so there is no shared state
  to drift. Single-tenant profiles (Pacific) set `ROLE: ''` and the feature is
  inert.

  The lifecycle: `detectCadetTransitions()` opens a row in a new **`Transitions`**
  sheet (authoritative state, so a human can see a stuck migration and step in) →
  `resolveTransitionDestinations()` fills in the senior address once that account
  appears → `migrateCadetTransitions()` imports Gmail in parallel batches →
  `migrateAllTransitionDrives()` copies owned Drive files →
  `migrateAllTransitionContacts()` copies personal contacts →
  `closeCompletedTransitions(false)` deletes the cadet account and forwards its
  freed address to the senior mailbox via a Group.

  Design points worth carrying forward:
  - **Suspended mailboxes stay readable.** SA impersonation works against
    suspended users (verified), so members are suspended on day 0 exactly as
    before — preserving the seat-cap discipline PATRON accounts blew in June
    2026 — and mail is migrated at leisure inside the hold window. Nothing here
    ever unsuspends an account.
  - **A 90-day hold, timed from `DetectedDate`, not the license clock.**
    `LICENSE_CONFIG` grace times a member *lapsing*, which a transitioning cadet
    is not doing. `TRANSITION_CONFIG.HOLD_DAYS` waits out National's fingerprint /
    Level I processing so a PATRON who later converts still has a mailbox to move.
    `deleteIneligibleSuspendedUsers()` now skips any CAPID with an open transition
    row (`getHeldTransitionCapids()`), so the license reaper never deletes a
    mailbox mid-flight. `FAILED` rows hold indefinitely and need a human.
  - **Resumable across the 6-minute limit.** A four-year mailbox holds thousands
    of messages, so each phase stops at a page boundary before the ceiling,
    records its cursor, and schedules its own continuation. Cursors advance only
    after a full page lands, so an expected time-limit stop cannot duplicate.
  - **Runs are serialized by a script lock** (`withTransitionLock_`) so an
    overlapping trigger and a manual run cannot import the same pages twice.
  - **`0 migrated` is a real value, not "unhandled."** The count fields treat a
    genuine zero (e.g. a member with no personal contacts) as done, via
    `isBlankField_()` — an earlier truthiness check read `0` as "never looked" and
    blocked the close for all four pilot members.
  - **The close is deliberately manual.** `armTransitionTriggers()` installs five
    daily triggers (detect → migrate, staggered 3–7 AM) but **no** close/delete
    trigger — the same discipline the license reaper landed on: automate the
    reversible copying, keep a human on the irreversible delete. Review with
    `closeCompletedTransitions(true)`, act with `(false)`. The triggers must be
    armed **as `automation@cawgcadets.org`** (triggers are owned by and visible
    only to their creator, and the completion email's Send-As is that account).

  Requires DWD scopes on the service accounts (set in the Admin console, *not*
  `appsscript.json`): cadets SA `gmail.readonly` + Drive/Contacts read; seniors
  (peer) SA `gmail.insert` + `gmail.labels` to import and `gmail.metadata` for the
  duplicate-guard check. Exercised on the initial pilot members. Going live is a
  deliberate operational step: run `armTransitionTriggers()` as the automation
  account and confirm with `listTransitionTriggers()`.

## [2026-07-15] — Commanders hear about background-check changes

### Added

- **`test/`** — the repo's first tests. `npm test`, plain Node, no framework and no
  dependencies, so it runs on a clean clone.

  Apps Script has no local test runner: the only way to run a `.gs` file for real
  is to push it to a live tenant and press **Run**, which here means production
  Drive, Gmail and Workspace. That is a bad place to learn that a notification
  module mails the whole wing. But a `.gs` file is only JavaScript whose globals
  arrive from the platform — so the harness reads the source, injects fakes for
  the globals it touches (`DriveApp`, `GmailApp`, `Logger`, `Utilities`, `CONFIG`,
  `PROFILE_`), and the module's own logic runs unmodified under Node. **Our code
  runs; Google is faked.** It does not replace a dry run against a tenant, it
  makes the dry run the second check rather than the first.

  Covers `LSCodeNotify.gs` end-to-end. The load-bearing case: a first run must
  mail nobody, because most seniors already hold an `A` and reading those as news
  would send every commander their entire roster. Inverting the first-seen rule
  fails both test files, so the guard is genuinely held rather than merely
  asserted.

  Two conventions, both learned from this codebase's own history: fixtures copy
  the **live** `Member.txt` / `Commanders.txt` headers verbatim, since modules
  resolve columns by name and an invented header would prove nothing; and stubs
  **throw** on anything unhandled rather than returning something plausible, so a
  stub cannot turn a real bug into a passing test.

  Nothing under `test/` reaches Apps Script — clasp's `rootDir` is `../src` and
  `.claspignore` excludes everything outside `src/`.

- **`notifications/LSCodeNotify.gs` (v1.1.0)** — the digest now dates each change,
  and the intended cadence is **weekly**.

  CAPWATCH publishes no date for an LSCode change. `Member.txt` carries `DateMod`
  (when the *record* was last modified) and `UsrID` (who modified it), but both are
  record-level — an address edit moves `DateMod` exactly as a background check
  does — and no background-check table exists anywhere in the extract. Quoting
  `DateMod` as "the date your member cleared" would therefore be wrong whenever
  anything else touched the record afterwards.

  So the digest reports the **window** instead: the change appeared somewhere
  between the last time we confirmed the old value and the run that saw the new
  one. On a weekly trigger that reads as "detected 8–15 Jul 2026", which is the
  resolution this data honestly supports. The footer says so in as many words, so
  a commander does not read the date as coming from eServices.

  The window is tracked **per member** (state file v2: `{ c, seen }`), not as one
  global last-run date. That is what makes a retry truthful — a digest that fails
  to send keeps its members' original `seen`, so when it lands a week later it
  still reports the week the change was really detected rather than the week of
  the retry. A v1 state file is re-baselined silently rather than misread.

- **`notifications/LSCodeNotify.gs` (v1.0.0)** — new module. Squadron commanders
  now get an email when a member under their command gains or loses their FBI
  background check. `Member.txt` `LSCode` carries that flag (`A` = passed, blank
  = not), and nothing in the codebase read the column until now — seniors and
  FIFTY YEAR members carry `A`; cadets and PATRON are blank. Note the column is
  **not** a per-person background-check flag: cadets over 18 have had a check and
  still show blank (checked against CAPID 612148), so it reflects the senior-side
  record. Both directions are reported: a grant is the expected case, but
  a clearance that stops being current is the one a commander needs sooner.
  Delivery is one digest per commander per run, so a squadron that fingerprints
  a dozen people at once produces one email rather than twelve.

  CAPWATCH is a snapshot with no history, so detecting a *change* needs saved
  state. `updateAllMembers()` already keeps one (`CurrentMembers.txt`) and this
  deliberately does **not** use it. That function is the account-provisioning
  job: if it sent commander mail and then threw before `saveCurrentMemberData()`,
  the next run would re-detect and re-send to every commander. Worse,
  `forceUpdateAllMembers()` writes that snapshot *without* diffing, so it would
  silently swallow pending transitions and those commanders would never be told.
  This module keeps its own `LSCodeState.txt`, runs on its own trigger, and
  cannot be affected by either.

  A member absent from the state file is recorded without notifying, which makes
  the first run after deploy a silent baseline (otherwise every commander would
  receive their entire roster) and keeps new members quiet — a joiner's existing
  clearance is not news. State advances per-org: a digest that fails to send
  leaves its members at their prior value to retry, so one bad send cannot
  re-mail everyone else. A unit with LSCode changes but no commander on record
  stays pending rather than being dropped, and is reported to IT support until a
  commander exists to receive it.

  Columns are resolved by header name, not position — per `docs/VERSIONING.md`
  and the `Expiration`-column lesson below, `parseFile()` strips the header and
  every positional index in this codebase is an unverified assumption. Off on the
  `cadets` profile (cadet records carry no LSCode at all) and on `pacific`
  (single-unit region HQ, pending a call from PCR/CC); on for `seniors` via
  `PROFILE_.RUN_LSCODE_NOTIFICATIONS`.

  Run `previewLSCodeChanges()` first — it sends nothing and writes nothing.

## [2026-07-15] — The ineligible-suspended reaper, repaired

### Fixed

- **`ManageLicenses.gs` (v1.1.0)** — `deleteIneligibleSuspendedUsers()` has never
  deleted anything since it was added in June 2026. Its `fields` selector asked
  for `employeeId`, which is **not** a field on the Admin SDK Directory User
  resource (it belongs to the People API), so the API returned 400 `Invalid field
  selection employeeId` and the function threw on its first page, every run.
  `manageLicenseLifecycle()` wraps that call in try/catch and filed the error into
  `summary.errors`, so it failed quietly into the monthly report for a month. The
  June 2026 cleanup of 257 stale accounts was `deleteIneligibleWorkspaceUsers()`,
  a different function.

  Removing `employeeId` alone would have been dangerous, which is why it wasn't
  done as a drive-by: the grace period measured **days since last login**, not
  days since the member lapsed, and every suspended account was already past that
  cutoff. The first successful run would have deleted the entire suspended
  population at once, permanently — this edition has no Archived User licenses, so
  there is no archive and no undo.

  Grace is now measured from the member's **CAPWATCH `Expiration`** date
  (`Member.txt` column 16, verified by name against the header — `parseFile()`
  strips it, so the index had never been checked). `lastLoginTime` is retained for
  human context only and drives no decision. Also fixed: Google returns
  `lastLoginTime` as the **Unix epoch** for accounts that never signed in rather
  than omitting the field, so the long-advertised `creationTime` fallback was dead
  code and such accounts read as ~20649 days stale — an account created yesterday
  and suspended today sorted as maximally stale.

  Behaviour changes: a **current member is never auto-deleted**, even when
  ineligible by type (a PATRON's expiry is in the future and cannot date their
  conversion, so there is nothing safe to measure) — they are surfaced for a human
  instead. **No CAPWATCH record** now means *deletable*: the extract retains only
  a rolling window of expired members (~3 months observed — `EXPIRED` rows carry
  just 3 distinct month-ends), so absence implies a lapse far beyond any grace
  period. That inference is only sound on a complete extract, hence the new
  `MIN_MEMBER_ROWS` guard.

  It is also only sound if the member actually lapsed, and that is checkable:
  lapsing gets you suspended on the next sync, and a suspended account cannot
  sign in — so a no-record account that was alive *after* the oldest lapse the
  extract still retains did not lapse at all. A **wing transfer** looks exactly
  like this: leave CAWG and you vanish from our extract while remaining a current
  member elsewhere. A live dry run surfaced one (CAPID 697618, last login 26 days
  prior, confirmed transferred to Nevada Wing), which the rule would otherwise
  have deleted. The window boundary is derived from the data each run rather than
  hardcoded, so it tracks CAPWATCH rather than assuming. `lastLoginTime` never
  sets the grace period, but it is used here to falsify it.

### Added

- **`ManageLicenses.gs`** — the **departure register**, giving members who leave
  the wing the same grace as members who lapse.

  No departure date is reachable, which was checked rather than assumed.
  Workspace records no suspension time. CAPWATCH *does* publish transfer dates in
  `MbrTransfer.txt` (`CAPID`, `TransferDate`, `ToORGID`, `FromORGID`), and an
  authoritative date would have beaten any proxy — but the table is **inbound
  only**: every `ToORGID` is a CAWG org, and the wing-scoped extract drops
  departing members wholesale, transfer row included. The live case (CAPID 697618,
  confirmed transferred to Nevada Wing on 07-02) appears in none of its 1368 rows.
  `debugCapwatchTransferFile()` re-checks this cheaply; if CAPWATCH ever carries
  outbound transfers, `TransferDate` should replace the register outright.

  `lastLoginTime` is not a stand-in either — someone who transfers today after six
  quiet months would read as six months elapsed and be deleted at once, which is
  the exact mistake the expiry basis exists to undo (the live case had last logged
  in 26 days before a departure that was hours old; a login-based timer would have
  left him 4 days instead of 30).

  So the timer runs from when the job **first saw them gone**. A transfer suspends
  them on the next sync, so first sighting lands within a sync cycle of the real
  departure, and any error runs long rather than short. State lives in a Script
  Property (`LICENSE_DEPARTED_FIRST_SEEN`, CAPID → ISO date) because `clasp push`
  overwrites code and would otherwise reset every timer on each deploy. The
  register is rewritten each live run from whoever is still departed, so returners
  prune themselves; it is written *before* deletions, so a mid-run crash cannot
  silently hand everyone a fresh 30 days; and **dry runs never write it**, so a
  preview cannot start a deletion clock. An unreadable register restarts timers
  rather than expiring them. `resetDepartedRegister()` clears it by hand.

  Worth noting the backstop: a member wrongly caught here spends 30 days suspended
  before anything irreversible happens, and a locked-out member complains.

  The function now **defaults to a dry run** and returns a result object
  (`candidates` / `withinGrace` / `activeIneligible` / `unknownExpiry`) rather than
  a bare array. `previewIneligibleSuspendedDeletion()` is a new entry point, also
  wired into `previewLicenseLifecycle()`, which previously covered only
  `previewArchival()` / `previewDeletion()` and left this path with no dry run at
  all. `manageLicenseLifecycle()` still calls it in dry-run mode; the monthly
  reaper is **not armed** until that flag is flipped. The report email reports
  candidates, spared-within-grace, and needs-review separately, and no longer
  claims deletions that did not happen.

### Added

- **`config.gs` (v1.5.0)** — `MIN_MEMBER_ROWS` (1000). `deleteIneligibleSuspendedUsers()`
  treats a missing CAPWATCH record as proof of a long-ago lapse; a truncated
  `Member.txt` would therefore make thousands of current members look deletable.
  `parseFile()`'s fallback parser can quietly return a partial row set, so the
  deletion path now refuses to run below this floor. The seniors extract carries
  ~5,000 rows.

- **`ManageLicenses.gs`** — `debugCapwatchMemberExpirationColumn()`, a read-only
  diagnostic printing the `Member.txt` header, the distribution of expiration
  values by member status, and raw rows for given CAPIDs. Written to verify the
  column-16 index before an irreversible policy was built on it.

## [2026-07-14] — Secondary-domain aliases for listed accounts

### Added

- **`SecondaryDomainAliases.gs` (v1.2.1)** — new module giving accounts a second
  address that keeps the local part of their primary but swaps in a secondary
  domain (`jane.doe@cawgcap.org` → `jane.doe@cawg.cap.gov`), as a **directory
  alias** via `AdminDirectory.Users.Aliases.insert`. Driven by a new, optional
  `Secondary Aliases` tab, which is a **curated opt-in list, not the roster** —
  only listed accounts are touched, and new members are not enrolled
  automatically. Entry points: `addSecondaryDomainAliases()` (trigger-safe) and
  `previewSecondaryDomainAliases()` (dry run, manual only).

  Gated on a new `TENANT_SECONDARY_EMAIL_DOMAIN` Script Property, blank on cadets
  and pacific, so the shared code is an explicit no-op there rather than an error.
  A preflight check confirms the domain is verified in the tenant, turning what
  would be one opaque HTTP 400 per row into a single actionable message.

  Unlike `addAlias()` in `UpdateMembers.gs`, a 409 does **not** fall back to a
  numbered variant (`jane.doe.1@`) — an address that does not mirror the primary
  defeats the purpose — and conflicts latch: they report once, then are skipped
  until the row changes, rather than logging an ERROR every night forever.

  Requires the new `admin.directory.domain.readonly` scope, so **every tenant
  re-authorizes on next run**.

- **`config.gs` (v1.4.0)** — added `SECONDARY_EMAIL_DOMAIN`, sourced from the new
  `TENANT_SECONDARY_EMAIL_DOMAIN` Script Property and wired into
  `setupTenantConfig()`. Blank on cadets and pacific.

### Fixed

- **`UpdateMembers.gs` (v1.5.0)** — `updateGmailSendAsDisplayName()` patched only
  `sendAs/{primaryEmail}`, so every *alias* Send-As identity kept whatever display
  name Gmail auto-assigned when the alias was created and went stale the moment the
  member was promoted. It now mirrors the name onto the user's org-owned alias
  identities as well (new step 3 + `updateSendAsDisplayNameForOrgAliases_()`), which
  fixes both callers — new-user setup and the `updateAllSendAsNames()` backfill.

  Only identities on `CONFIG.EMAIL_DOMAIN` / `CONFIG.SECONDARY_EMAIL_DOMAIN` are
  touched (`isOrgOwnedSendAs_()`, exact domain match): members add their own personal
  addresses as Send-As identities, and renaming someone's private Gmail to their CAP
  rank would be wrong. This is the concern that left the domain filter commented out
  in `updateSignatureForAllAliases()`. Patches only when the name differs, so a
  settled roster costs one list call per user and no writes.

  > ⚠️ The code half is inert on its own. `updateAllSendAsNames()` is the only bulk
  > caller and **was never on the nightly chain**, so promotions did not reach the
  > Gmail Send-As name for *anyone* — primary included — except when someone ran the
  > backfill by hand. It is now listed in [ADMIN_GUIDE §8](docs/ADMIN_GUIDE.md) at
  > 8–9 AM; **the trigger must still be created per tenant.**

- **`UpdateMembers.gs` (v1.6.0)** — new accounts received a **blank signature**.
  `runDelayedGmailSetup()` rebuilt `{ capsn }` from its queued Script Properties
  record and handed that to `generateEmailSignature()`, which therefore rendered an
  empty name line, an empty `(M)` phone row, and a duty of "Member" — then pushed it
  to the account five minutes after creation, on every tenant.
  `queueForDelayedGmailSetup()` now carries the fields the generator reads
  (`signatureMember`). Records queued by the older code have no such field; rather
  than reproduce the bug they are skipped with a warning, and the account still gets
  its Send-As display name.

  Also: an ungraded senior (CAPWATCH rank `SM`) rendered literally as "SM Jane Doe".
  The CAP style guide does not permit `SM` as a grade designation, and
  `getPublicRank()` has no mapping for it, so it passed straight through. New
  `getSignatureName()` shows ungraded members by name with a middle initial —
  "Jane M. Doe" — until their first promotion, after which the normal grade + name
  form resumes. A blank Rank column is treated the same way.

  > Existing users are deliberately untouched: `pushAllSignatures()` remains manual
  > and off the §8 schedule, by request. Only newly created accounts get a signature.

- **`UpdateMembers.gs` (v1.7.0)** — `updateSignatureForAllAliases()` wrote to **every**
  Send-As identity a member had, including personal accounts they had added themselves:
  the only guard was a hard-coded `endsWith("@pcrcap.org")` check that shipped
  **commented out**. It now writes only to identities on a domain this tenant owns, via
  the same `isOrgOwnedSendAs_()` used by the display-name sync.

  Note the old check was doubly wrong for this repo even if it had been enabled: it
  named the **Pacific** tenant's domain, so on seniors or cadets it would have matched
  nothing and skipped every identity — and it permits exactly one domain, so
  secondary-domain aliases would never receive a signature.

  Signature name lines now also include the member's **suffix** (`Maj. Isaac Wilson IV`),
  which `getSignatureName()` was dropping.

- **`UpdateMembers.gs` (v1.8.0)** — `generateEmailSignature()` reconciled with the CAP
  brand style guide. The guide itself lives behind a JS/auth wall on Frontify, so the
  reference used was the template inside CAP's own signature generator
  (`cap-brand-tools`, `signature-generator/script.js`), which emits the canonical block.

  | Element | Was | Now |
  |---|---|---|
  | "Civil Air Patrol, U.S. Air Force Auxiliary" | `<h2>`, normal weight, `margin 0 0 20px` | `<p>`, **bold**, `margin 0 0 5px` |
  | Duty block | all non-assistant duties, `line-height 12px`, `'Member'` when none | **max 2**, sorted highest→lowest org level, `line-height 14px`, **omitted** when none |
  | Phone row | always emitted — a bare `(M)` when the member had no phone | omitted when there is no phone |
  | Logo | no `width`/`height`/`alt`, negative margin | `width=200 height=42`, `display:block`, `alt` text |

  Two latent bugs fell out of this. `getDutyBlock()` checked emptiness *before* filtering
  assistants, so a member holding only assistant duties produced an **empty `<h2>`**; and
  `'Member'` was never a duty assignment — the guide's generator simply drops the element.

  Duty ordering uses CAPWATCH's `level` (`UNIT`/`GROUP`/`WING` per
  [API_REFERENCE](docs/API_REFERENCE.md), plus `REGION`/`NAT` for the region tenant).
  Unrecognized levels sort last rather than being guessed at; `Array.sort` is stable, so
  they retain CAPWATCH's own order.

  > Both open items from this entry — the logo host and the duty line's org prefix — are
  > resolved in v1.9.0 below.

- **`UpdateMembers.gs` (v1.9.0)** — signature duty lines named the **wrong org**. They
  were prefixed with `member.orgName`, the member's *home unit*, regardless of where the
  duty is actually held: a squadron member with a wing-level duty read "San Jose Sr Sqdn
  80 Director of IT". `addDutyPositions()` and `addCadetDutyPositions()` now carry the
  duty's own `orgName` (from the org that duty record points at), and `getDutyBlock()`
  uses it, falling back to the home unit.

  Unit names are also expanded for display by `formatOrgName_()`:
  `SAN JOSE SR SQDN 80` → **San Jose Senior Squadron 80**. Expansions: `Sq`/`Sqdn` →
  Squadron, `Cdt` → Cadet, `Comp` → Composite, `Sr` → Senior. Matching ignores case and
  a trailing period, and is **scoped to org names only** — it never runs over a person's
  name, so a member whose suffix is `Sr` stays "Vance Sr" rather than "Vance Senior".

  **Logo moved off the Frontify CDN token URL** to the copy served alongside CAP's own
  generator (`cap-brand-tools.netlify.app/.../LogoNoAux.png`) — a 2000×415 master
  rendered to 200×42, so it stays sharp on high-DPI displays.

  > ⚠️ Explicitly **not** the generator's own `LOGO_URL_OUTPUT`
  > (`civilairpatrolmac.github.io/CAP-Brand-Tools/...`): that URL **404s**, as does the
  > whole GitHub Pages site. Signatures produced by CAP's official tool therefore have a
  > broken logo. Verify any replacement with a HEAD request — a dead URL here is silent
  > in the logs and only shows up as a broken image in mail that has already been sent.

- **`UpdateMembers.gs` (v1.10.0)** — duty titles are used **verbatim** from CAPWATCH,
  plus one rename. Checked against a real CAWG `DutyPosition.txt` (4,085 rows,
  71 distinct titles):

  - The `Duty` column already holds full, **echelon-correct** titles. CAPWATCH varies
    them itself — `Information Technologies Officer` at `UNIT`/`GROUP` vs
    `Director of IT` at `WING`; likewise `Safety Officer` vs `Director of Safety`. No
    echelon logic belongs in this code.
  - **Do not add office-symbol expansion.** The symbol (`IT`, `AE`, `DC`) is a separate
    `FunctArea` column that this code never reads. `docs/API_REFERENCE.md` showed
    `id: 'CC'`, conflating the two — **corrected in this change**; the real value is
    `id: 'Commander'`.
  - `DUTY_TITLE_OVERRIDES` renames the retired `Recruiting & Retention Officer` →
    `Recruiting Officer`, per the ICL to CAPR 30-1. 2 of CAWG's rows still carried the
    old form against 69 correct ones. Fixing the record in eServices is the real
    remedy; this only stops a stale row printing a retired title.
  - Titles are whitespace-collapsed: `Communications Officer ` ships with a trailing
    space on all 196 of its rows.

  Verified across every distinct title in the feed: exactly **one** is rewritten (the
  retired Recruiting form) and the other **70 pass through untouched**, none of which
  are bare office symbols.

  > `Lvl` in real CAWG data only ever contains `UNIT`, `GROUP`, `WING` (one row has
  > trailing whitespace, which `dutyLevelRank_()` trims). `REGION`/`NAT` remain in
  > `DUTY_LEVEL_ORDER` for the Pacific tenant.

- **`UpdateMembers.gs` (v1.11.0)** — added **`previewSignatureForMember()`**, a
  read-only render of a single member's signature to the execution log.

  There was previously no safe way to look at a signature before it reached a member:
  `pushAllSignatures()` writes to every member at once, and the only other path fires
  five minutes after an account is created — so inspecting the output meant either
  spamming the wing or burning a licence on a throwaway account. This makes **no
  Gmail or Directory calls at all**; it reads CAPWATCH and formats a string.

  Set `SIGNATURE_PREVIEW_RUN_INPUTS.CAPID` at the top of the file and Run it (Apps
  Script cannot pass arguments to an editor Run — same convention as
  `GROUP_ADMINISTRATION_RUN_INPUTS`). It logs the name line, duty block, phone, and
  which identities *would* receive it, then the raw HTML last so it can be lifted out
  of the log and opened in a browser.

- **`UpdateMembers.gs` (v1.12.0)** — three defects the first live preview exposed.

  **A wing role could crowd out a squadron command.** Sorting on echelon alone, then
  taking two, meant a member holding two wing duties and a squadron command showed
  both wing rows and **dropped the command entirely**. The block now takes at most
  **one duty per echelon** before filling the second slot, so the span of someone's
  roles survives. If all their duties sit at one level the second slot is still used,
  rather than wasted.

  **Ties within an echelon were arbitrary.** CAPWATCH has no primary-duty flag, so two
  wing roles were ordered by whatever eServices listed first — putting "Web Security
  Administrator" ahead of "Director of IT". `dutyTitleRank_()` now breaks the tie on
  the title text: command, then directors, then everyone else. `Array.sort` is stable,
  so equal ranks keep CAPWATCH's order.

  **Wing and region orgs were named for the HQ unit.** Every one of CAPWATCH's 54
  wings is `<STATE> WING HQ`, so the line read "California Wing Hq Director of IT" —
  both wrong and mis-cased, since `toTitleCase()` lowercases before capitalising.
  `formatOrgName_()` now takes the org's scope and, for `WING`/`REGION`, drops
  everything after the echelon: `CALIFORNIA WING HQ` → **California Wing**. This also
  handles the one region not named "... REGION HQ" — `PACIFIC REGION CAP` →
  **Pacific Region** — which matters for that tenant. `HQ` no longer renders as "Hq".

  `addDutyPositions()`/`addCadetDutyPositions()` carry `orgScope` alongside `orgName`.
  Keyed on scope rather than the name, so a unit that merely has "wing" in its title
  is untouched.

  > Consequence worth knowing: with the cap at two, a member holding wing **and** group
  > **and** squadron duties still loses the lowest — the two highest echelons win.

- **`UpdateMembers.gs` (v1.12.1)** — `ORG_NAME_EXPANSIONS` gained `SQD` → Squadron,
  `GP`/`GRP` → Group, and `CALIF` → California.

  CAPWATCH spells Squadron three ways: `SQDN` (585 orgs), `SQ` (45) and `SQD` (1 —
  "FALLBROOK SENIOR SQD 87", a California unit, which was rendering as "Fallbrook
  Senior Sqd 87"). Trailing periods are stripped before lookup, covering `SQ.` and
  `SQDN.` too. `GP`/`GRP` appear nowhere in CAPWATCH's org list today — all 147 groups
  spell it out — but they are conventional and cost nothing to cover. `CALIF` fixes
  "CENTRAL CALIF GROUP 6" → **Central California Group 6** and "CALIF WING HQ SQ" →
  **California Wing HQ Squadron**; orgs already spelling out `CALIFORNIA` are
  unaffected, since lookup is whole-word.

  Verified by rendering **all 77 California orgs**: every remaining short word is a
  proper noun (San, Los, Diego, Santa, Beale, Pancho …) and correctly untouched.
  "Eugene L. Carnahan Cadet Squadron 85" keeps its initial, since the period is
  stripped only for the lookup, not the output.

- **`UpdateMembers.gs` (v1.13.0)** — `getPublicRank()` had **no cadet grades at all**,
  so a cadet signature rendered the raw CAPWATCH value: "C/Amn Jane Doe",
  "CADET Jane Doe". All 15 are now mapped.

  Display forms come from the grade list in CAP's own signature generator
  (`cap-brand-tools`), including the `Cadet ` prefix its `buildDisplayName()`
  prepends — so `C/CMSgt` → **Cadet Chief Master Sgt.**

  Two traps worth recording:

  - **CAPWATCH's cadet spellings are not the senior ones with `C/` glued on.** They
    carry no internal space: `C/2dLt`, `C/1stLt`, `C/LtCol` — against the senior
    `2d Lt`, `1st Lt`, `Lt Col`.
  - **`CADET` is C/AB**, the entry grade, and a *real* grade → "Cadet Airman Basic".
    It must not be folded into the ungraded-senior case that `isUngradedRank_()`
    handles for `SM`.

  Verified against every rank in a real CAWG `Member.txt`: all 14 cadet grades present
  (743 `C/Amn` … 9 `C/LtCol`) now map, and no non-cadet member carries a cadet-style
  rank, so senior output is untouched. `C/Col` is mapped for completeness though
  CAWG has none today.

  > ⚠️ Blocked on `cawg.cap.gov` being added and verified as a secondary domain of
  > the seniors tenant. As a subdomain of `cap.gov` this needs a DNS TXT record
  > published by CAP National; aliases **cannot** be created on the domain until
  > then, and there is no way to pre-create them and have them activate on
  > verification. Until it is verified `addSecondaryDomainAliases()` logs the
  > preflight error and exits without touching any account.
  >
  > `previewSecondaryDomainAliases()` deliberately still runs in that state (warning
  > rather than bailing), so the tab can be populated and validated ahead of the
  > domain going live — it resolves the address each row would get and flags any
  > listed account that does not exist.

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

- **`SquadronGroups.gs` (v1.3.0) + `config.gs` (v1.2.2)** — `SQUADRON_DISTRIBUTION_TOGGLES`
  was a hard-coded const, so the cadet tenant was creating senior-only lists
  (`.seniors`, Deputy Commander for Seniors) that don't apply there. The toggles
  now come from `PROFILE_.SQUADRON_DISTRIBUTION_TOGGLES` in `config.gs` (selected
  by the `TENANT_PROFILE` Script Property), read via
  `getSquadronDistributionToggles_()`; the const is a fallback default only. Same
  mechanism as the other per-tenant behavior — a shared-code `clasp push` can't
  make a tenant create the wrong lists.
- **Cadets profile = all-hands + cadets + parents lists.** Disabled on the cadet
  tenant: `.seniors` (no seniors here) and the command-staff lists (Commander /
  Deputy Commander / Deputy Commander for Cadets — those are senior duty positions
  whose holders have wing accounts, so the lists would be empty). `.all` is kept
  intentionally: on a cadet-only tenant it duplicates `.cadets`, but the lists
  already exist and are retained in case something references them. Seniors profile
  unchanged; pacific = all off (single-unit region, squadron sync not triggered there).
- **Cleanup follow-up:** disabling a toggle stops managing those lists but does
  not delete already-created groups. The existing `ca###.seniors@cawgcadets.org`
  and cadet command-staff groups become orphans on the cadet tenant and should be
  removed (stage them with `groupAdministration_stageOrphanedSquadronGroups()`,
  then `groupAdministration_bulkDeleteGroupsFromSheet`). `.all` groups are kept.

### Added — group-admin helpers (`groupAdministration.gs`)

- `groupAdministration_auditReceiveListPosting()` — read-only audit of
  `whoCanPostMessage` / `allowExternalMembers` on managed `.cadets`/`.parents`/`.all`
  receive lists; flags any whose posting policy would reject cross-tenant fan-out.
  Run on the tenant that owns the receiving groups (e.g. cadets).
- `groupAdministration_stageOrphanedSquadronGroups(sheetName)` — tenant-aware; writes
  squadron groups whose list type is currently disabled by `SQUADRON_DISTRIBUTION_TOGGLES`
  to a worklist tab (default "Delete Groups") for review. Reads groups + writes the sheet
  only; deletion stays a separate manual step (`groupAdministration_bulkDeleteGroupsFromSheet`).

## [2026-07-09] — Pacific go-live

The reconciled `src/` was deployed to the live "PCR Automation" project (`TENANT_PROFILE=pacific`)
and verified end-to-end; triggers rebuilt under `automation@pcr.cap.gov`. **All three tenants now
run identical source, differentiated only by configuration** — the reconciliation goal.
_This supersedes the "not yet deployed to Pacific / on hold pending 2SV" notes in the entries
below, which were accurate when written._

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
