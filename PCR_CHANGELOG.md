# PCR Changelog

Pacific Region (PCR) fork-specific changes to the CAPWATCH / Google Workspace
automation, layered on top of the upstream `cap-miwg/gsuite-automation` project.
Upstream changes are tracked in [CHANGELOG.md](CHANGELOG.md); **this file records
only what the Pacific Region deployment adds or diverges on.**

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).
Individual source files carry their own SemVer version in their header
(see [docs/VERSIONING.md](docs/VERSIONING.md)); the per-file version is noted
next to each entry below.

## [2026-08-19] — The admin help-desk app now runs on the cadets tenant too (`admin-webapp/`)

Adds the cadets clasp target and npm scripts (`push:admin:cadets` etc.) alongside the existing
seniors ones, and fixes three things in the shared source that were silently seniors-only.

### Added

- **`admin-webapp/Setup.gs`** — per-tenant setup values (seniors + cadets), each behind its own
  run-by-hand function (`setupSeniorsAdminWebApp()` / `setupCadetsAdminWebApp()`), sourced from
  `config-tenants/<tenant>.json`. `admSetupApply_()` **refuses** to write a `TENANT_PROFILE`
  that contradicts the one already on the project — pushed as-is to the cadet project, the old
  hardcoded seniors values would have written `@cawgcap.org` and the seniors CAPWATCH folder
  onto the cadet tenant without throwing, so every lookup would answer from the wrong extract
  while every action targeted a domain that tenant does not own.

### Changed

- **`Config.gs`** — `WEBAPP_CADET_TOOLS_URL` renamed to `WEBAPP_PEER_ADMIN_URL` (the old name
  is still read as a fallback), and `TENANT_CADETS_TENANT_DOMAIN` generalized to `XT_PEER_DOMAIN`
  (old name still read as a fallback too) — both were cadet-specific naming for a general
  "the other tenant" idea that now means `cawgcadets.org` on seniors and `cawgcap.org` on
  cadets. `WEBAPP_2SV_SETUP_GROUP` now also accepts `TENANT_2SV_SETUP_GROUP` as a fallback,
  matching the name `src/accounts-and-groups/TwoSvSetupGroup.gs` already uses for the same
  group. The cadet tenant's own 2SV group is deliberately left **blank** by default —
  `2sv-setup@cawgcap.org` lives on the other Workspace and cannot hold cadet accounts.

### Fixed

- **"This member is elsewhere" notice** pointed at the cadet Workspace unconditionally. On the
  cadets tenant, the members this app cannot help are **seniors** — unfixed, it would have
  pointed an admin at the page they were already standing on. The label and destination now
  come from the member's type and `XT_PEER_DOMAIN`.

## [2026-08-19] — The ".gov Account Processing" web app now runs the whole process on Add (`secondary-alias-webapp/` 1.1.0)

The live secondary-alias add/remove page (deployment `AKfycbz5mBOJ…`, its own script
project) already created the directory alias the instant you clicked Add. It now also
configures Gmail *Send mail as* and emails the member **at that moment** — so a member added
through the page gets the same complete treatment the nightly run gives a hand-added row,
without waiting until morning.

### Added

- **`secondary-alias-webapp/`** pulled into the repo and version-controlled (it previously
  existed only in the cloud). New clasp target `clasp-targets/secondary-alias-webapp.clasp.json`
  and npm scripts `*:alias:seniors`; `.claspignore` whitelists the dir like the other
  web-app projects.
- **`Notify.gs`** — `apiAddByCapid`, on a genuine new alias (`ADDED —`), now calls
  `webappConfigureSendAs_` (auto Send-As via the same impersonation `src/` uses) and
  `webappMaybeSendWelcome_` (the CAPR 120-1 §6.9 + compose-switching email). Both run
  **outside the sheet lock** and are **best-effort** — a Gmail hiccup cannot undo the alias
  or fail the request. Response gains `sendAsConfigured` / `welcomeSent`, surfaced in the UI.

### Why it had to live in the web app

An alias the page creates reads as `already present` by the time the nightly
`addSecondaryDomainAliases()` fires, so the nightly notifier would skip it — the member would
get an alias and **nothing else**. `Notify.gs` is a hand copy of the nightly logic;
`test/SecondaryAliasWebAppNotify.test.js` (30 assertions) pins the two together, including a
byte-for-byte template check and a render-parity check across all send-as outcomes.

### Deploy requirements (see `secondary-alias-webapp/README.md`)

- New scopes (`gmail.settings.sharing`/`.basic`, `gmail.send`, `script.send_mail`,
  `script.external_request`, `userinfo.email`) — the deploying account must **re-authorize**.
  `userinfo.email` is required by `Session.getActiveUser().getEmail()` in `Auth.gs`; without
  it the app denies everyone (fails closed — an availability bug, not a security hole). The
  original code relied on the old runtime behavior of handing a same-domain web app the
  caller's address without the scope; the current runtime requires it.
- New Script Properties on that project: `SA_IMPERSONATION_EMAIL` / `SA_PRIVATE_KEY` (Send-As),
  `TENANT_AUTOMATION_SENDER_EMAIL` / `TENANT_ITSUPPORT_EMAIL` (email). The email's wing name
  and ORG label DERIVE from `TENANT_WING` (`CA` → `California Wing` / `CAWG`), mirroring
  `src/config.gs`, so no separate `TENANT_WING_NAME` is needed.
- The live deployment is pinned to `@1`, so `clasp push` updates HEAD only; it goes live only
  on a **new version deploy**. Until then the alias still creates and Send-As/email degrade
  gracefully.

## [2026-08-19] — New secondary aliases are made send-ready and the member is told (`SecondaryDomainAliases.gs` 1.4.0)

When `addSecondaryDomainAliases()` newly **ADDs** a secondary-domain alias, it now makes the
address immediately usable and tells the member — so nobody follows setup steps only to find
the address does not work.

### Added

- **Auto-configures Gmail *Send mail as*** for the member on a new alias, so the address can
  **send**, not just receive — via the same service-account impersonation
  (`getImpersonatedToken_` + the `sendAs` REST endpoint) the send-as **name** sync already
  uses. Admin accounts (which cannot be impersonated for settings) and any setup error fall
  back to manual steps in the email. Uses `gmail.settings.sharing`, already in the manifest.
- **Emails the member** (`SecondaryAliasWelcomeEmail.html`): that the address exists; whether
  sending was set up for them or the manual *Send mail as* steps; **how to switch between
  their primary and the new address in the Compose window** (the From menu, per message, on
  desktop and mobile); and the **CAPR 120-1 §6.9** reminder that a `.gov` address is for
  official CAP business only — no commercial or fundraising use. Sent from
  `TENANT_AUTOMATION_SENDER_EMAIL`, reply-to `TENANT_ITSUPPORT_EMAIL`.
- The email sets the expectation that a brand-new address can take **up to a day** before
  Gmail will send from it (matching the existing "sendable within ~24 hours" note), so a
  not-yet-propagated address is not mistaken for a broken one.

### Behavior / safety

- Both actions fire **only on a genuine insert**, so a settled list (`OK — already present`)
  does nothing and re-runs never re-notify — no backfill blast.
- **Best-effort by contract:** the alias is already created before either runs, so a failure
  in Send-As setup or the email is swallowed to a `WARN` and never fails the run or the
  alias. A flaky Gmail call must never make a successful provisioning look broken and get
  "fixed."
- Opt out of the email per tenant with Script Property `SECONDARY_ALIAS_NOTIFY=false`.
  Optional `TENANT_ITSUPPORT_URL` sets the footer link (defaults to `https://support.pcrcap.org`).
- `test/SecondaryAliasWelcome.test.js` — 65 assertions: the opt-out toggle; Send-As setup
  degrading safely on admin / missing-credentials / API-error / token-throw; the email
  adapting between the "already set up" and manual variants; both variants carrying the CAPR
  citation and the compose-window switching note with no stray `{{placeholder}}`; and neither
  a throwing send nor a throwing Send-As call propagating.

> **Scope note.** This covers the nightly/hand-added path, the only secondary-alias creator
> present in this branch. The separate CAPID-driven alias web app (commit `330c99c`, never
> merged) would need the same wiring if revived — separate script projects cannot share the
> code.

## [2026-08-18] — two stuck migrations: bandwidth quota and an uninitialised mailbox

### Fixed

- **`CadetTransitionMigrate.gs`** — a migration failed outright at 1500 messages with
  *"Bandwidth quota exceeded … Try reducing the rate of data transfer"*. Gmail enforces that
  ceiling on raw-message downloads and throws it from `UrlFetchApp.fetchAll` **for the whole
  batch**, not as a per-request 429 — and the handler treated any fetchAll-level throw as
  fatal, so the retry logic never engaged. Transient throws (bandwidth/quota/rate/timeout) now
  back off hard (30s, 60s, 90s…) and retry the batch. If a mailbox keeps hitting it after
  that, lower `MIGRATE_PARALLEL_`.

- A 400 *"Precondition check failed"* on `messages.import` now reports what it actually means:
  the **destination mailbox is not initialised** — a new senior account whose owner has never
  signed in, so Gmail is not provisioned for it. Previously it landed in the sheet as a bare
  API error with no indication that the fix is "have the member sign in once, then retry".

## [2026-08-18] — move the destruction guards to expiry

### Fixed

- **`CadetTransitionCleanup.gs`** — since v2.0.0 the close **parks** the account instead of
  deleting it, which moved the irreversible step to `expireParkedAccounts()` — but the
  Drive/Contacts guards stayed behind on the close. Expiry checked **mail only**, so a parked
  account (live for the full 12-month window) had anything added to its Drive or Contacts in
  that year deleted unchecked, and a `DO NOT DELETE` note did not block it.

  Expiry now refuses via `whyNotExpirable_` (DO NOT DELETE note, unhandled Drive or Contacts,
  missing destination) and **re-sweeps all three** — mail, Drive, contacts — before deleting.
  Every sweep is idempotent (cursor / appProperties / userDefined marker), so it moves only
  what is new.

## [2026-08-18] — reap triggers whose function no longer exists

### Fixed

- **`CadetTransition.gs`** — a trigger for `runCadetTransitionLifecycle` (a duplicate
  orchestrator added and removed on 2026-08-03) kept firing nightly at 03:05, failing with
  *"Script function not found"* until 2026-08-16. **Deleting a function does not delete its
  triggers**, and triggers are owned by — and visible only to — the account that created them,
  so one armed under a different identity cannot be seen or removed by anyone else.

  `TRANSITION_RETIRED_HANDLERS_` now lists handler names that no longer exist, so
  `disarmTransitionTriggers()` reaps them; run it as **each** account that may have armed
  triggers. `listAllProjectTriggers()` also flags any trigger whose handler is missing, since
  nothing previously surfaced that — which is why this ran broken for two weeks.

## [2026-08-16] — the 2SV setup exemption stops being permanent

### Added — `src/accounts-and-groups/TwoSvSetupGroup.gs` 1.0.0: a nightly prune of the 2SV setup group

Membership of the 2SV setup group is a **security exemption**: it holds a member out of 2SV
enforcement for the window while they enroll. The help-desk app parks them there one at a
time, and nothing ever took them back out. An exemption granted for an afternoon quietly
became permanent — on precisely the accounts that have no second factor, which is why they
were put there.

`pruneTwoSvSetupGroup()` cross-references the group against 2SV enrollment nightly and
removes a member when **either** condition lands, whichever is first:

| Condition | Reason logged |
|---|---|
| enrolled in 2SV | `2sv-enrolled` |
| 7 days elapsed, still not enrolled | `grace-expired` |

Enrollment wins even on the day the member was added — the group's whole purpose has been
served at that point.

**Why a ledger.** `AdminDirectory.Members` carries no join timestamp, so "when was this
member added?" cannot be answered from the directory; the Reports API keeps admin activity
for six months but is a much heavier read for a group meant to hold a handful of people. So
each run stamps a first-seen date into `TwoSvSetupGroupLedger.txt` (CAPWATCH data folder)
for members it has not seen before, and the grace clock runs from there. **Members already
in the group when this first runs therefore get a full window from that first run, not from
whenever they were really added.** That is the safe direction to be wrong in — the
alternative is pulling the exemption out from under someone mid-enrollment — and it
self-corrects after one window. Leaving the group drops the ledger entry, so being parked
again later starts a fresh window rather than inheriting an expired one.

**What it will not do**, each pinned in `test/TwoSvSetupGroup.test.js`: end an exemption on
a 2SV read that *failed* (unreadable state is `null`, never `false` — but the grace clock
still applies, so it cannot keep anyone exempt forever either); touch a nested group or any
non-USER member; or restart the clock for a removal that failed, which would let a failing
removal renew the exemption it was trying to end.

`TENANT_2SV_SETUP_GROUP` (config.gs 1.14.0) is the group address and **blank disables the
module** — this one will not guess at an address. The help-desk app is a separate script
project and reads the same group as `WEBAPP_2SV_SETUP_GROUP`; set both. Run
`previewTwoSvSetupGroup()` once before arming the daily trigger (Admin Guide §8, slot 9) to
see who the first night would take out.

## [2026-08-16] — the help desk stops going through one person

### Added — `admin-webapp/` 1.0.0: a domain admin page for per-member account operations

Wing IT volunteers hold Google's **Help Desk Administrator** role, which covers a password
reset from the Admin console and nothing else this automation does. Resending the welcome
email an out-of-band account never received, parking a member in the 2SV setup group,
finding out *which* of a member's two accounts is the live one — each of those meant asking
the one person with Apps Script editor access to run a function by hand. That is a
bottleneck made of one person, and the editor is not somewhere a help-desk volunteer
belongs: every function in `src/` is one mis-click from a wing-wide run.

`admin-webapp/` is a **fourth standalone script project** (after `src/`, `signature-webapp/`,
and the alias admin UI that was never merged). One `doGet` per project is the mechanical
reason; the useful one is that its OAuth scopes are the four actions' and no more — no
calendar, no chat, no shared contacts, no group *settings*.

Four actions, each scoped to one member: **look up** (CAPWATCH record + every account
carrying that CAPID, with suspension, 2SV, last sign-in, admin flags, managed-group
membership and the welcome-ledger verdict), **reset password**, **resend welcome email**,
**group add/remove**.

**Two gates make it safe to delegate**, and both live in `Auth.gs`:

- *Who may call* — a super admin, or the holder of a directory role in `WEBAPP_ADMIN_ROLES`
  (default `_HELP_DESK_ADMIN_ROLE`). Roles rather than a maintained group of names, because
  the Admin console already records this: a volunteer who loses the role loses the app the
  same day, with nobody remembering to edit anything. A role lookup that *fails* denies.
- *Whom they may act on* — a non-super-admin may not act on **another admin's account**.
  This mirrors Google's own rule, and without it a help-desk volunteer could reset a super
  admin's password here and then sign in as them.

Two further boundaries: `WEBAPP_MANAGED_GROUPS` (plus the 2SV group) is an allow-list
checked against the address the *browser* sends, so the page cannot be used to join any
group in the domain; and credentials are never mailed to an address on our own domains —
not the account being reset, not the secondary domain — which is the rule that makes a
"sent" welcome email mean anything.

Every action writes one audit row naming the admin, the account and the outcome, to the
execution log always and to a sheet tab when `TENANT_AUTOMATION_SPREADSHEET_ID` is set.
The platform's own log names only the DEPLOYER — the app executes as them — so without this
the one question that matters about a delegated tool has no answer. **The temporary password
is never recorded**: it exists in the admin's browser and, if mailed, the member's inbox.

**Four things are copies of `src/`** — the welcome template, the resend eligibility policy,
the authoritative-account ranking, and the welcome-ledger format (which is *shared state*: a
resend from here records into `WelcomeEmailLedger.txt` so the monthly audit stops reporting
that member as MISSED). `test/AdminWebApp.test.js` runs both copies over the same inputs and
fails on divergence; its header records what each drift would cost. Two refusal *sentences*
differ on purpose — `src/` tells an operator to "pass `{force: true}`", which is meaningless
next to a checkbox — and those two are pinned by name so a third divergence cannot arrive
quietly.

Seniors only for now, and **pushed** to project `1miHntSI…Sq6ub` on 2026-08-16 — code only.
Script Properties and the by-hand deployment are still outstanding; see
[docs/ADMIN_WEB_APP.md](docs/ADMIN_WEB_APP.md).

### Added — credentials mail carries 2SV setup steps when the account lacks it

Both mails the help-desk app sends — the welcome email and the password-reset notice — now
include step-by-step 2-Step Verification instructions **when the target account is not
enrolled**. That is the one moment a member can act on it: they have a working password in
hand and the setup flow is three clicks from where they are about to land.

Conditional on the directory’s own `isEnrolledIn2Sv`, not on whether the account looks new.
Telling a member who already has 2SV on to go and turn it on reads as a system that does not
know what it is talking about, and the next real instruction from IT gets skimmed. It matters
for a FORCED welcome resend in particular, which can land on an established account.

The block is INSERTED AT SEND TIME rather than written into `WelcomeEmail.html`, which is a
byte-for-byte copy of the template `src/` sends and is pinned by the parity test. Editing it
would either break that pin or force the same conditional into provisioning, which only ever
mails accounts that cannot have 2SV yet. The template’s existing one-line mention stays in
both cases — this adds the how, not the whether. A template missing its footer marker gets
the block appended lower down rather than dropped: the password is live by then, and a
formatting problem must never become a failed send.

Also fixes the structural test that checks every client-reachable `api*` function is behind
`requireAdmin_()`. It matched a literal newline-brace, so on a CRLF checkout it tested an
empty string and reported all six entry points as ungated — failing loudly for a reason
unrelated to the property it guards.

### Changed — the 2SV group panel follows the member's 2SV state

Reported from live use: an admin saw **2SV off** on the account card and had nothing obvious
to do about it — the group panel was a bare Add/Remove pair with no connection to the fact
sitting directly above it.

The 2SV setup group is not a general-purpose group and no longer looks like one. It covers
the window while a member enrolls, so at any moment there is a right answer, and both facts
that decide it are already on the page. The panel now names the state and makes the matching
action primary: *2SV off and not a member* → **Add** ("add them while they enroll");
*2SV on and still a member* → **Remove** ("they no longer need the setup group"); the other
two states recommend nothing. Both buttons remain — an admin may have a reason the page
cannot know — but the no-op one is disabled rather than silently succeeding. The server side
was already idempotent in both directions and stays that way.

Group membership is now resolved **per account** rather than once for the authoritative one.
A duplicate pair can differ in both 2SV state and membership, and the previous build showed
the authoritative account's membership under whichever radio button was selected, with a
"membership shown for …" caveat papering over it. The caveat is gone because the gap is.

With no group configured the section now **explains itself instead of vanishing**. Everyone
reading this page is an administrator, and a section that silently disappears is
indistinguishable from a broken page — which is exactly how this was first reported.

### Added — cadets in a name search are flagged and pointed at the cadet tools site

CAPWATCH is scoped to the **wing**, not the tenant, so the seniors extract lists every cadet
in the wing and a name search returns all of them — none with an account here to act on.
Reported from the first real use of the page.

They are flagged rather than filtered: an admin who searched a name and got nothing back
would conclude the member does not exist, and one shown "no Workspace account" would
conclude it needs provisioning and go create one **on the wrong tenant**. Candidates now
carry an *other Workspace* tag with a count beneath the list; opening one replaces the
"no account" message with a note naming `WEBAPP_CADET_TOOLS_URL` as a link, repeated inside
the member card because a banner is replaced by the next action's result and this fact is not.

The rule is derived from `TENANT_PROFILE` against a copy of `MEMBER_TYPES_ACTIVE`, not
hardcoded to "cadet": **the region profile provisions cadets too**, so hardcoding would have
broken that tenant the day it deployed. The test compares the table against `src/config.gs`.
An unrecognised type reads as *not* ours — the worst case is a redundant note above a member
whose accounts are listed right below it, rather than a help desk quietly failing on someone.
A member who *does* hold an account here despite the flag (a cadet-to-senior transition in
flight) keeps the full action set, with the note saying so.

One naming trap worth recording: `Credentials.gs` holds the welcome email and would rather be
called `WelcomeEmail.gs`. It cannot — Apps Script addresses files without their extension, so
that name collides with `WelcomeEmail.html` and the push is rejected wholesale with *"A file
with this name already exists"*.

## [2026-08-03] — sweep phase added to the transition pipeline

### Added

- **`CadetTransition.gs` (v1.2.1)** — **`catchUpTransitionMail` is now a pipeline phase.** The
  close parks accounts rather than deleting them (see below), so a parked account stays live
  and keeps receiving mail; this sweep is what actually carries it to the senior tenant, and
  without it a parked mailbox silently accumulates mail nobody reads. Added a guard that
  `TRANSITION_TRIGGER_FUNCTIONS_` (the phase order) and the pipeline's phases map agree — they
  are separate declarations, so an edit to one can orphan a phase that then silently no-ops.

- **`listAllProjectTriggers()`** and **`clearAllTransitionContinuations()`** — Apps Script caps
  a script at **20 triggers** across the whole project, and nothing surfaced that shared
  budget; arming failed outright with *"This script has too many triggers."* These show the
  usage, flag orphaned continuation triggers, and reclaim them.

## [2026-08-03] — the close parks the account instead of deleting it

### Changed

- **`CadetTransitionCleanup.gs` (v2.0.0)** — **the close no longer deletes the cadet
  account.** Two live attempts proved the original design impossible: a Google Group can only
  take an address once Google releases it, and Google reserves a former primary address well
  beyond a single execution — ~20 days after a delete (recovery tombstone) and, measured, >75s
  after a rename (docs say up to 24h). Both returned `409 Entity already exists`. Any
  delete-now design therefore guarantees a bounce window of a day to weeks.

  The account is now **parked**: kept live, forwarding to the senior address, so the address
  never stops existing and no mail ever bounces. The trade is a license seat held until the
  forwarding window ends, at which point `expireParkedAccounts()` runs a final catch-up sweep
  and deletes the account.

  Delivery is guaranteed by the **daily catch-up sweep** (`catchUpTransitionMail`, now on the
  trigger schedule at 09:00), which moves anything arriving at a parked mailbox across to the
  senior tenant. Gmail auto-forwarding is *also* requested but is best-effort only:
  cross-domain forwarding requires the member to click a confirmation Google sends to their
  senior address, so it reports `pending` until they do — it makes delivery instant rather
  than sweep-delayed, but never gates it. Needs `gmail.settings.basic` +
  `gmail.settings.sharing` on the cadets local SA; without them it degrades to sweep-only.

  Removed the group/rename machinery (`freeAddressAndForward_`, `createForwardingGroupWhenFree_`,
  the alias-retention handling and their tests) — that whole approach cannot work, and leaving
  it invited its reintroduction.

### Fixed

- **`whyNotCloseable_`** now refuses a row whose previous close deleted the account but left no
  forwarding, routing it to `repairFailedTransitionCloses()` instead. Such rows previously
  passed every check, so re-running the close would have swept against a deleted mailbox.

## [2026-08-03] — free the address by renaming, not deleting (close incident fix)

### Fixed

- **`CadetTransitionCleanup.gs` (v1.3.0)** — the first live `closeCompletedTransitions(false)`
  run deleted four cadet accounts and then **failed to create any forwarding group**:
  `directory.groups.insert` returned `409 Entity already exists` for every one. A deleted
  Workspace account **holds its address for Google's ~20-day recovery window**, so the
  delete-then-create-at-the-same-address sequence could never work. (No member data was lost —
  migration and the final catch-up sweep both completed before each delete, moving 1–15 late
  messages each; the deleted accounts remained restorable.)

  The address is now freed by **renaming the user off it**: rename → remove the alias the
  rename auto-retains → create the group (retrying the brief, genuinely transient post-rename
  409) → delete the renamed shell, whose tombstone holds the `.transitioned` address rather
  than the real one. The **delete is now LAST**, so any earlier failure leaves a live,
  recoverable account instead of a deleted one.

  Added **`repairFailedTransitionCloses(dryRun)`** (+ `previewTransitionRepair()`) to finish
  the stranded rows once their accounts are restored from the Admin console — it rebuilds the
  forward without re-migrating anything. Added **`testAddressHandover(addr, forwardTo)`**,
  which exercises the whole sequence on a throwaway *user*; the earlier
  `testForwardingGroup()` validated only a group at a *fresh* address, which is exactly why
  this reached production. Removed the now-unused plain create helper so the unsafe order
  cannot be reintroduced.

## [2026-07-31] — two spellings, one mailbox

### Fixed — UpdateGroups.gs 1.11.0: membership compared on account identity, not string equality

`googleAccountKey()` (utils.gs) folds gmail dots and `+tags`, because two strings can be one
mailbox and Google resolves them while `===` does not. Its docstring records the cost: a group
holding one spelling while CAPWATCH supplies another reads as **a member to ADD and a stranger to
REMOVE** — the add returns 409 and is swallowed, the remove succeeds, and the member is off the
list until the next run. Silently, once per address change.

**`SquadronGroups` has compared this way since it started nesting external groups
(`diffGroupMembership_`). This pass never adopted it** — zero references. That was inert for as long
as these lists held only `@cawgcap.org` addresses: every one of those is already its own key.

It stopped being inert the moment `Add EXT` went on the wing `all` row. That run added **18 CAPWATCH
personal addresses** — seniors who have not completed Level I and so hold no Workspace account —
to `ca.all` and the unit lists, most of them gmail. Every one of them is now exposed to the gap the
helper was written for.

The comparison moves out of the delta loop into `reconcileCurrentAgainstDesired_()`, which is pure
and covered by `test/UpdateGroups.accountIdentity.test.js` — loading the *real* `googleAccountKey`
from utils.gs, so the assertions break if the folding rules ever change. An address the group
already holds under another spelling of one account is now neither added again nor removed.

Two behaviors deliberately unchanged: folding stops at gmail/googlemail (dots are significant
elsewhere, and folding them would merge two real people), and `MANAGER`/`OWNER` entries are still
never auto-removed.

**No action needed for the accountless seniors.** `buildWorkspaceEmailMapForGroups_` rewrites a
member's address to their Workspace one the moment the account exists, so the day one of them is
provisioned the personal address drops out of the desired set and the delta removes it — a different
domain is a different key, so this change does not interfere. Asserted in section 4 of the test.

## [2026-07-31] — why the wing-wide all list reaches no cadet, in one function

Members reported that `ca.all@cawgcap.org` does not deliver to cadets while the unit `.all` lists
do. That is not a contradiction — the two are reached by different code:

| List | Managed by | Nests the cadet group? |
| --- | --- | --- |
| `ca###.all` | `updateAllSquadronGroups()` (SquadronGroups.gs) | yes, and reconciles the settings that permit it |
| `ca.all` | `updateEmailGroups()` (UpdateGroups.gs) **only** | only if the sheet says so |

`shouldCreateDistributionLists()` returns false for anything that is not UNIT scope, and no CAPWATCH
org is wing scope, so **SquadronGroups never touches the wing list**. Every prerequisite it handles
for a unit list has to be satisfied by the Groups / User Additions tabs for the wing one, and each
fails silently and separately.

### Added — `groupAdministration_diagnoseWingAllFanout(groupEmail)` (groupAdministration.gs)

Read-only, run on the WING tenant. It checks the four things this tenant can see and names the one
that is failing, rather than leaving four candidates to be eliminated by hand:

1. **Is a cadet-tenant address a member at all** — with role, status and `delivery_settings`, because
   a member set to `NONE` is listed and still sent nothing.
2. **Does anything preserve it.** `updateEmailGroups()` removes any plain `MEMBER` it did not compute
   as desired, and a cross-tenant address only becomes desired through a **User Additions** row naming
   the group. Without that row the nesting is removed on the next sync, which is what "it does not
   hold" looks like from outside.
3. **Does the Groups row permit external members** (`Add EXT`, or `Add Lite` which implies it).
   Without it the apply loop declines the add *and* `applyManagedGroupSettings_()` writes
   `allowExternalMembers=false` onto the group every run — so fixing the flag in the Admin console
   reverts within a day, and the two failures look like one intermittent one.
4. **The live group settings**, so a false flag is reported as a fact rather than inferred.

It also flags a User Additions row that `updateCAWGCadetGroups()` will delete. That function rewrites
every row whose address matches `ca[###].(cadets|parents|all|…)@cawgcadets.org` but only *generates*
`.cadets@` and `.parents@` rows, so a hand-added `ca.all@cawgcadets.org` row is erased on its next
run and the member disappears one sync later. The address it generates for the wing list is
`ca.cadets@cawgcadets.org`.

Nothing here writes. See [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md).

### Added — `groupAdministration_diagnoseReceiveGroup(groupEmail)` (groupAdministration.gs)

The receiving half, run on the tenant that **owns** the group. A Workspace tenant cannot read the
other's group settings, so three prerequisites are invisible from the sending side and each fails
without a bounce anyone sees: the group **exists** (nesting an address that resolves to nothing is
accepted and delivers to no one), it **has members** (an empty receive list is indistinguishable
from a working one until somebody asks a cadet), and `whoCanPostMessage` is **`ANYONE_CAN_POST`**
(fan-out carries the *original* sender, who is external to the receiving tenant).

`groupAdministration_auditReceiveListPosting()` remains the bulk form of the third check; this is
for a specific address about to be nested, or already nested and not delivering.

### Fixed — UpdateCAWGCadetGroups.gs 1.3.0: a wing target nested an address that does not exist

`buildCAWGCadetSourceGroupEmail_()` took `scope` **and ignored it**, so the wing target resolved to
`ca.cadets@cawgcadets.org`. Nothing creates that address: `.cadets` groups on the cadet tenant come
from `updateAllSquadronGroups()`, which reaches an org only via `shouldCreateDistributionLists()` and
therefore walks UNIT scope only — and no CAPWATCH org is wing scope. It is the same structural gap
that leaves the wing `.all` list unmanaged, one tenant over.

Confirmed against live state on 2026-07-31: `ca.cadets@cawgcadets.org` returns **404**, while
`ca.all@cawgcadets.org` holds **2,646 members**, is `ANYONE_CAN_POST`, and is named "CAWG - Cadets".
On a cadet-only tenant that all-hands list *is* the wing's cadets, so nesting it is not a
substitution — it is the same population under the name that exists.

Three things followed from the old behavior, none of which surfaced an error:

- the generated User Additions row pointed at nothing; the add 404s and is logged as "cannot add
  external member", one line among thousands
- `ca.cadets@cawgcap.org` — the wing-wide cadet list — had no member that resolves, because unit
  cadets are nested into `ca###.cadets`/`ca###.all` and never into `ca.cadets`
- **it deleted the row for the address that works.** `isManagedCAWGCadetGroupEmail_()` matches
  `.all@` while the generator produced none, so a hand-added `ca.all@cawgcadets.org` row was dropped
  as an orphan of the function's own making

A WING-scope cadet source is now `<wing>.all@<cadet domain>`; UNIT scope is unchanged. The generator
now emits the row CAWG had added by hand, so the two stop fighting.

**Parents was left pending a check, which has now been made — see 1.4.0 below.**

### Fixed — UpdateCAWGCadetGroups.gs 1.4.0: the wing parents list is fed by the units

`groupAdministration_diagnoseReceiveGroup('ca.parents@cawgcadets.org')` on the cadets tenant, run
2026-08-02: **404.** Same structural gap as `.cadets` — nothing creates wing-level groups there.

Cadets had a substitute available, because that tenant's wing-wide all-hands *is* its cadets.
**Parents have none**: those groups exist per unit only. So the two take different routes now, and
the code says why in both places:

| | Wing source row | Units reach up? |
| --- | --- | --- |
| Cadets | `ca.all@cawgcadets.org` → `ca.cadets`, `ca.all` | no — the aggregate already covers every cadet |
| Parents | *none emitted* | yes — each `ca###.parents@cawgcadets.org` also nests into `ca.parents` |

The wing parents **source** row is no longer emitted at all. It could only ever point at a
nonexistent address, 404, and be swallowed — which is exactly how `ca.parents@cawgcap.org` came to
hold no member that resolves. `ca.parents` remains a managed **destination**, or the 56 unit groups
would be nested into something this function had stopped managing; that is asserted separately.

Sections 9 and 10 of `test/UpdateCAWGCadetGroups.nesting.test.js` drive the real
`buildCAWGCadetManagedRows_()` against a fake wing/group/unit org tree, so the routing decision is
covered rather than the helpers alone.

### Fixed — UpdateCAWGCadetGroups.gs 1.2.0: it built a dependency it did not satisfy

`updateCAWGCadetGroups()` set `Add EXT` on the `cadets`/`parents` rows it creates, and nested those
groups into `.all` rows written by a human — without setting it there. A blank column on that row is
not inert: `updateEmailGroups()` reads it as "this group holds no outside addresses" and both
declines the add **and** writes `allowExternalMembers=false` onto the group, every run. So the sheet
asked for a nesting it simultaneously forbade, and an Admin-console repair reverted by morning.

Found on `ca.all@cawgcap.org`: a nested cadet group present, `allowExternalMembers=false`, and no way
to re-add the member if it were ever lost. Every Groups row a generated nesting lands in is now
stamped `Add EXT = Y` — one cell; `Category`, `Attribute`, `Values` and `Description` stay as their
author wrote them, since rewriting those would silently redefine who is on the wing's all-hands list.
Both fixes are covered by `test/UpdateCAWGCadetGroups.nesting.test.js`. Neither was ever
CAWG-specific — they apply to any wing adopting the split-tenant pattern.

## [2026-08-03] — the duty CAPWATCH never told us about

A member's signature was missing their Pacific Region billet. Not mis-sorted, not filtered —
**absent from the data**. CAPWATCH scopes an extract to the echelon it is downloaded as, so the
wing pull (ORGID 188) carries the duties its members hold *at wing or below* and simply has no row
for a region or national assignment. Nothing logged, because there was nothing to log; the wing
extract is not aware the billet exists.

Confirmed by downloading the region list by hand and finding the row there, keyed to the same
CAPID.

### Added — `UpdateMembers.gs` 1.22.0, `signature-webapp/` 1.1.0

The region tenant already downloads a region-wide extract, so the wing tenant now reads **that
folder** as a second, read-only source — `TENANT_REGION_CAPWATCH_DATA_FOLDER_ID`, shared read-only
to the automation account. No second download, no additional eServices permission, no duplicated
quota. The tenant goes on operating as the wing in every other respect; the supplement answers one
question only: *what duties does this member hold that our own pull cannot see?*

**Two invariants, both pinned by tests, because this source is also the one that could do the most
damage:**

1. **It never widens the roster.** Rows are read only for CAPIDs the tenant's own extract already
   has. The region extract holds every member of every wing in PCR; treating it as a source of
   members would provision Nevada.
2. **Nothing that already reads a member can see them.** They land in their own array,
   `member.outOfWingDutyPositions`, merged only by `allSignatureDuties_()` inside the signature
   generator.

   This shipped wrong first. It originally pushed into `dutyPositions` while staying out of
   `dutyPositionIds`/`dutyPositionIdsAndLevel`, on the belief that those two arrays were the
   group-matching contract. **They are one contract, not the only one.** `UpdateGroups.gs` iterates
   `dutyPositions` directly to match duty ids and levels, `SquadronGroups.gs` does the same for unit
   distribution lists, and `addOrUpdateUser()` builds the Workspace **directory job title** from it.
   A member's region billet would have appeared in their GAL title and in any wing duty group whose
   title happened to match — caught by grepping the consumers *after* the property went live on the
   main project, and before any sync ran. The separate array makes the invariant structural rather
   than something ten call sites have to remember.

Only orgs **outside** this wing are taken; our own pull is authoritative for ours. Filtering on the
org rather than de-duplicating keeps the rule legible — "duties our own pull cannot see" — instead
of depending on which extract downloaded last. Org names come from the supplement's own
`Organization.txt`, the only place "Pacific Region" and "National Headquarters" exist, so a region
billet renders as *"Pacific Region Assistant Director of IT"* rather than falling back to the
member's home unit. `DUTY_LEVEL_ORDER` already ranked NAT/NHQ above REGION above WING.

Self-disabling throughout: unset property, missing folder, unreadable folder and empty file all
degrade to "no extra duties", never to an error. A member still gets their signature.

### Added — `previewOutOfWingDuties()`, because the first failure was undiagnosable

When it did not work, the log said nothing — "the supplement is switched off" and "the supplement
found nothing" were indistinguishable from outside. The normal path now logs whether the supplement
was consulted and what it found, and `previewOutOfWingDuties()` walks the same path on demand and
reports where it stops: property unset, folder unreadable, file absent, no rows for that CAPID, or
rows found and for each whether it was TAKEN or SKIPPED as ours.

Run it **from the editor**, which executes the project's current code — the deployed app serves
whatever version was last published, and "pushed but not re-deployed" is a common cause that this
distinction isolates immediately.

> ⚠️ **The share is a manual, cross-tenant grant.** Folder `1lU9…` lives in the PCR Workspace and is
> shared read-only with the wing's automation account. If that lapses, or the PCR pull stops being
> armed, out-of-wing duties quietly stop appearing — the failure mode is silence by design, which is
> why the log line and the diagnostic exist.

> Two things that look like failure and are not: an **assistant** billet is never in the default
> pick, so it sits unticked in the chooser rather than appearing in the preview (most region
> assignments are assistant ones); and the resolved record is **cached for ten minutes**, so a page
> load from before the property was set keeps serving the old one.

`src/` carries the change but has **not been pushed to any tenant**, so provisioning and
`pushAllSignatures()` do not yet see region duties. The web app is live on seniors.

## [2026-07-28] — members can set their own signature, and change exactly one thing about it

A CAP email signature reached a member three ways: five minutes after their account was created,
whenever somebody ran `pushAllSignatures()` by hand over the whole wing, or by hand-building one
in CAP's own generator. Nothing covered a member who joined before signatures were automated, was
promoted, picked up a new duty assignment, or simply cleared their signature — for those, the
answer was a ticket to IT or an afternoon with the brand-tools page.

### Added — `signature-webapp/` 1.0.0 (a NEW, separate Apps Script project)

A member opens a link, signs in as themselves, and sees the signature their CAPWATCH record calls
for. If they approve it, it is written to their own CAP addresses.

**One thing is theirs to decide: whether the phone row appears.** Grade, name, duty assignments and
unit come from CAPWATCH and the layout comes from the CAP brand style guide, so the browser sends
exactly one value — a boolean — and the server rebuilds the entire signature from CAPWATCH on
every call, **including on apply**. The preview is a rendering of the record, never a document the
client can hand back edited. The page says as much: a wrong grade is fixed in eServices, and
arrives with the next extract.

Two consequences of following `src/`'s own rules rather than inventing new ones: a **cadet never
has a phone row at all** (`addContactInfo()` refuses to publish cadet numbers to the directory or a
signature), and neither does a senior with no cell in eServices. Both cases replace the toggle with
a sentence explaining why there is nothing to toggle.

**Why a separate script project.** An Apps Script project has one `doGet`, and `src/` already
spends its on the FileMaker mission webhook (anonymous, runs-as-deployer) while `webapp/` spends
its on the alias ADMIN interface. This is the opposite of that one in who may use it — every
member of the domain — so it gets its own project, manifest and scope list rather than sharing a
deployment with alias-mutation code.

**Access model.** `access: DOMAIN` + `executeAs: USER_DEPLOYING`: the CAPWATCH read and the Gmail
write need rights a member does not have and should not be given. `DOMAIN` is load-bearing, not a
preference — under anonymous access `getActiveUser()` is blank and every visitor is
indistinguishable. There is deliberately **no admin-acts-for-a-member mode**; the account acted on
is derived from the session and nothing else, so no request can reach another member's mailbox.
`SIGNATURE_WEBAPP_ALLOWED_GROUP` restricts the app to a pilot group; blank means every member,
which is the reverse of `webapp/`'s fail-closed default and deliberate — that app hands out
addresses, this one lets people format their own name.

**The write is bounded twice.** Only the caller's own mailbox, and within it only Send-As
identities on a domain the tenant owns. Members add their personal accounts as Send-As identities,
and stamping a CAP signature on someone's private mail would be a real intrusion. The domain check
compares the whole domain, so `…@cawgcap.org.example.com` is refused rather than matched as a
suffix. A mailbox with no org-owned identity gets an error, not a success that changed nothing.

**The cost, stated plainly.** The signature template is *duplicated* from
`src/accounts-and-groups/UpdateMembers.gs`, because two script projects cannot share code without
becoming a library with its own deploy-version step. A copy that drifts is worse than no copy:
a member would approve one signature and have the next bulk push replace it with a subtly
different one. So `test/SignatureWebApp.test.js` renders eight fixture members through **both**
generators and fails on any byte of difference — and does the same for the record builder, running
`src/`'s own `createMemberObject` / `addContactInfo` / `addDutyPositions` against this project's
single-CAPID reader over one set of CAPWATCH fixtures. 100 assertions, including a cadets-tenant
section over cadet, cadet-lite and cadet-duty records, and one that loads the real `Config.gs` over
a properties store where an unset key is absent rather than empty.

Impersonation needs `SA_IMPERSONATION_EMAIL` / `SA_PRIVATE_KEY` in **this project's** Script
Properties as well: Gmail settings have no admin-on-behalf-of endpoint, and Script Properties are
per project. That is a second copy of a credential that can act as any user, so a key rotation must
now update both projects — recorded in the Admin Guide's standing rotation item.

### Fixed — an assistant duty claimed the billet (`UpdateMembers.gs` 1.21.1, web app 1.1.1)

Reported from a real record within minutes of the duty chooser going live: picking the assistant
Supply Officer printed **"California Wing Supply Officer"**. Not a formatting slip — the signature
stated that the member held a billet somebody else holds.

CAPWATCH keeps the assistant flag in its **own column**; `DutyPosition.txt`'s `Duty` value reads
"Supply Officer" for the officer and their assistant alike. The signature printed that value bare,
which was harmless for exactly as long as assistants were filtered out of every signature — and
wrong from the moment a member could choose one. Shipped inside 1.21.0, found in the first hour of
use.

`signatureDutyTitle_()` now prefixes "Assistant " for an assistant duty, skipping titles that
already carry the word.

**Deliberately NOT in `formatDutyTitle_()`.** That function normalizes a title so `UpdateGroups.gs`
and `SendRetentionEmail.gs` can MATCH it against a configured list; an "Assistant " prefix there
would quietly change who lands in a duty group. Display and matching want different strings, and
this is the display one.

Default signatures are unaffected — they exclude assistants entirely — so nothing the automated
path produces changes. Four assertions had encoded the bug and now encode the fix, alongside new
ones for the no-double-prefix case, a renamed title keeping its rename under the prefix, and a
principal never gaining the word. The page also drops its "assistant" badge: the line now says it.

### Added — `Index.html` 1.1.1: the signature's NAME in Gmail, explained rather than attempted

Asked for as an option on the page. It cannot be one: Gmail's multiple-signatures feature labels
each signature, and **the Gmail API does not expose that label**. The `SendAs` resource carries nine
fields — `sendAsEmail`, `displayName`, `replyToAddress`, `signature`, `isPrimary`, `isDefault`,
`treatAsAlias`, `smtpMsa`, `verificationStatus` — and `signature` is the HTML, nothing else. No API
client can set the name, so no code here could have.

What the page can do is stop a member hunting for a control that could not exist. After a
successful apply it now says the name is not something this page can set, and points at the only
place it can be changed — **Gmail → Settings → See all settings → General → Signature** — adding
that the name is shown only to them and never appears on mail they send. Shown only once something
was actually written: an apply that failed must not send someone off to rename a signature that
isn't there.

Not to be confused with `displayName`, the "From:" name, which *is* settable and is already set at
provisioning by `updateGmailSendAsDisplayName()`. This app still does not touch it.

### Added — members choose which duty assignments appear (`UpdateMembers.gs` 1.21.0, web app 1.1.0)

From a tester: give members a way to pick which duties the signature shows, keeping the two-position
limit and the echelon hierarchy.

The block previously guessed: assistants excluded, at most one duty per echelon, most senior title
first. Those rules exist to choose well **in the member's absence** — and on this page the member is
not absent. So when a selection is present, the defaults give way to it; **what never gives way is
the ordering or the cap.**

`getDutyBlock()` gained an optional `member.selectedDutyKeys`. The web app filters the member's own
duty list and hands it to the **same generator** every other signature goes through, so:

- **Highest organizational level still comes first.** A wing *assistant* precedes a squadron
  *principal* — the new principal-before-assistant rule orders duties **within** one echelon, not
  across them, because the style guide asks for highest level first and that is not a member's to
  reorder. Ticking order is irrelevant.
- **Two lines, still.** Enforced in the generator, again in the API, and shown in the page by
  greying the remaining boxes once two are ticked.
- **Assistants are selectable but never default.** Untouched, the page produces exactly what the
  automated path produces.
- **Every key is checked against the member's own CAPWATCH record** and a stranger's duty is
  *refused, not dropped* — quietly ignoring it would publish a signature the member never saw.
  Three duties are refused, not truncated, for the same reason.

**Nothing in `src/` sets the new field.** Provisioning and `pushAllSignatures()` pass a member
straight from `getMembers()`, so their output is byte-identical to before it existed — asserted,
along with the port's agreement, in the test suite.

The choice is session-only, like the phone: reopening the page starts from the CAPWATCH default
again, and a bulk push resets it.

Also drops the duty row from the "What CAPWATCH says" card, which now duplicated the chooser
directly beneath it.

### Fixed — the signature's vertical rhythm (`UpdateMembers.gs` 1.20.1, `SignatureTemplate.gs` 1.0.1)

Two stray blank lines, both visible only once a real signature sat in a real mail client:

- a **leading `<br />`** before the name line, which Gmail rendered as a blank line at the top of
  every signature this project has ever pushed;
- the logo's **20px bottom margin**, which opened a gap above the tagline that no other row in the
  block has — every other margin is `0 0 5px`. Now `15px 0 5px 0`: the space above the logo is
  kept, the one below it matches everything else.

**Changed in `src/` first and mirrored into the web app's port**, which is the rule the parity test
exists to enforce — `test/SignatureWebApp.test.js` compares the two generators byte for byte and
would have failed on a one-sided fix. Three new assertions pin the spacing itself, so neither can
creep back: nothing precedes the name line, the logo carries the 5px margin, and `0 0 5px` /
`15px 0 5px 0` are the *only* margins in the block.

⚠️ The `src/` change reaches members' signatures only when a tenant is next pushed **and**
`pushAllSignatures()` or an account provisioning runs. Until then the two paths agree in the repo
and disagree on the tenants — the web app (already pushed) writes the corrected spacing, the
tenants' own copy still writes the old.

### Changed — `Index.html` 1.0.2 / `SignatureApi.gs` 1.0.1: the format has an authority, and it is not us

The page now names **CAP's own signature generator**
(`cap-brand-tools.netlify.app/signature-generator`) in a footer, and — more usefully — on both
pages that *refuse*: not signed in to a wing account, and tenant not configured yet. A member
turned away by either had otherwise been told only what did not work. The generator gives them a
correct signature by hand whatever is wrong with their account or this tenant's setup, and it
answers "why does my signature look like this?" with something better than "because the wing's
script says so". It is the same site the signature's logo is served from, and the URL is defined
once, server-side, in `Config.gs`.

Also dropped "Nobody else's" from the header copy.

### Fixed — `Index.html` 1.0.1, first live test: a ticked box that sent `false`

The first real apply dropped the phone row from a signature the member had **ticked the box for**,
and then said so: *"Signature applied to … The phone row was left out."*

The server was right and honest — the browser lied to it. `apply()` called `setBusy(true)`, which
**disables the checkbox** to show a request is in flight, and `includePhone()` read
`!box.disabled` as part of deciding its answer. By the time the value was computed the box was
always disabled, so the answer was always `false`. `refreshPreview()` had the same trap, which is
why re-ticking never brought the phone back in the preview either — the control looked broken as
well as acting wrongly.

Two changes, either of which would have been enough, kept because they say different things: the
value is read **before** the UI starts mutating itself for the request, and "is this control
currently disabled" is no longer treated as an opinion about what the member wants. `STATE.phone`
is the only gate that belongs there.

**Everything on the server side was already correct and tested to the point of pedantry** — only
an explicit `true` includes the phone, every other value suppresses it, the render is rebuilt from
CAPWATCH on apply. None of it could help: the unit that was wrong was the wire between the checkbox
and the call, and nothing tested that. New `test/SignatureWebApp.page.test.js` loads the real
`<script>` out of `Index.html` over a small fake DOM and asserts what actually reaches
`google.script.run` — 19 assertions, and it fails on the original code.

**One project per tenant, same source.** A script project lives inside one Workspace, so a cadet
cannot sign in to one on the senior domain: seniors and cadets each get their own project, clasp
target (`signature-webapp-{seniors,cadets}.clasp.json`) and `/exec` URL. Nothing in the code
branches on `TENANT_PROFILE` — what differs for a cadet differs because their CAPWATCH record does:
**never a phone row** (`addContactInfo()`'s rule, so a cadet with a clean `CELL PHONE`/`PRIMARY`
row still resolves to none), duties out of `CadetDutyPositions.txt`, and no secondary domain, so a
cadet's senior-domain Send-As is left to the seniors deployment. One deliberate divergence:
cadet-lite grades are filtered out of `getMembers()` because they get no *account*, but this app is
reached *by* a signed-in account — if such a member has one, they get a correct signature rather
than a refusal.

**An unset Script Property is ABSENT, not blank** — the Apps Script UI will not store an empty
value, so "leave blank" means the property is never created and `getProperty()` returns null. The
config reader already treated null and blank alike, which is right for the optional ones (no
allowed-group = every member; no abbreviation = derive `CAWG`; no support address = "ask your wing
IT director"). It was **not** right for a missing required one: with `TENANT_EMAIL_DOMAIN` unset,
`isOnATenantDomain_()` matches nothing and every member is turned away with a message about *their*
account — a configuration gap wearing the costume of an authorization failure. `sigMissingConfig_()`
now runs before identity in both entry points, refuses with "this page is not set up yet", and names
the missing property **in the execution log rather than to the member**. Missing service-account
credentials count as missing config too, so the page never offers a button it cannot honor.

The seniors script project exists (`1iNlP98…`) and the code is **pushed** to it; the cadets one is
still to be created and its ID filled into its clasp target.

⚠️ **Deploy from the Apps Script editor, never with clasp.** `clasp update-deployment` /
`create-deployment` leave the app serving Google's "You need access" page to visitors, and only an
editor redeploy clears it — observed twice on seniors. The Apps Script API replaces a deployment's
config wholesale and the web-app entry point does not survive it; the manifest's `webapp` block is
only a default applied when a deployment is *created*, so the access setting chosen in the editor is
not restored from it. `clasp push` is fine and is how code reaches the project — the deploy step is
a human one.

The `push:signature:*` scripts carry `--force` on purpose. clasp asks before overwriting a
manifest, and Apps Script normalizes `appsscript.json` server-side, so it sees a difference on
*every* push of this project. Run non-interactively it answers its own prompt: prints
`Skipping push.`, exits 0, changes nothing — which reads like success and leaves you testing code
you never uploaded. It cost one round of exactly that here. Setup, per-tenant Script Properties, the
cadet notes and troubleshooting are in [docs/SIGNATURE_WEB_APP.md](docs/SIGNATURE_WEB_APP.md).

> **Known interaction.** `pushAllSignatures()` rebuilds every member's signature *with* the phone
> row, so running it wing-wide undoes anyone's choice to omit theirs (they can untick it again). It
> cannot produce a different-looking signature — that is what the parity test is for. Making the
> opt-out survive a bulk push would need a store both projects can read, and belongs in `src/`.

## [2026-07-27] — two groups that could never be created, and nobody could tell

`groups.insert` refuses a name over **73 characters** outright — 400 "Invalid Input: groupName"
— and creates nothing. On the CAWG senior tenant, `ca404.deputy-commander-cadets` and
`ca404.deputy-commander-seniors` had therefore never existed.

**It stayed invisible because the name is only sent on CREATE.** Sixty-six other units looked
healthy for the simple reason that their groups already existed, so the over-long name was never
re-submitted. The two that failed were the only ones in the whole run that needed creating.

The offending name came from the **abbreviation** path, which is supposed to shorten:

| | |
|---|---|
| Unit name | `California Aerospace Composite Squadron` |
| Matches a shortening pattern? | no — the patterns want a trailing number or a leading ordinal |
| Falls through to | `Sqdn 404 California Aerospace Composite Squadron` — squadron, twice |
| With the label | **78 characters.** Refused. |

Measured across the wing, those two names were the only ones over the limit; the next longest
was 70.

### Changed — `SquadronGroups.gs` 1.9.0

`fitGroupName_()` joins the unit and the label within the limit. **The label never gives way** —
it is what tells one of a unit's lists from another, and two groups whose names differ only past
the cut are the same name to anyone reading the console. The unit is trimmed at a **word
boundary**, and if the cut would leave only a fragment of a word the unit is dropped entirely
rather than shipping `W - Deputy Commander for Cadets`.

**Trimming happens only when the limit is exceeded, so no existing group is renamed** — the
70-character `ca404.parents` name is untouched. For ca404 the trim removes the duplicated
"Squadron" and lands at 69: `Sqdn 404 California Aerospace Composite - Deputy Commander for Cadets`.

### Added — `test/SquadronGroups.groupName.test.js`

25 assertions. The last one is the point of the exercise: across every combination of unit and
label, including absurd ones, nothing the function returns can be refused for length. Caught a
real defect during development — the word-boundary rule was not applied when the cut left a
single mid-word character. All unit names synthetic.
## [2026-07-27] — two spellings, one mailbox, one member quietly dropped

`first.last@gmail.com` and `firstlast@gmail.com` are the same Google account. Dots carry no
meaning on gmail.com and everything after a `+` is a tag. Google resolves all of it. String
equality does not.

So a group holding one spelling while CAPWATCH supplied another looked, to the diff, like a
member to **add** and a stranger to **remove**. The add came back 409 "Member already exists"
and was swallowed. The remove succeeded. **The member was dropped from their unit's list and
only restored on the next run** — a day off the list, once per address change, with nothing in
the log naming who it happened to.

Found on the cadet tenant while chasing something else, and it has two shapes:

| | What happens |
|---|---|
| Group holds one spelling, CAPWATCH sends the other | member removed, absent until the next run |
| Both spellings in the desired set | one added, the other 409s **every run, forever** |

### Changed — `utils.gs`

`googleAccountKey()` reduces an address to the account it identifies. **Only gmail.com and
googlemail.com are folded** — dots are significant everywhere else, and treating
`a.member@example.org` and `amember@example.org` as one person would delete a real member. The
result is a comparison key, never a deliverable address. Input that is not an address comes back
unchanged rather than collapsing onto a shared key.

`executeWithRetry()` gains an optional `quietCodes` list: statuses the **caller** handles as an
expected outcome are still thrown, but no longer logged as errors on the way out. A 409 from a
membership add means "already a member", which the caller treats as success — logging it at
ERROR while the summary reported `errors: 0` taught readers that ERROR lines here are noise.

### Changed — `SquadronGroups.gs` 1.8.0

`diffGroupMembership_()` decides adds and removes, keyed on account identity, on both sides.
Pure — no API calls, no logging — so the decisions are testable without Google. Add-before-remove
ordering is unchanged; it is the safer order when the input data is suspect, and identity
matching removes the need to reorder.

A 409 now records **which member**, at INFO. Both membership bugs found this month were
invisible in the log because this branch swallowed the address along with the error.

### Added — a worklist for addresses Google refuses

A 404 from `members.insert` on a gmail.com address means Google checked its own domain and found
no such account: a typo or closed account in CAPWATCH. Code cannot fix those, and one ERROR line
per failure buried in a run of thousands is not something anyone acts on. They are now also
collected per execution and reported once, as a short list a unit can take to eServices —
**alongside** the per-failure ERROR lines, not instead of them; a failed add is a real error and
carries context the rollup does not. Grouped by address rather than per occurrence, because one
bad contact on a cadet's record reaches every list that cadet belongs to, and reading the same
address three times invites someone to fix it once and assume they are done.

Both refusals Google issues are collected, because they are one situation for the wing:

| Code | Meaning | Reported as |
|---|---|---|
| 404 | Google checked gmail.com and found no such account | `no such account` |
| 400 | Google would not parse the address at all | `malformed` |

The first pass recorded only the 404s, and the live run promptly turned up a **double dot in a
domain** — which `sanitizeEmail()`'s format check accepts, since `[^\s@]+\.[^\s@]+` is satisfied
by `icloud..com`. That address failed every run and appeared in no worklist. Thirteen addresses
failed; twelve were reported.

First live run on cadets: **12 occurrences across 11 addresses** for the 404s alone — one
plus-addressed, one with underscores in a Gmail username (both structurally impossible), the rest
well-formed but not resolving.

`sanitizeEmail()` is deliberately left alone. Rejecting consecutive dots there would be correct,
but it is a shared validator used by member provisioning and every mail path, and the worklist
already gets the address in front of the people who can fix it. Worth doing separately, on its
own terms.

### Fixed — a run that could not add twelve members still reported `errors: 0`

`updateGroupMembership()` has always counted failures into `result.failed`, and **nothing has
ever read that field.** The run summary's `errors` counts squadrons that threw; a member add
failing underneath a squadron that otherwise succeeded is not one of those. So the first run of
the worklist above sat next to `errors: 0` — the same shape of dishonest log that hid the starved
squadron tail for two weeks. `memberFailures` is now carried in the summary beside `errors`,
because they are different numbers and only one of them was being told.

### Added — `test/GroupMembership.identity.test.js`

38 assertions, weighted toward **not merging people who merely look alike** as much as matching
the ones who are the same. All addresses synthetic.
## [2026-07-27] — a run that stops is not a run that finished

`updateAllSquadronGroups()` had timeout protection and no memory. It gave up gracefully when
it ran out of time — and recorded nothing about where it got to, so the next execution
started at the top of the same list and gave up in the same place.

The units past that line were never reached. **Not "reached late" — never, on any run.**

On the CAWG cadet tenant that was the last **9 of 68** squadrons. Their `.all` lists still
carried `allowExternalMembers: false` two weeks after the sync began enforcing `true`, which
is how they were found: the setting is a fossil recording that no run ever arrived. Their
Cadet Lite members — added by personal email, since those cadets hold no account — had never
been added either. Nothing raised an error the entire time, because a run that stops early
still reports success for the part it did.

Confirmed rather than assumed: the nine were exactly the **last nine in iteration order**.

### Changed — `SquadronGroups.gs` 1.7.0

`updateAllSquadronGroups(options)` accepts `{deadlineMs, resume}` and returns its position and
a `complete` flag. **No-argument behavior is unchanged.** New:

```
updateAllSquadronGroupsBatch()        // 25-minute slice; point the daily trigger here
updateAllSquadronGroupsBatch(5)       // shorter slice, e.g. the 6-minute tier
checkSquadronGroupsBatchStatus()      // read-only
resetSquadronGroupsBatchProgress()    // discard the parked run
```

Only the **position** is parked, not the computed rosters — those are rebuilt from CAPWATCH
each slice. Cheap next to the API calls, and a resumed slice then acts on today's data instead
of replaying a snapshot from before the pause.

**It shares `SQUADRON_BATCH_INDEX` with the batch driver that already existed.** This nearly
shipped as a second, parallel mechanism with its own state file. `updateSquadronGroupsBatch()`
has been slicing the same list by count — 10 squadrons a run, a week to come round on a 68-unit
wing — and parking its position in that Script Property all along. Two entry points walking one
list with two private cursors is the original bug wearing a hat: a daily trigger on one and a
manual run of the other would each advance their own, and units would fall between. One pointer,
two paces (count or time), and mixing them is safe. The companion keys (`_CHARTER`, `_SAVED_AT`)
are new; a legacy bare index is still honoured rather than discarded.

**So the operative fix is a trigger change, not just code.** The daily trigger points at
`updateAllSquadronGroups()`, which is exactly the entry point that cannot resume. The Admin
Guide now says so at the trigger table.

**Progress counts are cumulative, and that is not cosmetic.** First live test parked correctly
but logged `resumingAt: "18/0"` and a per-slice `processedSquadrons`, because neither the total
nor the running count was persisted. A progress line that resets every slice looks exactly like
a run starting over — which is the appearance the starved tail hid behind for two weeks. The
total, the running count and the original start time are now parked alongside the position, and
a legacy position with no total prints `?` rather than `0`.

**The fix has the same failure mode as the bug, and that is the interesting part.** A parked
index is a position in an order, and the order comes from CAPWATCH; one unit chartering or
folding shifts every index after it, and resuming on the stale number would skip squadrons
silently — precisely what this exists to stop. So the charter sitting at that index is parked
alongside it, and any disagreement means start over. Re-walking squadrons is idempotent and
cheap; skipping them is the whole bug. `resolveSquadronResumePosition_()` never returns a
mid-list guess — every rejection lands on 0.

### Changed — `config.gs` (comment only)

`MAX_EXECUTION_TIME_MS` is 1,750,000 ms — **29.2 minutes**, not the "5.5 minutes" the comment
above it claimed. That text described the 6-minute consumer tier and had been wrong for a long
time. The real number leaves 48 seconds before a hard kill, and the elapsed check runs only
between squadrons, so one slow unit can carry a run past the cap and have it killed outright —
which parks nothing. Hence the batch driver's shorter default.

### Added — `test/SquadronGroups.resume.test.js`

20 assertions on the resolver, weighted toward the ways a resume could silently skip: a list
that grew, a list that shrank, an index past the end, and malformed state. Redundant work is
always preferred to a possible skip.
## [2026-07-27] — the other tenant is external, and the lists were told to reject external

A senior on the wing domain could not post to `ca.all@cawgcadets.org`. Nothing was broken
about the account or the address; the cadet-side all-hands lists sit at
`ALL_IN_DOMAIN_CAN_POST`, and **`@cawgcap.org` is not that domain.** Two Workspace tenants
means every sender on the far side is external, however much the same wing they are.

`ANYONE_CAN_POST` is the only setting that admits them. Google has no value meaning
"members plus my other domain", so this is the whole menu.

### Changed — `SquadronGroups.gs` 1.6.0

Managed distribution lists are now created at `ANYONE_CAN_POST` with
`spamModerationLevel=MODERATE`, and `applyGroupSettings()` reconciles
`whoCanPostMessage` and `spamModerationLevel` alongside `allowExternalMembers` on every
sync. Running `updateAllSquadronGroups()` on a tenant repairs its existing lists.

**Why this is now safe to enforce, having been unsafe in 1.2.9.** That version narrowed
the managed keys to `allowExternalMembers` alone, because the callers passed
`ALL_MEMBERS_CAN_POST` for every list and applying it would have dragged the cadet
receive lists *down* from `ANYONE_CAN_POST` and silently re-broken the fan-out — which
carries the original external sender and so fails for exactly the same reason this bug
does. The callers now pass `ANYONE_CAN_POST`, so enforcement only ever moves a group
toward accepting outside mail. The hazard was the value, not the key.

**The openness is real and is not free.** `ANYONE_CAN_POST` accepts mail from anywhere on
the internet, including cadet-facing lists. `spamModerationLevel` is managed in the same
key list so the two cannot drift apart and no caller can widen posting while forgetting
moderation. A list that must genuinely stay closed belongs outside the managed set, not
hand-set in the console — the next sync overwrites that.

### Changed — `UpdateGroups.gs` 1.8.1, `groupAdministration.gs` (comment only)

The sheet-driven path already enforced `ANYONE_CAN_POST`; it now manages
`spamModerationLevel` too, so the two paths agree about a group they both touch. The
receive-list audit's docs no longer describe cross-tenant fan-out as the only victim —
a person writing across directly hits the identical wall.

### Added — `groupAdministration_repairReceiveListPosting(dryRun)`

Enforcing a setting on every sync only helps groups the sync visits. It reaches a list through
`shouldCreateDistributionLists()`, which returns false for anything that is not **UNIT** scope,
so three populations are permanently out of reach however often it runs:

| | Why the sync never sees it |
|---|---|
| `ca.all@...` — the wing all-hands | no CAPWATCH org at all; the loop iterates orgs |
| `ca006.all`, `ca006.dty.all` — group HQ | `scope=GROUP`, excluded by the UNIT filter |
| lists whose unit left CAPWATCH | the group outlived the org that justified it |

Confirmed on both tenants: the cadet-side blocking set was `ca.all` plus the eight group HQs
(006, 008, 070, 188, 205, 213, 303, 445), and the wing side showed the same eight orgs across
`.cadets`/`.parents`/`.dty.all`. **The wing all-hands is in this population** — which is to say
the group a member actually writes to was the one nothing would ever fix.

DRY RUN by default. Scope is whatever the audit flags, so the audit stays the single definition
of "should accept outside mail" and the two cannot drift apart. **Settings only — it never
changes membership**, because the orgs involved are precisely the ones the sync does not model,
and inventing a roster for them is a policy question, not a repair.

### Added — `test/SquadronGroups.groupSettings.test.js`

The managed key list is now pinned by test, in both directions: the delivery-governing
keys are reconciled, everything else the callers pass is left to console/GAM. This
setting has been wrong twice, once by applying nothing and once by nearly applying too
much, so which keys are in the set is a decision worth holding still. 17 assertions.

## [2026-07-27] — two writers, one group, 1,643 members removed every night

`ca.all@cawgcadets.org` had 1,004 members. The tenant has 1,004 account-holding cadets and
**1,683 cadet-lite** ones. Pulling that thread found something worse than a missing wing list.

**Every unit's `.all` group was being emptied of cadet-lite members daily and refilled an hour
later.**

| | |
|---|---|
| **05:24** | `updateEmailGroups` removes ~1,643 external addresses from every `ca###.all` |
| **06:01** | `updateAllSquadronGroups` adds them all back |

Members were off their unit's all-hands for roughly half an hour every morning, and mail sent in
that window did not reach them. Measured, not inferred: `getEmailGroupDeltas()` reported 1,643
pending removals across 55 units.

**One cause, three symptoms.** The sheet-driven path could not see cadet-lite members —
`getMembers()` filtered them out before the desired set was built — so every address
SquadronGroups had added looked like a stranger. The same blindness meant `ca.all` never had
them, and the blank `Add EXT` column on those rows meant `allowExternalMembers` was being set
**false** nightly, undoing the repair applied that morning.

### Changed — `UpdateGroups.gs` 1.9.0

A new **`Add Lite`** column. Rows that opt in include cadet-lite members, addressed by their
personal CAPWATCH address — the same address SquadronGroups uses, so the two paths finally agree
about what a `.all` group contains.

**`Add Lite` implies external members.** A row asking for accountless members while leaving
`Add EXT` blank asks for two contradictory things, and the losing side is silent: the adds fail,
or the flag flips nightly against another writer. One column now settles both.

**Cadet-lite members are in the member set by default and still reach no group without the
opt-in.** `createMemberObject` sets `email: null` and only the Workspace directory fills it, so
they arrive addressless and every group's `isMatch && .email` test skips them — exactly as when
they were filtered out entirely. Nothing changes for any row that does not ask.

### Fixed — 1.10.0: the apply loop could not add an external member at all

With the gate working and the delta correct, `ca.all` still gained **one** member out of 1,644
pending. The delta was right — `desired=2649 add=1644 drop=0` — so the fault was downstream:

```javascript
// Skip external emails except for groups whose Attribute is 'contact'
if (!finalEmail.endsWith(CONFIG.EMAIL_DOMAIN) && groupAttributeByName[category] !== 'contact') continue;
```

Every off-domain address was dropped by a bare `continue` unless the row's Attribute was literally
`contact`. **No log, no counter, no error.** The run reported `totalAdded: 2, totalErrors: 0` while
declining 1,644 adds it had just computed. The two that succeeded were the only internal ones.

The unit `.all` groups hid it: SquadronGroups had already put those members there, so the delta
said *keep* rather than *add* and the skip never fired.

Permission now comes from the Groups sheet — `Add EXT`, or implied by `Add Lite` — with `contact`
kept as an alias so untouched rows behave as they always have. **Declined adds are counted and
reported** in the per-group log line, because an add that vanishes without a number is how 1,644
of them went unnoticed.

This was never specific to cadet-lite. Any row marked `Add EXT` has been unable to add an
external member for as long as that line has existed.

### Fixed — the gate, in 1.9.1, before it reached a mailbox

The first cut gated on **a member having no address until one was supplied**, reasoning that a
row which had not opted in would skip cadet-lite members naturally. That belief was wrong:
`addContactInfo()` fills `.email` from the CAPWATCH PRIMARY contact for **every** member,
cadet-lite included. The gate never engaged — and cadet-lite members were eligible for *every*
group whose criteria they matched, not the two rows that asked for them.

The live preview said so plainly: `reachable: 0` from the helper that was supposed to be doing
the work, next to removals dropping from 1,643 to 12. The fix appeared to work while doing
nothing it was designed to do.

Eligibility is now decided by **CAPID**, against the same cadet-lite definition
`SquadronGroups.gs` uses, so the two paths cannot disagree about who is cadet-lite. Rows without
the opt-in are handed a member set with them removed.

**A membership rule must not rest on a field another module populates for its own reasons.**
That is what the tests now pin: the first assertion is that a cadet-lite member *does* carry an
address, so an address-based gate would leak — and the fourth is that the same person with the
same address is in the group or not purely according to whether the caller included them.

## [2026-07-27] — six transition triggers become one, freeing five slots

Apps Script allows **20 triggers per script per user**, and the cadets project was at 20 — six of
them the cadet→senior lifecycle, one per phase. Adding anything meant displacing something.

### Added — `CadetTransition.gs` 1.2.0

```
armTransitionPipelineTrigger()      // one daily trigger, 03:00; replaces the six
runCadetTransitionPipeline()        // the driver; run by hand to test
checkTransitionPipelineStatus()     // which phase the next run starts from
```

**Collapsing them is safe, and the existing scheduler said why.** The phases were never an hour
apart because each takes an hour — the comment in `armTransitionTriggers()` has always read
*"each phase only acts on rows the previous one made ready, so a member who does not finish one
phase in a day is simply picked up the next day."* And the three expensive phases — Gmail, Drive,
Contacts — already self-limit and schedule their **own** continuation triggers a minute out, so
none of them ever needed a private 30-minute window to make progress. Those continuations are
untouched; they belong to the phases.

**Two properties six triggers gave for free, kept deliberately:**

| | |
|---|---|
| A failing phase does not stop the others | Six triggers meant a 3 a.m. exception never prevented the 4 a.m. run. Each phase is caught and logged and the pipeline carries on — aborting would quietly make this *less* resilient than what it replaces. |
| Order is preserved | detect → resolve → Gmail → Drive → Contacts → remind |

**And one that is new.** The run parks the phase it stopped at and resumes there, wrapping when
the cycle completes. Without it, a day that ran out of budget mid-list would re-run the early
phases and never reach the late ones — which is precisely how nine squadrons went unvisited for
weeks in the group sync. On a six-item list that would have taken far longer to notice.

`armTransitionTriggers()` still installs the original six for anyone who wants them back, and
`listTransitionTriggers()` now reports the consolidated driver too — filtering on the six phase
handlers alone would have said "none installed" on a project running the pipeline.

## [2026-07-27] — the trigger repoint, as one call instead of three panel steps

Merging the resume fix changed nothing on its own: the daily trigger still pointed at
`updateAllSquadronGroups()`, the entry point that **cannot** resume. Repointing it by hand is
three steps in the Triggers panel — delete the old handler, add the new one, do not leave both —
with a wrong outcome available at each, and the wrong outcome is **silent**. The sync goes on
reporting success while never reaching the end of the list. That is precisely how nine CAWG
cadet units went unvisited for weeks.

### Added — `SquadronGroups.gs` 1.10.0

```
installSquadronGroupsBatchTrigger()     // daily at 06:00
installSquadronGroupsBatchTrigger(4)    // or whatever hour suits
```

Removes any trigger on `updateAllSquadronGroups` **and** on the batch handler itself, then
installs one daily trigger on `updateAllSquadronGroupsBatch`. Leaving both installed would mean
two schedules walking one shared `SQUADRON_BATCH_INDEX` position.

**Deletes before it creates**, so the swap needs no free slot — which matters, because the cadets
project sits at the Apps Script ceiling of 20 triggers per script per user. If the create still
fails, the error names what was already removed: the operator is one trigger down at that point
and needs to know it rather than discover it later.

> Found while installing the parent-email trigger: the cadets project was at 20, and one of them
> was `cleanupLegacyCrossTenantContacts` — a **one-time manual migration tool** that lives in
> `config-tenants/`, never in `src/`. Someone pasted it into the project and put a time-based
> trigger on it; a later `clasp push` removed the pasted file, since `.claspignore` ships only
> `src/**`. It had been firing at a function that no longer existed, at a 100% error rate, for
> long enough to be background noise.

## [2026-07-27] — units get told which parent addresses Google refuses

A unit's parents list silently fails to reach some families, and the unit has no way to know.
The addresses are wrong in eServices; only the automation ever finds out, in a log nobody reads.

### Added — `ParentEmailNotify.gs` 1.0.0

```
previewBadParentEmails()                        // read-only; sends nothing, records nothing
testParentEmailDigestForOrg(orgid, recipient)   // one unit's REAL digest, to a test address
notifyBadParentEmails()                         // sends, and records who was told
installParentEmailMonthlyTrigger()              // 1st of each month, ~08:00
resetParentEmailNotifyState()                   // discard the cooldown; re-reports everything
```

`testParentEmailDigestForOrg()` exists because the alternative to having it is discovering how
the email reads by sending it to ten commanders. It applies the same suppression the real run
does — so what you see is what they would get — but **writes nothing**, leaving everything shown
still reportable afterwards. It renders against the real addressee so the greeting is honest,
while the mail goes only to the test recipient with no Cc, and sends without the automation
`from` override so it works from whoever is signed in.

Recipients are the unit **Commander and Deputy Commander for Cadets**, and nobody else — a
parent contact hangs off a cadet's record, so the cadet-side command staff are the ones who act
on it. Recipient resolution is shared with `RecoveryEmailNotify.gs`, which already solves the
awkward part: on the cadets tenant both of those people are **seniors**, absent from
`getMembers()` and holding accounts on the other domain.

**The data has to come from a ledger, and that is not an implementation detail.** An address is
only ever discovered to be bad by asking Google — a well-formed `gmail.com` address that does not
exist is indistinguishable from a good one until `members.insert` refuses it. There is no
read-only check. So the squadron sync writes what it learned to
`RejectedMemberAddresses.json` (`SquadronGroups.gs`), and this module reports only what was
actually observed. It invents no validation rules of its own and so cannot accuse a working
address.

Three consequences, all deliberate:

- **Until `updateAllSquadronGroups()` has run, this reports nothing.** An absent ledger means
  "not yet looked", which is not the same as "nothing wrong", and the two must not be conflated.
- Only addresses on lists the sync manages are ever tested.
- A corrected address stops being refused, drops out of the ledger on the next complete run, and
  stops being reported — with nobody pruning anything.

**Suppression is per member AND per address**, not per member. A cadet with two bad parent
addresses who gets one corrected must still be told about the other; keying on the member alone
would hide it for three months. Three calendar months, matching the other digests — reported on
the 3rd, eligible again on the 3rd.

**The person is rebuilt from the current extract, not stored.** The ledger holds an address; the
cadet is re-resolved through `MbrContact` at send time. A cadet who has left, or whose record was
corrected, resolves to nothing — so a commander is never mailed about someone who is not theirs.
Siblings sharing one address produce one row each.

### Changed — `SquadronGroups.gs`

`reportRejectedMemberAddresses_()` now writes the ledger. A run that saw every squadron
**replaces** it; a **paused slice only merges**, because treating a partial run's findings as the
whole picture would delete what earlier slices found and leave the digest reporting a fraction of
the problem while looking complete.

### Added — `test/ParentEmailNotify.test.js`

25 assertions on the two decisions that can hurt someone: the cooldown boundary (a day short,
exactly on it, across a year end, and never-reported treated as reportable rather than
suppressed), and resolution refusing to produce a row for a cadet who is not in the extract. All
names, CAPIDs and addresses synthetic.

## [2026-07-26] — a missing welcome email is now detected, not stumbled over

The resend below repairs a member you already know about. This is how you find out.

**"Was this member ever welcomed?" is not answerable from the directory** — nothing on an
account records whether mail was sent to it. So it gets written down.

### Added — `WelcomeEmailAudit.gs` 1.0.0

```
seedWelcomeLedger()                    // DRY RUN by default — one-time baseline
seedWelcomeLedger(false)               // write it
scanUnwelcomedAccounts()               // read-only report
notifyUnwelcomedAccounts()             // mails IT the MISSED list; silent when empty
installWelcomeAuditMonthlyTrigger()    // 1st of each month, ~08:00
```

`sendWelcomeEmail()` now records each send in a Drive ledger
(`WelcomeEmailLedger.txt`, alongside the other state files). One line, at the send site, so
provisioning **and** the resend **and** any future sender are covered by it. A ledger write
that fails is caught and logged, never rethrown — the email is already gone, and bookkeeping
must not turn a successful send into a reported failure (`UpdateMembers.gs` 1.20.0; nothing
about who is mailed or when is changed).

**The seeding problem is the whole design.** On day one nobody has an entry, so the entire
wing looks unwelcomed. `seedWelcomeLedger()` sets the baseline on the one honest piece of
hindsight evidence available: an account that **has been signed into** plainly received
working credentials at some point, whatever the route. Those are recorded as welcomed.
Accounts that have **never** been signed into cannot be judged either way — so the seed
deliberately does *not* vouch for them. Marking them welcomed on no evidence would
permanently bury the very members this exists to find, including the one that started this.

Three verdicts, and they are worth different amounts:

| | Meaning | Mailed to IT |
|---|---|---|
| **MISSED** | created after the baseline, no send ever recorded | yes |
| **UNKNOWN** | predates the baseline, never signed into — a genuine maybe | no |
| **WELCOMED** | a recorded send, or seeded from login history | — |

Only MISSED is mailed. UNKNOWN cannot be acted on without judgement — a member who was
welcomed and never logged in is indistinguishable from one who was never welcomed — and a
monthly message that is mostly noise stops being read. That population is *already* surfaced
to unit command staff by `RecoveryEmailNotify.gs`'s LOGIN condition; this is not a second copy
of it, and it does not mail units.

**Guards, both tested:**

- An **unseeded or lost ledger produces UNKNOWN, never MISSED.** The failure mode that would
  discredit the report on its first run is a wing-wide list of false accusations, so the
  no-baseline case cannot accuse anyone.
- **Re-seeding is refused** once a baseline exists (`{force: true}` overrides). Moving the
  baseline forward silently reclassifies every confirmed MISSED account as merely UNKNOWN.
- Accounts younger than `NEW_ACCOUNT_GRACE_DAYS` (2) report PENDING, not MISSED — provisioning
  creates the account and sends the mail as separate steps, and a resend may be in flight.

**A useful side effect:** a welcome email that *fails* to send during normal provisioning now
shows up too. That failure is caught and logged at the call site and nothing revisits it, so
it was previously as invisible as an out-of-band account.

**Still manual by design.** The audit reports; it never resets a password on its own.
`resendWelcomeEmail()` stays the fix, with its own guards.

## [2026-07-26] — a welcome email can be sent to an account created out-of-band

A new senior never received a welcome email. They joined in June, cleared Level I in July, and
had an account by 01Jul — a week *before* Level I. That ordering is the whole story:
`REQUIRE_LEVEL_I_FOR_SENIORS` was withholding provisioning the entire time, so the script
cannot have created that account. Someone made it by hand.

**The welcome email has exactly one call site** — inside the `if (!user)` insert branch of
`addOrUpdateUser()`. When the gate finally lifted, `Users.update()` at that address succeeded,
`user` came back truthy, and the whole insert branch (welcome email included) was skipped. The
member read as an ordinary existing account getting a routine attribute update. Every account
created in the Admin console or by GAM has this hole, nothing detects it, nothing repairs it.

### Added — `WelcomeEmailResend.gs` 1.0.0

```
previewWelcomeEmailResend(capid)        // read-only: would this send, and to whom
resendWelcomeEmail(capid)               // resets the password, sends the welcome email
resendWelcomeEmail(capid, {force:true}) // bypass the login-history / suspended guards
```

**A resend is a password RESET plus a send, not a re-send.** The original temp password is
generated at insert time and stored nowhere, so the message cannot be reproduced — the only
way to deliver working credentials is to mint new ones. That makes this destructive on a live
account, so the decision of whether a member may be resent is a pure function
(`welcomeResendEligibility_`) with hard guards, covered by `test/WelcomeEmailResend.test.js`:

- **An account with login history is refused.** The member is demonstrably using it; a reset
  locks them out of a working account to fix a cosmetic gap. `{force: true}` overrides, and the
  log records that a guard was bypassed.
- **Credentials are never mailed to the mailbox they unlock.** If CAPWATCH holds no address
  outside the tenant, the welcome email would land in the account the new password is needed to
  open. Refused outright — `force` does not override, because the send would be useless rather
  than merely unwise. This one is silent in the wild: it looks like a successful send.
- Suspended (refusable with `force`), archived, no account, and no CAPWATCH record are each
  refused with their own reason slug.

The account is resolved by **CAPID through the duplicate-guard path**
(`findExistingAccountsByCapid_` + `chooseAuthoritativeAccount_`), never by deriving
`first.last` — an out-of-band account frequently is not at the derived address, which is how
it came to be missed in the first place.

`UpdateMembers.gs` is untouched: `sendWelcomeEmail()` and `generateTempPassword_()` are reused
as-is, so provisioning behavior is unchanged.

**Not fixed here:** nothing yet *detects* an account that never got a welcome email — this is
a repair tool an admin runs against a known CAPID. Closing that gap needs a welcome-sent marker
and a one-time backfill, or the entire existing wing gets re-welcomed on first run.

## [2026-07-26] — updateEmailGroups() can be run in slices

`updateEmailGroups()` timed out. The run's length tracks the number of membership **changes**,
not the number of groups — each one is an API call with `API_DELAY_MS` pacing behind it, plus a
15-second cooldown every 25 adds. A steady day is minutes. The day the level rule changed to
highest-rung-only, thousands of memberships moved at once, and the execution limit arrived first:
the run was killed partway with no record of where it got to, and the next run started over.

### Added — `updateEmailGroupsBatch(budgetMinutes)` (`UpdateGroups.gs` 1.8.0)

```
updateEmailGroupsBatch()        // 25 minutes, sized for the 30-minute execution limit
updateEmailGroupsBatch(5)       // a shorter slice, e.g. on the 6-minute tier
checkEmailGroupsBatchStatus()   // how far along, changes nothing
resetEmailGroupsBatchProgress() // discard the parked run, recompute next time
```

The first call computes the deltas and parks them in Drive with the group metadata; each call
applies what fits in its budget and saves its position, down to the individual member. Run it
again — by hand, or by pointing the daily trigger at it — until it reports complete.

**Rather than a second copy of the add/remove logic, the existing loop learned to stop.**
`updateEmailGroups()` now takes an optional `{deadlineMs, resume}` and returns where it got to;
called with no arguments it behaves exactly as before, so the daily trigger is unaffected.
Two details that matter:

- **Metadata and settings are not re-applied** when resuming into the middle of a group.
- **The error sheet is written only on the run that completes.** `saveErrorEmails()` clears and
  rewrites the tab, so writing it per slice would leave only the last slice's errors on it;
  errors accumulate in the parked state instead.

Re-applying a slice is harmless in any case — an add that already happened returns 409 and a
removal 404, both already caught — so a hard timeout between saves costs at most one group.
A parked run older than 12 hours is discarded rather than resumed: past that point its deltas
describe a directory that has moved on.

### Not changed — the pacing itself

`CONFIG.API_DELAY_MS` is 250 ms before every call, and the add path sleeps another 250 ms plus
15 seconds every 25 adds. That is ~0.5 s per change before the API is even consulted, well under
what the Directory API allows (~15 queries/second). Loosening it would cut these runs
substantially, but it is a shared constant every module paces on, so it is a deliberate,
separate decision rather than something a timeout fix should take unilaterally.

### Added (`test/UpdateGroups.batch.test.js`)

A faked clock and directory: pausing mid-group and at a group boundary, the exact resume
position in both cases, a two-pass run applying all five changes exactly once with totals
carried across, and metadata not re-applied on resume.

## [2026-07-26] — Level groups are rungs, not badges

Two corrections from the wing DA after the first live run.

### Changed — a member is in the group for their HIGHEST level only (`UpdateGroups.gs` 1.7.0)

1.6.0 put a member in every level group they had ever completed, so a Level V holder sat in
`all-level-ii` through `all-level-v` at once. The levels are **rungs, not badges**: a wing that
mails `all-level-ii` means the people sitting at Level II, not everyone who passed through it.

`professionalLevel` now matches a member's highest completed level. The rungs come from
`getProfessionalLevelLadder_()`, which reads them out of `PL_Paths` by name (`Level 2 Part 1`
and `Part 2` both land under 2), so a level re-split upstream needs no code change.

**The level groups will shrink on the next run.** Everyone who has moved past a level is removed
from that level's group and appears in their own. The run log counts them as
`excludedHoldingAHigherLevel`, and a group that empties out entirely for this reason says so
rather than reporting a mystery.

A value naming something off the ladder — `Squadron Commander Training`, `TLC Basic` — has no
rung to be highest and keeps plain "holds this path" semantics.

### Added — `professionalLevelInProgress` (`UpdateGroups.gs` 1.7.0)

Members holding **some** parts of a level but not all of them: built for "finished Level II
Part 1, hasn't finished Part 2", the list a wing sends a nudge to.

```
Group Name: all-level-ii-part-1-only
Attribute:  professionalLevelInProgress
Values:     Level 2
```

Approved credit only, same as the completed-level rule, so a pending Part 1 is not progress. A
level CAPWATCH stores as a single path (3, 4, 5) has no in-progress state and warns instead of
creating a group that can never fill.

### Added (`test/UpdateGroups.professionalLevel.test.js`)

A member who climbed the whole ladder, pinned at each rung; Part 1 alone neither completing
Level 2 nor promoting past Level 1; the in-progress list; single-path levels rejected for
in-progress; PathID-of-a-part still meaning the whole level.

## [2026-07-26] — The E&T level groups were reading a retired program

The `ca.all-level-ii` … `-v` groups came back empty even after achievement names and IDs both
resolved correctly, and the reason turned out to be more interesting than a bad Values column:
**they were pointed at the wrong CAPWATCH subsystem entirely.**

`Achievements.txt` still lists `Level II`–`Level V` as AchvIDs 131–134 under FunctionalArea
`ET-Senior` — leftovers from the pre-2018 professional development program. The Groups sheet
was keyed on those IDs, correctly, and `MbrAchievements.txt` carries **no row** for any of them,
because the current PD program records nothing there. A rule keyed that way resolves cleanly,
reports its resolved AchvIDs, and matches nobody — the most convincing possible way to be wrong.

The live program lives in the **`PL_*` tables**:

| File | Holds |
|---|---|
| `PL_Paths.txt` | `PathID` → `PathName`; Level 1 = 4, **Level 2 = 7 + 8** (Part 1 / Part 2), Level 3 = 3, Level 4 = 2, Level 5 = 1 |
| `PL_MemberPathCredit.txt` | `MemberPathCreditID, PathID, CAPID, StatusID, Completed, Expiration, ExtraCreditEarned` |
| `PL_Lookup.txt` | `StatusID` 8 = `APPROVED`, 26 = `PENDING`, 27 = `DISAPPROVED` |

### Added — `professionalLevel` attribute (`UpdateGroups.gs` 1.6.0)

```
Category: education-training   Group Name: all-level-ii
Attribute: professionalLevel   Values: Level 2
```

- **Level 2 requires both parts.** CAPWATCH splits it into `Level 2 Part 1` and `Level 2 Part 2`,
  so a value naming the level resolves to both PathIDs and the member must hold *each*. Part 1
  alone is progress, not the level, and is counted separately in the log as `partialCredit`.
- **Values may be a path name or a PathID**, with Roman numerals folded to digits, so `Level 2`,
  `Level II` and `3` all work. Several values are an OR.
- **Approval is read from `PL_Lookup`, not hardcoded** — the extract states which StatusID means
  APPROVED, so a future renumbering does not silently admit pending credits.
- Any path name works, not just levels: `Squadron Commander Training` is a path too.
- `listProfessionalLevelPaths('level')` prints the real names.

### Changed — the PL tables are read by column name

`readCapwatchTable_()` reads a CAPWATCH file *with* its header and returns objects.
`parseFile()` drops the header, which encodes column meaning as a magic index — the habit that
left Member.txt's Expiration column unverified at index 16 for months. The PL_* tables are new
here and have no such folklore, so they start out read by name.

> **The Groups sheet needs editing for this to take effect**: the seven `ca.all-level-*` rows
> must move from `achievements` + `131`–`134` to `professionalLevel` + `Level 2`…`Level 5`.
> Until then those rows keep resolving to retired AchvIDs and keep matching nobody — which the
> run now says out loud rather than silently creating an empty group.

## [2026-07-26] — Group-echelon DLs for positions the echelon does not have

### Fixed (`UpdateGroups.gs` 1.5.0)

Duty and rank rules **seed one DL per group-echelon org before matching anybody**, so that a DL
whose holder has just left stays in the desired state long enough for that person to be removed
from it. Right for a position the echelon *has* and is merely vacant; wrong for one it does not
have — and nothing distinguished the two. CAP puts a **Director of Information Technology at wing
and above, an IT Officer below it**, and gives groups **no Inspector General**, so each of those
rows minted one permanently empty DL per group, every run.

The two cases are now told apart by asking whether the group already exists. An empty seed for a
group already in the tenant is kept (vacant seat, stale members still to clear); an empty seed for
a group that does not exist is dropped rather than created. That reads the doctrine off the data
rather than off a hardcoded list of wing-only offices, so it covers the rest of the class —
Director of Operations, Director of Personnel, and anything else the wing adds later — without
maintenance.

The group index behind it is built lazily, once per run, and only when a seed comes back empty; a
failure to list is not fatal, it just skips clearing vacant DLs until the next run. The same
pruning applies to the `dutyPositionLevelStaff` GROUP seeding.

### Added — `cleanupEmptyEchelonGroups(dryRun = true)`

Deletes the empty group-echelon DLs already created. Candidates come from the Groups sheet, so
nothing outside the managed naming scheme is considered, and a group is offered **only when it is
empty** — which is also why nothing is lost: an empty DL has no manual additions in it, and one
whose position does exist at that echelon is recreated as soon as somebody holds it. Preview by
default.

## [2026-07-26] — Three ways a group rule matched the wrong people, all silent

Reported together, and they turn out to be one habit rather than three bugs: a rule that
matches nothing, or matches the wrong record, produces a **real Google Group with the wrong
membership and no error anywhere in the log**. An empty DL is worse than a missing one — it is
in the GAL, it accepts mail, and it delivers to nobody.

### Fixed — a wing office DL held the assistant and not the primary (`UpdateGroups.gs` 1.4.0)

`dutyPositionIdsWingHQ` tested **the member's home unit** (`squadrons[member.orgid]` being Wing
HQ) rather than the org the duty is actually held at. Nearly all wing staff are members of a
squadron and hold their wing duty on top of that, so the rule quietly reduced to *"wing staff who
are also Wing HQ members"*. For the wing recruiting office that was the assistant; the primary,
being a squadron member, was dropped. The header comment ("Wing HQ ONLY for the configured duty
position titles") described the intent, not the code.

It now reads the duty's own org via `getDutyAssignmentOrg_()` — the same resolution
`dutyPositionLevelStaff` and `dutyPositionIdsGroupScope` already used. The rule reads the same
way from the other side too: a Wing HQ member holding the title **at a squadron** is now
excluded, where before they were included on their home unit alone.

### Fixed — retired duty titles no longer miss (`UpdateGroups.gs` 1.4.0)

Duty matching compared raw CAPWATCH strings. `DUTY_TITLE_OVERRIDES` (the ICL to CAPR 30-1 rename
of `Recruiting & Retention Officer` → `Recruiting Officer`) was applied only to *displayed*
titles, so a Groups row keyed on the current title silently skipped every holder whose eServices
record still carried the old one — 2 of CAWG's rows at last count. `normalizeDutyId_()` now runs
both sides through `formatDutyTitle_()`, so the sheet can be written against current doctrine
regardless of what is stale upstream. Fixing the record in eServices is still the real remedy.

### Fixed — the Education & Training level groups were unfillable (`UpdateGroups.gs` 1.4.0)

`MbrAchievements.txt` identifies an achievement **only by numeric AchvID** (Level I is `96`); the
name lives in `Achievements.txt`. The `achievements` rule compared the Values column directly
against that ID column, so `ca.all-level-ii` and its siblings — written with names, which is the
obvious thing to type — matched zero rows and created empty groups.

Values may now be **either form**. Names resolve through a `Achievements.txt` index, matched with
case, whitespace, punctuation and Roman numerals normalized, so `Level II`, `level ii` and
`Level 2` all land on the same achievement. `listAchievementNames('level')` prints the real
strings.

### Changed — a rule that matches nobody no longer creates a group

- An **unknown `Attribute`** (a typo in the sheet) used to fall through to a warning *after* its
  wing-level group had already been pre-created, so the typo shipped a permanent empty group.
  Nothing is created now.
- An **achievements row matching no member** is left out of the desired state and warned about,
  with the row counts by status attached. This also means a mistyped row can no longer *empty* a
  group that already holds the right people — the failure mode that made the previous behavior
  dangerous rather than merely untidy.

### Added — see what a rule does before it does it (`UpdateGroups.gs` 1.4.0)

`previewEmailGroupRows(filter, showMembers)` walks the Groups sheet and prints the group IDs each
row generates and how many members land in each, marking the empty ones. It creates nothing,
adds nobody and removes nobody. `previewEmailGroupRows('recruiting', true)` is the fastest way to
confirm a duty rule holds the people it should.

### Fixed — command-staff DLs for positions the unit does not have (`SquadronGroups.gs` 1.5.0)

CAP establishes a plain **Deputy Commander only at senior units**; cadet and composite units have
**Deputy Commander for Cadets** and **for Seniors** in its place. Every cadet and composite
squadron was nevertheless getting a `ca###.deputy-commander` DL, which no CAPWATCH duty can fill.

Selection moved into `COMMAND_STAFF_SUFFIXES_BY_UNIT_KIND_`, keyed by what the unit's type
establishes. A **flight** is classified by its own membership (CAPWATCH types it `FLIGHT`
whatever it runs), and a unit whose kind cannot be determined — unknown type, or a flight with no
members — now gets the **Commander DL only** instead of all four. Creating all three deputy
flavors "to be safe" guaranteed at least two were wrong.

`cleanupUnnecessaryDistributionLists()` (preview by default) now offers the wrong ones already
created: `.deputy-commander` at a cadet or composite unit, `.deputy-commander-cadets/-seniors` at
a senior one. Flights are left alone. **Deleting a group takes its manual "User Additions"
members with it** — read the preview.

> Cadet units get `deputy-commander-seniors`, same as composite: **Deputy Commander for Seniors
> is a valid billet at both** (confirmed by the wing DA, 2026-07-26). The prior code comment
> claiming CAPWATCH shows no DCS on cadet units was reading a vacancy as a rule — an unfilled
> billet is still a billet, and its DL should exist for the day it is filled.

### Added (`test/UpdateGroups.dutyGroups.test.js`, `test/UpdateGroups.achievements.test.js`, `test/SquadronGroups.commandStaff.test.js`)

The primary-vs-assistant case pinned directly (synthetic members), the home-unit rule inverted in
both directions, retired-title matching, achievement names vs IDs vs Roman numerals, status
filtering, and the command-staff matrix including what cleanup will and will not delete.

## [2026-07-25] — Per-unit renewal digest; senior notices carry no CC

### Added — the unit renewal digest (`SendRetentionEmail.gs` 1.7.0, `config.gs`)

Each unit now receives one message listing **every member under its command expiring this month**,
cadets and seniors alike. Addressed to the commander with the recruiting officer copied; where a
unit has only one of the two it goes to whichever exists, and a unit with **neither** is reported
in the run summary rather than silently skipped.

**A senior's renewal notice carries no unit CC at all.** That is the reason the digest exists: a
senior's renewal is between them and the wing, so their unit gets a worklist addressed to it rather
than a blind copy of somebody else's mail. **Cadet renewals are unchanged** and still CC the
commander and recruiting officer — that is a cadet protection matter, not a retention one — and
cadets also appear in the digest, so the unit sees one complete list.

| | |
|---|---|
| Squadron Commander | Cadet mail only (turning 18/21, cadet renewals) + digest addressee |
| Deputy Commander for Cadets | Turning 18/21 |
| Recruiting Officer | Cadet renewals (CC) + digest copy |

Deduplicated per unit per month on the same Log sheet as individual sends, keyed
`RENEWAL_DIGEST|<orgid>` — the ORGID sits in the CAPID column because the unit is what was mailed.
It lists every expiring member **including those whose own notice was skipped as already-sent**:
the unit wants the full picture, not a record of what one run happened to do.

`logEmailSent()`'s sheet handling is extracted to `retentionLogAppend_()` so both writers share one
sheet and one header.

### Changed — duty holders decoupled from the commander

1.5.0 had duty holders included only where the commander was already being CC'd. That could not
survive the digest, which must reach a recruiting officer at a unit whose commander is unreachable.
`retentionCcList_()` now takes an explicit `includeCommander` flag and adds each recipient
independently.

**Side effect worth knowing:** an age-milestone mail at a unit with no reachable commander now
still reaches the Deputy Commander for Cadets, where under 1.5.0 it reached nobody. Recorded so the
reversal is not mistaken for drift.

`previewRetentionCcLists()` now prints the digest addressee and copy per unit, and flags units that
would hear nothing.

## [2026-07-25] — Retention CC preview, and a trigger installer

### Added — `previewRetentionCcLists()` (`SendRetentionEmail.gs` 1.6.0)

Prints the resolved unit CC for every unit the next run would touch, then the three things worth
fixing before it runs. Sends nothing, writes nothing.

The reason it exists: a cadet email now carries up to **three** unit addresses, and two of them
may be **derived** — reconstructed as `first.last@<command domain>` because the directory had no
account for that CAPID. A derived address reproduces the DEFAULT account name, so it is wrong for
a rename, a manual creation, or a `.2` duplicate. Gmail accepts it and bounces afterward, per
recipient. That is invisible until it happens, so the preview lists them up front.

Each resolved address now records which of the three routes produced it (`directory` / `derived` /
`capwatch`) — classification only; the resolution order itself stays in
`rcResolveRecipientEmail_()`. The summaries are: units with no reachable commander (which send
with no CC at all), CC'd duty positions nobody holds, and every derived address in one list.
`testRetentionEmail()` calls the dump, and it can be run on its own.

### Added — `installRetentionMonthlyTrigger()`

Mirrors `installRecoveryComplianceMonthlyTrigger()`: idempotent (drops existing
`sendRetentionEmails` triggers first, leaves others alone), 1st of the month at ~10:00, after the
daily `getCapwatch()`.

It **refuses to install** where `PROFILE_.RUN_RETENTION_EMAILS` is false, throwing with the reason
rather than creating a trigger that fires into a no-op — a trigger that exists to do nothing is
worse than no trigger, because it looks like the feature is running.

It does not and cannot solve the identity problem: a time-driven trigger runs as whoever creates
it, so this must be run **signed in as the automation account**, and the Triggers panel confirmed
afterward.

## [2026-07-25] — Retention unit CC carries the staff who own the subject

The commander was the only unit recipient on retention mail, which put the notice in the command
channel but not in front of the person whose job it actually is.

### Changed (`SendRetentionEmail.gs` 1.5.0, `config.gs` 1.13.0)

- **Turning 18 / 21 now also CC the unit's Deputy Commander for Cadets** — the person who walks a
  cadet through the decision the mail describes.
- **Renewals now also CC the unit's Recruiting Officer** — retention is that officer's job.
- Titles live in `RETENTION_CONFIG.CC_DUTY_TITLES` and are matched through `formatDutyTitle_()`,
  reused rather than reimplemented, so the legacy `Recruiting & Retention Officer` rows the ICL to
  CAPR 30-1 renamed still match and the trailing whitespace the feed ships on every duty value is
  handled in one place.

Four behaviors worth stating, because they are decisions rather than fallout:

1. **A duty nobody holds is simply absent** — "if one is assigned" — and the commander is CC'd
   alone.
2. **The extra staff ride on the commander CC, never replace it.** With no resolvable commander
   there is no CC at all, so the mail cannot quietly redirect to a different person. This is also
   what keeps renewals cadets-only without a second condition: seniors have never had a commander
   CC, so they gain no unit CC.
3. **Primary beats assistant**, so "the unit's recruiting officer" is one person rather than a
   unit's whole staff. An assistant is used only when nobody holds the duty primary.
4. **Deduplicated by address** — in a small unit the commander is frequently also the recruiting
   officer, and should appear once.

`retentionCommanderIndex_()` becomes `retentionUnitStaffIndex_()`, now also walking
`DutyPosition.txt` plus a `Member.txt` name pass (duty rows carry only a CAPID). Both are skipped
entirely when no email type asks for staff. Duty holders resolve through the same
directory → derived → CAPWATCH chain as commanders. `getCommanderInfo()` is unchanged for callers,
including `CadetTransitionMigrate.gs`.

Reading `Member.txt` raw is deliberate, matching `RecoveryEmailNotify.gs`: on the cadets tenant
`getMembers()` returns cadets only, but a cadet unit's DCC and Recruiting Officer are seniors —
present in the extract under the cadet ORGID, filtered out of the member set.

## [2026-07-25] — Retention: duplicate protection, and a one-tenant guard

The two remaining reasons the retention trigger was unsafe to arm. Neither is a bug in what the
module sends — both are about what happens when it runs more than once.

### Added — already-sent guard (`SendRetentionEmail.gs` 1.4.0)

The Log sheet has always been **written and never read**, so nothing in the module knew what a
previous run had done. An execution that hit the 30-minute limit partway through the expiring
batch — the largest category — would, on the next firing, start again from the top of the list.
A manual re-run after fixing a problem re-mailed everyone already reached.

Sends are now filtered against `(email type, CAPID)` for the current **calendar month**. That
grain is not arbitrary: it is exactly what member selection already keys on (birth month for
18/21, expiration month and year for renewals), so "already mailed this month" and "already
mailed for this occurrence" are the same statement. The key includes the type, so a cadet who
turns 18 and expires in the same month still gets both.

**It fails open, loudly.** An unreadable or unconfigured log leaves the run behaving exactly as
it did before the guard existed, rather than refusing to send — a spreadsheet read failure should
not become a silent outage of the whole feature. But the run says so in the execution log *and*
in a banner on the summary email, so a low send count is never ambiguous about whether protection
was in effect. A missing sheet is treated differently from an unreadable one: no sheet yet is a
legitimate empty history, and proceeds without a warning.

Only successful sends reach the Log (`logEmailSent()` runs after the send returns), so a failed
send is correctly retried next run. The converse gap is real but narrow: if a send succeeds and
the Log write then fails, that member is re-mailed on a re-run.

### Added — tenant guard (`config.gs` 1.12.0)

`sendRetentionEmails()` is now gated on `PROFILE_.RUN_RETENTION_EMAILS`: **true on seniors, false
on cadets and region**. This module hardcodes `'CADET'`/`'SENIOR'` instead of reading
`MEMBER_TYPES.ACTIVE`, and both wing tenants download the same wing-wide extract, so it addresses
the entire wing from wherever it runs. Arming both tenants did not split the work between them —
it mailed every member twice. Off for region because that tenant has no retention log sheet and
no recruiting role group.

The summary email and `testRetentionEmail()` both report skipped counts, and the preview now
states the tenant profile and whether duplicate protection is available, so what it shows is what
a real run would send.

### Added (`test/SendRetentionEmail.dedupe.test.js`)

Period keying, the type-specific key, last-month non-suppression, missing sheet vs unreadable
sheet, junk rows, numeric-vs-string CAPID cells, and the tenant guard short-circuiting before it
touches CAPWATCH.

## [2026-07-25] — Wing recruiting mailbox: placeholder string was being mailed into every unit group

`SQUADRON_GROUP_CONFIG.PUBLIC_CONTACT.RECRUITING_MAILBOX` held the literal
`'<recruiting email DL here>'`. The consuming code gated on truthiness, and a placeholder is
truthy — so `updatePublicContactGroup()` passed that string to
`AdminDirectory.Members.insert()` **once per unit, per run**. On the seniors tenant, where
`SQUADRON_PUBLIC_CONTACT_AUTO_CREATE` is on, that is every squadron every day. Each call failed
and was caught and logged, so the symptom was a standing bank of member-add errors rather than a
visible break, and the intended feature — wing recruiting seeing unit public inquiries — has
never actually worked.

It was also a **tenant-specific value living in `config.gs`**, a file every `clasp push`
overwrites identically on all three tenants, so it could not have been correct on more than one
of them regardless.

### Changed (`config.gs` 1.11.0, `SquadronGroups.gs` 1.4.0)

- Value moves to the **`TENANT_RECRUITING_MAILBOX` Script Property**. Added to all five tenant
  configs and both setup scripts.
- **Blank disables the behavior, and blank is the default everywhere** — including CAWG.
  Enabling it adds an address to the membership of every squadron public-contact group in the
  wing, which should be a deliberate act, not something a bug fix does on the way past. Set the
  property when the wing decides it wants that fan-in.
- The consuming site now **validates with `sanitizeEmail()`** rather than checking truthiness. A
  set-but-invalid value warns and is skipped instead of being handed to the Directory API.

## [2026-07-25] — Retention addresses corrected; summary only, no per-member BCC

`TENANT_RETENTION_EMAIL` was version-controlled as `recruiting@cawgcap.org` on **both** the
seniors and cadets tenants. That address **does not exist as a group** — confirmed by the wing
DA on 2026-07-25. It is the `To:` of the retention run summary, so the summary had nowhere to
land, and `sendRetentionSummaryEmail()` catches and logs rather than raising: arming the trigger
would have produced a monthly run reporting its results into a black hole.

### Changed (`SendRetentionEmail.gs` 1.3.0, `config-tenants/seniors.json`, `config-tenants/cadets.json`)

- **Both addresses now point at the role group `ca.dty.director-recruiting@cawgcap.org`**, created
  for this purpose and matching the established `ca.dty.*` duty-group convention.
  `TENANT_RETENTION_EMAIL` addresses the summary; `TENANT_DIRECTOR_RECRUITING_EMAIL` is the
  `replyTo` on member-facing mail.
- **Both are now version-controlled** in `config-tenants/`, reversing the deliberate blank. The
  blank existed because the value was an individual's mailbox; a role group survives a change of
  incumbent, so the reason no longer applies. `TENANT_DIRECTOR_RECRUITING_NAME` still names an
  individual and stays blank / Script-Properties-only.
- **Dropped `bcc: RETENTION_EMAIL` from all three member-facing sends.** At wing scale that was a
  few hundred messages a month into one mailbox, and it duplicated a record the Log sheet already
  keeps per send (timestamp, type, CAPID, name, address, commander). The group now gets the
  monthly summary and nothing else.

> ⚠️ A tenant whose Script Properties still hold `recruiting@cawgcap.org` keeps the broken value —
> `config-tenants/` is the canonical copy, not the live one. Re-run `setupTenantConfig()` or fix
> the property by hand, then confirm with a test send.

### Duty-title finding (no code change)

Cross-referencing the CAPWATCH extract for the wing recruiting role turned up a naming trap worth
recording. The CAPWATCH duty string at CAWG is **`Recruiting Officer`**, held at ORGID 188
(CALIFORNIA WING HQ = CA-001, scope WING) by the incumbent as a primary and by one other member as
an assistant. **`Director of Recruiting` is also a real CAPWATCH title, but no CAWG member held it**
in the January 2026 extract — its holders were at PCR region and two other wings.

Group rules match the duty string exactly, and `dutyPositionIdsWingHQ` deletes a group that comes
out empty. So a `Groups` sheet row keyed on `Director of Recruiting` would match nobody at CAWG and
**fail silently** rather than erroring. `SQUADRON_GROUP_CONFIG.PUBLIC_CONTACT.DUTY_POSITIONS`
already uses the correct `Recruiting Officer`.

## [2026-07-25] — Commander CC uses the CAP account, not a personal address

Retention mail CC'd the squadron commander at whatever address sat in their CAPWATCH
record — normally a personal mailbox. `RecoveryEmailNotify` had already worked out how to
reach command staff properly; retention was the one module still doing it the naive way.

### Changed (`SendRetentionEmail.gs` 1.2.0)

- **Address order is now org account first, CAPWATCH last:** the real Workspace account read
  from this tenant's directory → the derived `first.last@<command domain>` → CAPWATCH
  PRIMARY. This is **`rcResolveRecipientEmail_()` reused from
  `notifications/RecoveryEmailNotify.gs`**, not a reimplementation, so the two modules cannot
  drift on what "reach the commander" means. (The dependency already ran the other way: that
  module calls `createEmailMap()` from this one.)
  - The directory step catches what derivation cannot see — a `.2` duplicate, a manual
    creation, a rename.
  - On the **cadets** tenant, command staff are seniors, so `COMMAND_EMAIL_DOMAIN` points at
    the senior domain, the cadet directory yields nothing usable, and the derived senior
    address is correctly preferred over a same-CAPID cadet account.
  - A commander with no usable name and no directory entry still falls through to their
    CAPWATCH address; `null` (no route at all) drops the CC and sends to the member alone.
  - Derived addresses are unverified, so a commander whose account does not follow the
    default naming will have their **CC** bounce. The member's own send is unaffected —
    Gmail accepts the message and bounces per recipient.
- **`getCommanderInfo()` is backed by `retentionCommanderIndex_()`, built once per run.** The
  previous version re-parsed `Commanders.txt` **and rebuilt the entire CAPWATCH email map on
  every call** — once per cadet email sent. The index is reset alongside `clearCache()` so it
  can never outlive the data it was derived from.

### Also changed: cadet-transition email (`CadetTransitionMigrate.gs` 1.1.0)

`getCommanderInfo()` has a second caller — `sendTransitionCompleteEmail_()` CCs the unit
commander on the "your email has moved" notice. It inherits the new addressing, so that CC
also moves from a personal address to the commander's CAP account. **Reviewed and kept
deliberately**, not an incidental side effect: the notice is CAP business either way, and one
resolver for both callers is the whole point — giving this path its own would reintroduce the
drift the change exists to prevent. No code in that module changed; its version note records
the inherited behavior so it is discoverable from the file itself.

Worth knowing: that path runs on the **cadets** tenant, so the address now derives onto
`COMMAND_EMAIL_DOMAIN` (the senior domain) instead of whatever CAPWATCH held.

It also inherits one `getActiveUsers()` listing per execution. The index is cached
module-globally, so this is once per run and not once per member — but note the cadets tenant
pays for a listing whose result `rcBuildCommandDirectoryMap_()` then discards, since command
staff are not on that tenant. The domain guard is deliberately left in one place rather than
restated at the call site to avoid the two copies drifting.

### Added (`test/SendRetentionEmail.commander.test.js`)

First tests for this module: the resolution order on seniors and cadets tenants, the
CAPWATCH last-resort, the null case, build-once caching, graceful degradation when the
directory read throws, and template rendering across two wings (which fails if a hard-coded
wing name or role holder reappears, the regression this module has had before). The test
loads the **real** resolver out of `RecoveryEmailNotify.gs`, so changing the order there
fails here.

## [2026-07-25] — Retention email templates genericized; dead feedback form removed

The three member-facing retention templates were the last hard-coded wing in the codebase.
The 1.7.0 pass genericized the *report* footer but not the mail members actually receive,
which still opened `CIVIL AIR PATROL / CALIFORNIA WING` and closed with a named CAWG role
holder. A wing adopting this module by Script Property alone would have mailed its members
under California Wing's masthead and someone else's signature.

Found while verifying the flow ahead of arming the monthly trigger — which is **still not
armed**. See the pre-arm checklist below.

### Changed (`SendRetentionEmail.gs` 1.1.0, `config.gs` 1.10.0)

- **`{{wingName}}` / `{{orgLabel}}` / `{{signature}}`** in `Turning18Email.html`,
  `Turning21Email.html` and `ExpiringEmail.html`, filled from `CONFIG.WING_NAME`,
  `CONFIG.ORG_LABEL` and the new signature builder. Matches the `{{wingName}}` convention
  already used by `TransitionCompleteEmail.html`.
- **New optional Script Property `TENANT_DIRECTOR_RECRUITING_NAME`** (`CONFIG`-level const
  `DIRECTOR_RECRUITING_NAME`) carries the signature name. Like the director's *address* it
  names an individual, so it is deliberately **not** version-controlled in `config-tenants/`.
  **Blank is valid** and signs with the office title alone — the right default for a tenant
  that has not named a role holder, and for one that would rather not put a personal name on
  automated mail.
- **`retentionRenderTemplate_()`** replaces the `.replace()` chains that were duplicated
  across all seven render sites (three send paths, four test paths), so a new placeholder is
  added in one place. It substitutes via a **replacer function**, not a string, so a value
  containing `$&` or `` $` `` cannot corrupt the output; unknown placeholders are left intact
  so a typo is visible in the mail instead of silently blanking.

### Removed

- **The feedback survey from `ExpiringEmail.html`.** It shipped with the literal placeholder
  `LINK TO FORM HERE` in three places — the button `href`, the fallback `href`, **and** the
  visible link text — so arming the trigger would have mailed every expiring member (the
  largest of the three categories) a dead button and a paste-this-URL block reading
  "LINK TO FORM HERE". The feedback ask now routes to `replyTo`, which is already the
  Director of Recruiting.
- **The commented-out MIWG "Phoenix Senior Flight" block** from `Turning21Email.html`,
  upstream content for a unit this deployment does not have.
- CSS orphaned by both removals (`.button`, `.button-container`, `.survey-link`,
  `.highlight-box`).

### Fixed

- **Unbalanced `<p>` tags** in all three signature blocks (a `<p>` opened for the name line
  was never closed before the next opened).

### Not changed — still required before the monthly trigger is armed

The verification pass turned up operational blockers that are **not** code fixes:

1. **`TENANT_DIRECTOR_RECRUITING_EMAIL` must be set** on whichever project is armed. Every
   send passes it as `replyTo`; blank makes Gmail reject **every** send, which the per-member
   catch turns into a silent 100%-failure run.
2. **The trigger must be created while signed in as the automation account**, for the same
   Send-As reason that broke both notification digests on 2026-07-16. Retention fails *worse*:
   `sendRetentionSummaryEmail()` also passes `from`, so a wrong identity means no member mail
   **and** no failure summary — unlike the notification modules, which deliberately send their
   IT alarm without a `from` override.
3. **Arm on one tenant only.** This module hardcodes `'CADET'`/`'SENIOR'` and ignores
   `PROFILE_.MEMBER_TYPES_ACTIVE`, and both projects pull the same wing-wide extract
   (ORGID 188, `unitOnly=0`). Armed on both, every member gets two copies.
4. **`sendRetentionEmails()` does not download CAPWATCH.** `clearCache()` only clears the
   in-memory `_fileCache`, which is already empty at the start of any fresh execution.
   Freshness depends entirely on `getCapwatch()` having run that morning.
5. **No idempotency.** The Log sheet is written but never read back, so a timeout, a re-run,
   or a double firing re-sends to everyone already mailed.

### Notes

- CAPWATCH carries **no true date of birth** — only month and year, with the day defaulting
  to the 1st. Selecting on birth *month* is therefore the only thing the data supports, and a
  1st-of-month run lands on the recorded date. The parser already ignores `dobParts[1]`.
- Commander CC resolves to the commander's **CAPWATCH personal** address, not their CAP
  Workspace account. `RecoveryEmailNotify` prefers the org address with a CAPWATCH fallback;
  retention does not. Unchanged here, but worth a decision.

## [2026-07-23] — 2SV and never-signed-in checks join the monthly compliance digest

Commanders were told about email records that block a password reset, but not about the
two account facts sitting next to them: members whose account has **2-Step Verification
turned off**, and members whose account was **created 60+ days ago and never signed
into** — the latter silently missing every communication sent to them.

### Added (`RecoveryEmailNotify.gs` 1.1.0, `UpdateMembers.gs` 1.19.0)

- **Two new conditions in the same digest**, read from the tenant's own directory and
  joined to CAPWATCH members by CAPID:
  - **2SV not enabled** on an account that is actually in use (has been signed into).
  - **Never signed in**, 60+ days after account creation (`FIRST_LOGIN_GRACE_DAYS`);
    younger accounts wait out the grace period.

  A never-used account is flagged **only** for the sign-in, never also for 2SV — you
  cannot enroll an account you have never entered. When a CAPID holds duplicate accounts,
  the one **most recently signed into** is judged, not an abandoned twin. The digest
  gains per-issue guidance blocks (eServices self-service / 2SV enrollment link, with the
  support portal for members already locked out by 2SV enforcement / first-sign-in
  instructions), included only when the issue is present; the subject broadens from
  "email records" to "account issues".

- **Per-category suppression (state v2).** The 3-month quiet window now runs per issue
  category (`EMAIL` / `TWOSV` / `LOGIN`), so a member already inside the email window is
  still reported the month a 2SV gap appears — and the digest shows **only** the
  newly-reported issue, not a monthly re-listing of the suppressed one. v1 state files
  are **migrated in place** (their member-level date becomes the `EMAIL` category's
  date), so windows already running keep running instead of re-mailing every
  previously-reported member on deploy.

- **Directory failure aborts the run** without touching state, exactly like an empty
  roster: a failed or empty `getActiveUsers()` read would make every member look
  account-less, dropping every `TWOSV`/`LOGIN` window and re-mailing them all when the
  directory recovered. The run now also fetches the directory **once**, shared between
  the account checks and the command-staff address map.

- **`getActiveUsers()` (UpdateMembers.gs 1.19.0)** now also returns `isEnrolledIn2Sv`,
  `lastLoginTime` and `creationTime`. Purely additive — the active-only contract and
  every existing field are unchanged.

- **Never-signed-in guidance routes to the support portal (RecoveryEmailNotify.gs
  1.2.1)** — "file a support ticket at support.pcrcap.org" instead of the IT mailbox,
  matching the 2SV block. Came out of the first one-unit live test.

- **`testRecoveryDigestForOrg(orgid, recipient)` (RecoveryEmailNotify.gs 1.2.0)** —
  renders ONE unit's real digest, post-suppression (exactly the rows the next run
  would send that unit), and mails it to a test recipient only. Reads state, writes
  nothing — nobody lands on the cooldown, so the real run still reports everyone
  shown. No `from` override, so it runs from any signed-in account. The Run dropdown
  cannot pass arguments; call it from a scratch.gs wrapper.

## [2026-07-19] — Domain-aware duplicate authority + derived-address drift scan

### Changed (`DuplicateAccountGuard.gs` 1.3.0, `DuplicateAccountScan.gs` 1.5.0, `UpdateMembers.gs` 1.18.1)

- **`chooseAuthoritativeAccount_` is now domain-aware.** An account on the tenant's
  configured email domain outranks a legacy-domain twin — below login recency and
  active (never retire the account a member actually uses), above everything else.
  Motivated by the region tenant's domain-migration pair (identical localpart, legacy
  vs configured domain, both never used), where the old ordering fell through to
  "newest created" and picked correctly only by luck. All three call sites pass
  `CONFIG.EMAIL_DOMAIN`, so scan preview and cleanup keep deciding identically.

### Added (`DuplicateAccountScan.gs` 1.5.0)

- **`scanDerivedAddressDrift()`** — read-only; for every account whose CAPID matches a
  CAPWATCH member, compares the address it HAS against what provisioning would DERIVE
  today, classified. Built to settle where the odd in-use addresses came from: a
  `punctuation` drift means the CAPWATCH name lacks the punctuation the address
  carries — i.e. `baseEmail` has been deriving faithfully and the odd addresses
  predate provisioning (original population). Mismatches are stable, not defects —
  the CAPID map updates those accounts in place.

## [2026-07-19] — Rescue: cross-tenant contacts 0.2.1/0.2.2 existed only on the tenants

`CrossTenantContacts.gs` was **0.2.2 on the seniors and region tenants (byte-identical
on both) but 0.2.0 in git** — two versions of live, working code that existed nowhere in
the repo. Because `clasp push` deploys the whole `src/` tree, the next push from any
branch would have silently reverted it and re-broken the cross-tenant sync. Captured
here verbatim from the live projects so git and the tenants agree again. Cadets was
already 0.2.0 and is unaffected.

### Fixed (`CrossTenantContacts.gs` 0.2.0 → 0.2.2)

- **0.2.2 HOTFIX.** 0.2.1 had added `employeeId` to the self-directory `fields` param,
  but the Directory API has **no selectable top-level `employeeId` field**, so every
  self-directory read 400'd and aborted the entire sync. Reverted to externalId-only
  (where the Admin console's "Employee ID" actually lives). Same trap recorded in
  `ManageLicenses.gs` 1.1.0 — worth treating as a standing gotcha.
- **0.2.1 card-shape hash.** `xtHash_` keyed only on display fields, so the 0.2.0
  sort-key change (moving the sort value into `givenName`) left the hash unchanged for
  any contact whose display text didn't also change — reconcile skipped the rewrite and
  they kept sorting by first name. New `XT_CARD_FORMAT` is folded into the hash to force
  a one-time rewrite whenever the card *shape* changes.
- **0.2.1 own-domain guard.** A self-published contact whose CAPWATCH email is on one of
  this tenant's own domains (`cfg.ownDomains`) is now skipped — such a member has an
  account, and publishing it shadowed the real directory user.

> **Note on the 0.2.1 "self-publish shadow".** Its root-cause note was a misdiagnosis:
> the reported GAL double-listing was **two real Workspace accounts for one member**, not
> a self-published contact. That is the duplicate-account defect fixed in the entry below
> and cleaned up on cadets (27 duplicates retired). The own-domain guard is still correct
> as a backstop, so it is kept.

## [2026-07-19] — Monthly recovery-email compliance digest to unit command staff

A member resets their Workspace password through a personal, non-CAP address — you cannot
recover an account from the account it locks. Nothing told command staff which of their
people were set up in a way that makes that impossible.

### Added (`RecoveryEmailNotify.gs` 1.0.0, `config.gs` 1.9.0)

- **New `src/notifications/RecoveryEmailNotify.gs`.** Monthly digest to each unit's
  **commander** (To), copying its **personnel officers** (primary *and* assistant) and
  **deputy commanders**, listing members under their **direct command** (their own ORGID)
  who trip either of two independent conditions:
  - **No CAP address in the PRIMARY slot** — covers a personal address sitting in PRIMARY
    *and* no PRIMARY at all.
  - **No personal (non-CAP) address anywhere** — no way to receive a reset.

  The digest names which condition each member tripped, asks the recipient to contact the
  member, and reminds them that as command staff they may correct it themselves at
  `https://www.capnhq.gov/CAP.PersonnelInfo.Web/`.

- **A member reported once is not reported again for three months**, even if uncorrected —
  a monthly job that re-mails the same commander about the same member every month gets
  filtered to trash. State is per member (`RecoveryComplianceState.txt`), and the window
  runs from the date they were actually told, so a failed digest's retry does not reset it.
  A member who **becomes compliant is dropped from state**, so a later relapse is reported
  on the next run rather than sitting silently inside a stale window.

- **Runs on seniors *and* cadets** (`RUN_RECOVERY_EMAIL_NOTIFICATIONS`; off for region,
  a single-unit HQ). Two cadet-specific behaviours are deliberate:
  - **Cadet-lite members are excluded automatically** — they get no account, so they have
    no password to reset. This falls out of reusing `getMembers()` rather than being a
    second copy of the grade rule.
  - **A parent/guardian email counts as a recovery address.** The check reuses
    `member.recoveryEmail` from `UpdateMembers.gs`, which is the same value that populates
    the account's real Workspace recovery address — so the module reports on exactly what
    recovery will use, rather than a second, subtly different opinion of it. Flagging
    cadets covered by a parent address would bury the real gaps under most of the cadet wing.

- **`CONFIG.COMMAND_EMAIL_DOMAIN`** (optional Script Property `TENANT_COMMAND_EMAIL_DOMAIN`,
  blank defaults to `EMAIL_DOMAIN`). Command staff are **senior** members, so on the cadets
  tenant their account is on the senior domain; deriving `first.last@cawgcadets.org` for a
  cadet unit's commander would address an account that does not exist. **Required on the
  cadets tenant** — `validateTenantConfig()` does not flag it, because blank is correct
  everywhere else.

### Notes

- **Unlike `LSCodeNotify.gs`, the first run is deliberately loud.** That module diffs against
  recorded state and is silent on first run by construction; this one reports a *standing*
  condition, so the first run surfaces the whole existing backlog. Run
  `previewRecoveryEmailCompliance()` (sends nothing, writes nothing) and check the volume
  before scheduling.
- Recipient addresses resolve **directory → derived → CAPWATCH primary**. CAPWATCH primary is
  last deliberately, since a wrong or personal CAPWATCH primary is the very thing this module
  exists to report. On a cadet tenant the local directory is skipped (command staff are seniors
  in the peer tenant), so addresses there are derived and unverified.
- **Interacts with the duplicate-account work below, which landed the same day.** The command
  directory map is built from `getActiveUsers()`, which is active-only — so where a member has
  a duplicate pair, this addresses whichever twin is *not* suspended. That is why the first
  live run correctly reached a `.N`-suffixed commander account rather than the canonical
  `first.last` one. Once `suspendOrphanDuplicates()` retires the orphans, the map resolves to
  the surviving account on the next run; the map is rebuilt every run, so nothing is pinned.

### Fixed (`RecoveryEmailNotify.gs` 1.0.1) — both found by the first live preview

- **Derivation alone produced dead addresses.** `first.last@<domain>` reproduces only the
  *default* account name; `addOrUpdateUser` prefers the real directory address when an account
  is not the default. The live wing turned up five classes that would have been silent
  dead-letter sends: an apostrophe stripped in the real account, a **`.3` duplicate-account
  suffix**, a CAPWATCH nickname vs the legal first name, a surname changed since account
  creation, and a middle name concatenated into the CAPWATCH first name. Recipients are now
  resolved from this tenant's directory first, and only derived when it has no entry — which
  is always the case on a cadet tenant, and why that path is retained rather than replaced.
- **The preview misreported recipients.** It printed the raw duty list, so a unit where one
  person holds several of these duties showed them repeatedly (four times, in one case) — while
  the send had always deduplicated. The preview is the surface an operator checks recipients
  on, so it now runs the same reduction as the send and reports the addressee and Cc separately.
- Members merged from the **ManualMembers sheet are skipped**: they never pass through
  `addContactInfo()`, so read naively every one of them would look non-compliant.
- Own state file and own trigger, never touching provisioning — same isolation rationale as
  `LSCodeNotify.gs`. The monthly trigger **must be created as the automation account**
  (Send-As identity); the IT failure-summary sends without a `from` override so that
  misconfiguration still gets reported.
- Verification: `npm test` — 2 new suites (118 assertions), mutation-checked (disabling the
  suppression window fails 9 assertions across both files; removing the manual-member guard
  fails 3).
## [2026-07-19] — Stop provisioning from creating duplicate Workspace accounts

A member (and, it turned out, a wider population) held two active accounts, neither
ever signed in — one in `last.first` order (a format the code never generates) and
one in `first.last` order (the derived format). Root cause: `addOrUpdateUser`
decides update-vs-create purely on whether `Users.update` succeeds at the derived
`first.last` email, and the CAPID→email map it consults (`getActiveUsers` →
`workspaceEmailByCapid`) omits **suspended** accounts and reads the CAPID **only**
from `externalIds[type='organization']`. Any member whose real account was suspended,
tagged under a different externalId type, or created out-of-band (e.g. by hand in
last.first order) was invisible to the map, so provisioning **inserted a second
account** instead of updating the first. There was no rename path and no directory
lookup by CAPID before inserting.

A scan of the cadets tenant found **28 CAPIDs holding 56 accounts** out of 1,189 users
(plus 12 accounts carrying no readable CAPID). Two bulk creation events: an import on
2025-11-24 and a provisioning run on 2026-01-23 that mass-created canonical twins for
members it could not match by CAPID. The scan corrected two assumptions:

- **The canonically-named account is usually the DEAD one.** In nearly every group the
  older, oddly-named account carries the login history and the newer `first.last` twin
  has never been signed into. Authoritative selection therefore ranks **login history
  first**, not name shape — the reverse would retire accounts in active use.
- **The dominant trigger is punctuation/collision drift, not name swaps.** Roughly a
  third are `.N` collision suffixes on the in-use account (our code suffixes only
  aliases, never primary emails, so those came from the import); another third are
  hyphen drift between the localpart and the derived address (`baseEmail` strips
  whitespace but **not** hyphens); the rest are typo/preferred-name corrections. A
  true `last.first` → `first.last` reversal was a single group.

### Added (`UpdateMembers.gs` 1.17.0–1.18.0, `DuplicateAccountGuard.gs` 1.1.0, `DuplicateAccountScan.gs` 1.1.0)

- **Duplicate-create guard.** Before the create branch inserts, `addOrUpdateUser`
  now calls `findExistingAccountsByCapid_()` — a live directory lookup by CAPID that
  **does** see suspended accounts and every externalId type. If a real account for
  the member exists, it is updated in place at its own address and no second account
  is created. `chooseAuthoritativeAccount_()` picks the right twin when more than one
  already exists (canonical `first.last` > active > unsuffixed > newest). Both are
  pure and unit-tested (`test/DuplicateAccountGuard.test.js`).
- **Read-only scanner** `scanDuplicateAccountsByCapid()` groups the directory by
  CAPID (reading the org externalId AND a top-level employeeId) and reports every
  CAPID with >1 account: emails, created dates, suspended/never-signed-in status, and
  whether localparts are reversed (`first.last` vs `last.first`) or `.N`
  collision-suffixed. Uses only `admin.directory.user.readonly` (already scoped).
- **Gated cleanup** `suspendOrphanDuplicates(dryRun=true)` retires the extra accounts
  that already exist: it retypes the orphan's `organization` externalId to a
  `duplicate_retired_capid` marker **and** suspends it. The retype is what makes the
  suspension stick — `reactivateRenewedMembers()` un-suspends any suspended user whose
  CAPID is active (reading the org externalId), so a plain suspend would be reversed on
  the next reactivation run. **Nothing is deleted** (permanent on this edition — no
  Archived-User licenses); orphans that have login history are skipped for a human.
  Defaults to a **dry run**; reversible by changing the externalId type back.
- **Provisioning CAPID map** `buildProvisioningEmailByCapid_()` now backs
  `workspaceEmailByCapid` in `updateAllMembers()`/`forceUpdateAllMembers()`. It sees
  suspended accounts and every CAPID carrier, and resolves a CAPID with several
  accounts to the one the member signs into. This is what makes a retirement stick:
  suspending a dead twin alone does nothing, because provisioning still derives that
  twin's address, `Users.update` **succeeds** against it, and the 404-triggered guard
  never fires — so it would be unsuspended and re-maintained every run.
- `getActiveUsers()` is left unchanged on purpose — `suspendExpiredMembers()` and
  `ManageLicenses.gs` depend on its active-only contract, so provisioning got its own
  map builder rather than a widened one.

### Verified against the live tenant

The first live scan confirmed the guard covers the whole population: **0** of the 56
duplicate accounts are invisible to the `externalId=<capid>` lookup — the carrier
histogram is uniformly `externalId:organization`. So the 2025-11-24 import *was*
CAPID-tagged, and the 2026-01-23 run duplicated it only because the old map filtered
suspended users and never saw those accounts.

Reviewing that scan also caught two defects, fixed in `DuplicateAccountGuard.gs` 1.2.0:

- **Login recency, not a boolean.** One pair has BOTH accounts signed into — one used
  days ago, one months earlier. A has-ever-signed-in boolean tied them, so "newest
  created" won and marked the **actively used** account for retirement. Ranking is now
  on the `lastLogin` timestamp.
- **Preview must equal action.** `suspendOrphanDuplicates` re-ranked with a canonical
  `first.last` drawn from CAPWATCH that the scan's preview does not use, so the account
  an admin reviewed as KEEP could differ from the one cleanup kept. Cleanup now consumes
  the scan's own decision.

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
