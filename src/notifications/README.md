# notifications/

Push signals to unit command staff, derived from the CAPWATCH extract.

Both modules follow the same shape, deliberately: **own Drive state file, own trigger, no
contact with provisioning.** A failure here cannot affect account creation, and a
`forceUpdateAllMembers()` cannot swallow a pending notification. Neither reuses
`CurrentMembers.txt` — that snapshot belongs to `updateAllMembers()`, and sharing it would
mean a throw mid-run either re-mails every commander or silently drops transitions.

| Module | Cadence | Tenants | Profile flag |
|--------|---------|---------|--------------|
| `LSCodeNotify.gs` | Weekly (Mon) | seniors | `RUN_LSCODE_NOTIFICATIONS` |
| `RecoveryEmailNotify.gs` | Monthly (1st) | seniors + cadets | `RUN_RECOVERY_EMAIL_NOTIFICATIONS` |

Each module's header is the authoritative documentation — read it before changing behaviour.
Operational steps live in [Admin Guide §8–9](../../docs/ADMIN_GUIDE.md#8-what-runs-when-the-automation-schedule).

## LSCodeNotify.gs — background-check changes

Mails a unit's commander when a member's FBI background check (`Member.txt` `LSCode`) is
granted or lapses. **Seniors only**: cadet records carry no `LSCode` at all, including cadets
over 18, so the column tracks the senior-side record rather than the person.

**Change detector.** A member absent from state is recorded silently, so the first run — and
every new member after it — mails nobody. The digest dates a **window**, not a day, because
CAPWATCH publishes no date for an `LSCode` change; changing the trigger interval therefore
changes what the digest claims.

## RecoveryEmailNotify.gs — password-reset readiness

Mails a unit's commander (copying personnel officers, primary and assistant, and deputy
commanders) about members under their **direct command** whose CAPWATCH email record would
stop them resetting their Workspace password:

- no CAP address in the **PRIMARY** slot (a personal address there, or no primary at all), and/or
- no personal, non-CAP address **anywhere** to receive a reset at.

**Standing condition, not a change detector.** The first run is meant to be loud — it surfaces
the whole existing backlog. Run `previewRecoveryEmailCompliance()` and check the volume first.
A member reported once is quiet for **three months**; a member who fixes it drops out of state,
so a relapse is reported immediately rather than inside a stale window.

Reuses `member.recoveryEmail` from `UpdateMembers.gs` rather than re-deriving it, so it reports
on exactly the value that populates the real Workspace recovery address. Two consequences are
intended: cadet-lite members are excluded (no account, no password), and a cadet covered by a
**parent/guardian** address is not chased.

⚠️ On the **cadets** tenant, set `TENANT_COMMAND_EMAIL_DOMAIN` to the senior domain. Command
staff are seniors and hold no cadet-domain account.

## Sending identity — read before scheduling either

Digests send as `AUTOMATION_SENDER_EMAIL`, which Gmail permits only when the **executing**
account owns that verified Send-As alias. A time-driven trigger runs as whoever created it, so
**both triggers must be installed while signed in as the automation account.** Running as
`it@` failed every digest with `Invalid argument` on 2026-07-16.

Both IT failure-summaries deliberately send **without** a `from` override, so the alarm still
arrives when that identity is wrong — an "attention needed" email listing every unit as failed
is the signature of exactly this misconfiguration.

## Tests

`npm test` runs both modules under Node with faked Google globals
([`test/helpers/apps-script.js`](../../test/helpers/apps-script.js)). What is verified is the
modules' own decisions — who gets mailed and when — not that Google's APIs behave.
`makeClock()` moves "today" so the three-month window can be tested without waiting.
