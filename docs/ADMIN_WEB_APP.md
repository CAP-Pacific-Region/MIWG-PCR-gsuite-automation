# Domain admin help-desk web app

`admin-webapp/` — a page where wing IT staff run the per-member account operations
that previously required someone with access to the Apps Script editor.

---

## 1. Why it exists

Wing IT volunteers hold Google's **Help Desk Administrator** role. That role lets
them reset a member's password from the Admin console, and nothing else this
wing's automation does. Everything past a password reset — resending a welcome
email to an account created out-of-band, parking someone in the 2SV setup group
while they enroll, finding out *which* of a member's two accounts is the live one
— meant asking the one person with editor access to run a function by hand.

That is a bottleneck made of one person, and the editor is not somewhere a
help-desk volunteer belongs: every function in `src/` is one mis-click from a
wing-wide run.

This app exposes four operations, each scoped to a single member, each recorded
against the name of the admin who ran it.

| Action | What it actually does |
| --- | --- |
| **Look up a member** | CAPWATCH record + every Workspace account carrying that CAPID, with suspension, 2SV, last sign-in, org unit, admin flags, managed-group membership and the welcome-email ledger verdict |
| **Reset password** | New temporary password, `changePasswordAtNextLogin` set. Shown on screen to read out; optionally mailed to the member's eServices address |
| **Resend welcome email** | Password reset + the wing's welcome email, guarded by the same policy as `resendWelcomeEmail()` in `src/` |
| **Group add / remove** | Membership of the 2SV setup group, and any other group explicitly configured as managed |

## 2. Why it is a fourth Apps Script project

An Apps Script project has exactly one `doGet`, and the three that exist are
spoken for: `src/` on the FileMaker mission-provisioning webhook (anonymous),
`signature-webapp/` on member self-service, and now this.

Separation also keeps the blast radius honest. This project's OAuth scopes are
the ones these four actions need and no more — no calendar, no chat, no shared
contacts, no group *settings*. A bug here cannot reach them.

The cost is duplication, and it is paid deliberately: the welcome email template,
the resend eligibility policy, the CAPID account ranking and the audit ledger
format are all copies of `src/`. Every one of them is pinned by
`test/AdminWebApp.test.js`, which runs both copies over the same inputs and fails
on divergence. **Read that file's header before changing either side.**

## 3. Access and trust model

The app is deployed `executeAs: USER_DEPLOYING`, `access: DOMAIN`. Every
directory write therefore runs with the **deployer's** rights, not the caller's —
which is the entire point, since the people who use it cannot reach the editor.

That means the deployment grants nothing by itself and `Auth.gs` *is* the access
control. Two gates:

**Who may call — `requireAdmin_()`.** A super admin, or the holder of a directory
role named in `WEBAPP_ADMIN_ROLES` (default `_HELP_DESK_ADMIN_ROLE`, Google's
built-in Help Desk Administrator). Roles rather than a group of names, because
the Admin console already records this and a second list would drift: a volunteer
who loses the role loses the app the same day, with nobody remembering to edit
anything.

**Whom they may act on — `admAssertMayActOn_()`.** A caller who is not a super
admin may not act on **another admin's account**. This mirrors Google's own rule
and it is the difference between a delegated tool and a privilege-escalation
ladder — without it, a help-desk volunteer could reset a super admin's password
here and then sign in as them.

Two further boundaries worth knowing:

- **The managed-group list is a security boundary.** The browser names a group in
  every membership request and `admGroupFromRequest_()` refuses anything not on
  the configured list. Without it, this page would be a way to add yourself to
  any group in the domain.
- **Credentials are never mailed to an address on our own domains.** Not the
  account being reset, not another account on it, not the secondary domain. That
  rule is what makes a "sent" welcome email meaningful.

Every action writes one row to the audit log — success, refusal, or failure —
naming the admin, the account, and what happened. **The temporary password is
never written anywhere**: it exists in the admin's browser and, if mailed, the
member's inbox.

## 4. Setting it up on a tenant

One script project per tenant (a project lives in one Workspace). These steps are
for the seniors tenant; repeat with a new clasp target for any other.

### 4.1 Create the project and push

The seniors project already exists and its ID is in
`clasp-targets/admin-webapp-seniors.clasp.json`. For a new tenant:

```bash
clasp create-script --title "CAWG Account Admin" --type standalone
```

Paste the returned script ID into a new clasp target, then:

```bash
npm run push:admin:seniors
```

`--force` is baked into the push script: Apps Script normalizes `appsscript.json`
server-side, so clasp always sees a manifest diff and, non-interactively, answers
itself with **`Skipping push.` and exit 0** — which reads like success while
uploading nothing.

> **On filenames:** Apps Script addresses files *without* their extension, so
> `Credentials.gs` cannot be called `WelcomeEmail.gs` — it would collide with
> `WelcomeEmail.html` and the whole push is rejected with *"A file with this name
> already exists"*. Keep that in mind before renaming anything in this project.

The manifest declares the **Admin SDK** advanced service. If the page throws
`AdminDirectory is not defined` on first use, enable the Admin SDK API in the
Cloud project the script is attached to.

### 4.2 Script Properties

Copy the `TENANT_*` values straight from `config-tenants/seniors.json` — they are
the same names the main project uses.

| Property | Required | Notes |
| --- | --- | --- |
| `TENANT_EMAIL_DOMAIN` | yes | `@cawgcap.org` |
| `TENANT_DOMAIN` | yes | `cawgcap.org` — the welcome email's sign-in link |
| `TENANT_WING` | yes | `CA` |
| `TENANT_CAPWATCH_DATA_FOLDER_ID` | yes | the extract folder; also holds the welcome-email ledger |
| `TENANT_SECONDARY_EMAIL_DOMAIN` | no | `@cawg.cap.gov` on seniors |
| `TENANT_ITSUPPORT_EMAIL` | no | CC'd on credentials, and named in refusals |
| `TENANT_AUTOMATION_SPREADSHEET_ID` | no | where the audit tab is written |
| `TENANT_WING_ABBREVIATION` | no | derived from `TENANT_WING` when blank |
| `TENANT_PROFILE` | no | `seniors` (default) / `cadets` / `region` — decides which members this tenant provisions, see §5.1 |
| `WEBAPP_CADET_TOOLS_URL` | no | linked when an admin opens a member from the other tenant |
| `TENANT_CADETS_TENANT_DOMAIN` | no | named as a fallback when the URL above is unset |
| `WEBAPP_2SV_SETUP_GROUP` | no | the 2SV group's address. **Blank disables the group panel entirely** |
| `WEBAPP_MANAGED_GROUPS` | no | comma-separated extra groups the app may change |
| `WEBAPP_ADMIN_ROLES` | no | defaults to `_HELP_DESK_ADMIN_ROLE` |
| `WEBAPP_ADMIN_GROUP` | no | additive escape hatch for someone with no admin role |
| `WEBAPP_AUDIT_SHEET_NAME` | no | defaults to `Admin Web App Log` |

The audit tab is created on first use, so there is no sheet to prepare.

### 4.3 Deploy

**From the Apps Script editor, by hand** — Deploy → New deployment → Web app,
execute as **Me**, access **Anyone in <domain>**.

> 🚫 **Never deploy this from clasp.** `clasp update-deployment` /
> `create-deployment` replaces a deployment's config wholesale and the web-app
> entry point does not survive: visitors get Google's "You need access" page
> until a human redeploys from the editor. Observed twice on the signature app.
> `clasp push` is fine and is how code gets there — stop at the push.

### 4.4 Confirm the gate before handing out the URL

Sign in as a help-desk admin and confirm the page loads; sign in as an ordinary
member and confirm it refuses. The refusal is the feature.

## 5. Members who belong to the other tenant

CAPWATCH is scoped to the **wing**, not to the tenant. The seniors tenant's
extract therefore lists every cadet in the wing, and a name search finds them —
with no account here to act on, because cadet accounts live on the cadet
Workspace.

They are **flagged, not filtered out**. An admin who searched a name and got
nothing back would conclude the member does not exist; one who saw "no Workspace
account" would conclude it needs provisioning and go and create one on the wrong
tenant. So:

- the candidate list tags them **other Workspace**, with a count underneath;
- opening one replaces the usual "no account" message with a note saying the
  accounts are elsewhere and linking `WEBAPP_CADET_TOOLS_URL`;
- the same note is repeated inside the member card, because a banner is replaced
  by the next action's result and this fact is not.

In the rare case where such a member *does* hold an account here — a
cadet-to-senior transition in progress — the accounts and actions are shown as
normal, with the note saying so rather than claiming nothing can be done.

### 5.1 Which members count as ours

From `TENANT_PROFILE`, matching `MEMBER_TYPES_ACTIVE` in `src/config.gs`:

| Profile | Provisions |
| --- | --- |
| `seniors` | SENIOR, FIFTY YEAR, INDEFINITE, CADET SPONSOR |
| `cadets` | CADET |
| `region` | SENIOR, FIFTY YEAR, INDEFINITE, CADET |

Note the region tenant provisions cadets too, which is why this is a table and
not "cadets are somebody else's problem". An unrecognised type reads as *not*
ours — the worst case is an unnecessary note above a member whose accounts are
listed right below it. The table is compared against `src/config.gs` by the test.

## 6. What the actions refuse, and why

The welcome resend carries the guards from `src/WelcomeEmailResend.gs`, because a
resend is a **password reset plus a send** — the original temporary password was
never stored, so nothing else is possible.

| Refusal | Meaning | Force? |
| --- | --- | --- |
| `no-capwatch-record` | No member with that CAPID (expired, wrong wing, typo) | no |
| `no-account` | No Workspace account carries it — this is a provisioning job | no |
| `archived` | Restore the account first | no |
| `no-off-tenant-recipient` | CAPWATCH has no address outside this tenant, so credentials would be mailed into the mailbox they unlock | **no** |
| `suspended` | The new password would not work | yes |
| `already-signed-in` | The member is using this account; a resend would lock them out | yes |

The plain **password reset** deliberately carries fewer guards: an admin on the
phone with a member who has lost access is entitled to reset a working account,
which is exactly the case the resend refuses. It still refuses to act on another
admin, and it still refuses to *mail* credentials with no off-tenant address to
send to.

A successful resend records the send in `WelcomeEmailLedger.txt` — the same file
`src/accounts-and-groups/WelcomeEmailAudit.gs` reads — so the monthly audit stops
reporting that member as MISSED. This is shared state: the format, including its
version number, must match, and the test pins both.

## 7. Troubleshooting

**"This page is not set up for … yet."** A required Script Property is unset. The
names are in the execution log, not on the page — an admin cannot act on a
property name, and whoever can is reading the log anyway.

**"Your account does not hold an administrator role…"** for someone who plainly
does: check `WEBAPP_ADMIN_ROLES` against the role's real name. Built-in roles use
Google's internal names (`_HELP_DESK_ADMIN_ROLE`); custom roles use the name you
gave them. A role lookup that *fails* also denies — look for
`Role assignment lookup failed` in the log.

**The group panel is missing.** `WEBAPP_2SV_SETUP_GROUP` is unset and no
`WEBAPP_MANAGED_GROUPS` are configured. That is the fail-safe: the app will not
guess at a group address.

**"This app may only change membership of the groups it is configured for."** The
group is not on the managed list. Widening the list is an Admin console edit, on
purpose.

**A member shows two accounts.** That is the finding, not a bug — see
`scanDuplicateAccountsByCapid()` in `src/`. The one marked *authoritative* is the
one provisioning maintains: most recent sign-in first, which is usually **not**
the canonically-named `first.last` twin.

**A member has no CAPWATCH record but does have an account.** Either their
membership lapsed (the record is shown with its status when it exists) or they
are a *manual member* from the Manual Members sheet, which this app deliberately
does not read. The account card is still complete and every action still works.
