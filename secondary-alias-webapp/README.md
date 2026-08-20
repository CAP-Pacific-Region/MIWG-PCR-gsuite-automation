# `secondary-alias-webapp/` — the ".gov Account Processing" web app

A domain-restricted web app that lets an authorized person add or remove a member on the
`Secondary Aliases` tab **by CAPID**. On a genuine add it does the **whole thing
immediately** — creates the directory alias, turns on Gmail *Send mail as*, and emails the
member — so nobody waits for the nightly `addSecondaryDomainAliases()` run.

**This is its own Apps Script project**, separate from `src/` (which spends its one `doGet`
on the FileMaker webhook) and version-controlled here after being pulled from the live
project. Manage it with:

```bash
npm run status:alias:seniors   # what a push would send
npm run pull:alias:seniors     # pull live source back into this dir
npm run push:alias:seniors     # push HEAD (does NOT change the live @1 deployment)
npm run open:alias:seniors     # open in the Apps Script editor
```

| File | Role |
|---|---|
| `AliasAdminApi.js` | `doGet` + the server functions the browser calls. `apiAddByCapid` runs the immediate process. |
| `Directory.js` | CAPID → account resolution, address arithmetic, alias add/remove. |
| `Notify.js` | **Immediate Send-As + welcome email** — ported from `SecondaryDomainAliases.gs` 1.4.0. |
| `Auth.js` | The entire access control (group membership, fails closed). |
| `Config.js` | Script-Property config + the `Logger` shim. |
| `Index.html` | The UI. |
| `SecondaryAliasWelcomeEmail.html` | The member email (byte-identical copy of the src template). |

## Why the immediate process lives here, not in the nightly run

The nightly `addSecondaryDomainAliases()` only configures Send-As and emails on a **genuine
insert**. An alias this page creates is `already present` by the time the trigger fires — so
the trigger would notify **nobody**. The page therefore does it itself, at the click.
`Notify.js` is a hand copy of the src logic; `test/SecondaryAliasWebAppNotify.test.js` pins
the two together, including a render-parity check and that the email template matches src
byte-for-byte.

## Deploying a change (READ THIS — the live deployment is pinned)

The live deployment `AKfycbz5mBOJ…` ( *.gov Account Processing* ) serves **version `@1`**, so
`clasp push` alone updates only HEAD and the live page is unchanged. To make a change live:

1. **Set the Script Properties** the new behavior needs (Project Settings › Script
   Properties). Copy values from `config-tenants/seniors.json` where noted:

   | Property | For | Notes |
   |---|---|---|
   | `TENANT_EMAIL_DOMAIN`, `TENANT_SECONDARY_EMAIL_DOMAIN`, `TENANT_AUTOMATION_SPREADSHEET_ID`, `WEBAPP_ALIAS_ADMIN_GROUP`, `TENANT_WING_ABBREVIATION` | (already set — the original app) | |
   | `SA_IMPERSONATION_EMAIL`, `SA_PRIVATE_KEY` | **auto Send-As** | the tenant's delegated service account (the secret key; same values `src/` uses) |
   | `TENANT_AUTOMATION_SENDER_EMAIL` | the welcome email's From | without it, no email is sent |
   | `TENANT_ITSUPPORT_EMAIL` | the email's Reply-To | |
   | `TENANT_WING_NAME` | email copy (e.g. `California Wing`) | |
   | `TENANT_SENDER_NAME` | optional From name | defaults to `CAP Information Technology` |
   | `TENANT_ITSUPPORT_URL` | optional email footer link | defaults to `https://support.pcrcap.org` |

2. **Re-authorize the added scopes.** v1.1 adds `gmail.settings.sharing`,
   `gmail.settings.basic`, `gmail.send`, `script.send_mail`, and `script.external_request`.
   Open the editor and run any function once to trigger the consent screen, and grant them as
   the deploying account (a super admin, so impersonation works).

3. **Redeploy the version.** Deploy › Manage deployments › edit *.gov Account Processing* ›
   Version **New version** › Deploy. This points the live URL at the new code.

Until steps 1–2 are done the app still works — a new alias is still created — but Send-As
falls back to the manual steps in the email, and if `TENANT_AUTOMATION_SENDER_EMAIL` is unset
no email goes out. Everything in `Notify.js` is best-effort and never blocks the alias.
