# admin-webapp/

A **separate Apps Script project** — the domain admin help-desk page. It is not
deployed by `npm run push:seniors`; it has its own manifest, its own scopes, and
its own clasp target — one per tenant. It runs on both CAWG tenants, each
project linked to the other as its peer.

```bash
npm run push:admin:seniors     # code only — NEVER deploy from clasp
npm run open:admin:seniors
npm run push:admin:cadets
npm run open:admin:cadets
```

Full setup, access model and troubleshooting: **[docs/ADMIN_WEB_APP.md](../docs/ADMIN_WEB_APP.md)**.

## Files

| File | Holds |
| --- | --- |
| `AdminApi.gs` | `doGet` and the complete list of functions the browser can reach. **Every one starts with `requireAdmin_()`** |
| `Auth.gs` | Who may call (super admin / Help Desk role) and whom they may act on (not another admin) |
| `Config.gs` | Script Properties, and why each one is optional or not |
| `Accounts.gs` | Directory lookups — accounts by CAPID, the authoritative pick, managed-group membership |
| `MemberRecord.gs` | One member out of the CAPWATCH extract, read-only |
| `Actions.gs` | The four things this app can do, and every guard deciding whether it may |
| `Credentials.gs` | Temporary passwords, the welcome email, the shared audit ledger |
| `WelcomeEmail.html` | Byte-for-byte copy of `src/recruiting-and-retention/WelcomeEmail.html` |
| `Setup.gs` | One-time, run-by-hand Script Property setup per tenant (`setupSeniorsAdminWebApp()` / `setupCadetsAdminWebApp()`), sourced from `config-tenants/<tenant>.json` |

> `Credentials.gs` would rather be called `WelcomeEmail.gs`. It cannot: Apps Script
> addresses files without their extension, so it would collide with
> `WelcomeEmail.html` and the push is rejected.
| `AuditLog.gs` | One row per action, naming the admin — the platform log names only the deployer |
| `Index.html` | The page. Holds no authority; every check it appears to make is re-made server-side |

## Before you change anything here

Four things in this project are **copies of `src/`**: the welcome email template,
the resend eligibility policy, the authoritative-account ranking, and the welcome
ledger's format. `test/AdminWebApp.test.js` runs both copies over the same inputs
and fails on divergence — read its header, which explains what each copy breaks
if it drifts.

```bash
npm test
```
