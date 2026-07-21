# `webapp/` — Secondary Alias admin web app

**This is a separate Apps Script project.** It is not part of `src/` and is never deployed
by `npm run push:seniors`. It has its own clasp target
([`clasp-targets/alias-webapp.clasp.json`](../clasp-targets/alias-webapp.clasp.json)) and its
own manifest, Script Properties and OAuth scopes.

```bash
npm run push:webapp     # deploy this project
npm run open:webapp     # open it in the Apps Script editor
```

Setup, the access model, and the sheet contract it shares with the nightly
`addSecondaryDomainAliases()` run are documented in
**[docs/ALIAS_WEB_APP.md](../docs/ALIAS_WEB_APP.md)** — read that before changing anything
here.

| File | Role |
|---|---|
| `appsscript.json` | Manifest. `access: DOMAIN` is load-bearing — see `Auth.gs`. |
| `Config.gs` | Script-Property config + the `Logger` shim the whole codebase uses. |
| `Auth.gs` | The entire access control. Fails closed. |
| `Directory.gs` | CAPID → account resolution, address arithmetic, and the only two calls that change an account. |
| `AliasAdminApi.gs` | `doGet` + the server functions the browser may call. |
| `Index.html` | The UI. Holds no authority; everything is re-checked server-side. |

Tests: [`test/AliasWebApp.test.js`](../test/AliasWebApp.test.js), run by `npm test` with the
rest of the suite.
