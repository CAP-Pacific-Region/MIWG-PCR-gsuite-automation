# `signature-webapp/` — member self-service email signature

**This is a separate Apps Script project, one per tenant.** It is not part of `src/` and is never
deployed by `npm run push:seniors`. The same source is pushed to its own script project inside each
Workspace — a project lives in one tenant, and a cadet cannot sign in to one on the senior domain —
each with its own manifest, Script Properties, OAuth scopes and clasp target
([seniors](../clasp-targets/signature-webapp-seniors.clasp.json) ·
[cadets](../clasp-targets/signature-webapp-cadets.clasp.json)).

```bash
npm run push:signature:seniors     # deploy to the seniors tenant's project
npm run push:signature:cadets      # ...and the cadets tenant's
npm run open:signature:seniors     # open it in the Apps Script editor
```

Nothing in this code branches on seniors-vs-cadets. What differs for a cadet — never a phone row,
duties out of `CadetDutyPositions.txt`, no secondary domain — differs because their CAPWATCH record
and their tenant's Script Properties differ.

A member opens the link, signs in as themselves, sees the signature CAPWATCH says they should
have, and — if they approve it — has it written to their own CAP addresses. Two things are theirs
to decide: whether the phone row appears, and which of their own duty assignments do (up to two).
Not what any of it says, and not the order it prints in.

Setup, the access model, and what happens when `pushAllSignatures()` in `src/` is run over the
same member are documented in
**[docs/SIGNATURE_WEB_APP.md](../docs/SIGNATURE_WEB_APP.md)** — read that before changing
anything here.

| File | Role |
|---|---|
| `appsscript.json` | Manifest. `access: DOMAIN` is load-bearing — see `Auth.gs`. |
| `Config.gs` | Script-Property config + the `Logger` shim the whole codebase uses. |
| `Auth.gs` | Who is calling, and whether they may. The account it returns is the only one anything else may touch. |
| `MemberRecord.gs` | CAPID → CAPWATCH record, read-only, built the way `src/` builds it. |
| `SignatureTemplate.gs` | The CAP signature. A port of the generator in `src/` — output must match byte for byte. |
| `GmailSignature.gs` | The only writes: the caller's own org-owned Send-As identities. |
| `SignatureApi.gs` | `doGet` + the server functions the browser may call. |
| `Index.html` | The UI. Holds no authority; the account comes from the session, server-side. |

Tests: [`test/SignatureWebApp.test.js`](../test/SignatureWebApp.test.js), run by `npm test` with
the rest of the suite. Section 1 is the one that matters most — it renders every fixture member
through **both** this project's generator and `src/`'s and fails on any difference.
