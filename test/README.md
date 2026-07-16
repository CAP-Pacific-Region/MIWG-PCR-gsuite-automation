# Tests

```bash
npm test
```

No dependencies, no framework, no install step — plain Node, so it runs on a clean
clone.

## Why these exist

Apps Script has no local test runner. The only way to run a `.gs` file for real is
to push it to a live tenant and press **Run**, which for this codebase means
touching production Drive, Gmail and Workspace. That is a bad place to find out
that a notification module mails the entire wing, or that a deletion path
misreads a column.

A `.gs` file is just JavaScript whose globals happen to be injected by the
platform. So these tests read the source, inject fakes for the globals it reaches
for (`DriveApp`, `GmailApp`, `Logger`, `Utilities`, `CONFIG`, `PROFILE_`, …), and
run the module's own logic unmodified under Node.

**What is verified is our code. What is faked is Google.** These tests prove the
decisions — who gets mailed, and when — not that Google's APIs behave. They are
not a substitute for a dry run against a real tenant; they are what makes that
dry run the *second* check rather than the first.

## Layout

| Path | Purpose |
|---|---|
| `run-all.js` | Runs every `*.test.js`, one child process each |
| `helpers/apps-script.js` | The stubs: `loadModule`, fake Drive/Gmail/Logger, `Utilities`, assertions |
| `helpers/capwatch-fixtures.js` | `Member.txt` / `Commanders.txt` fixtures, real headers |
| `LSCodeNotify.run.test.js` | End-to-end through `notifyLSCodeChanges()` |
| `LSCodeNotify.diff.test.js` | The diff, state rollback, and window rendering in isolation |

Nothing here ships to Apps Script: clasp's `rootDir` is `../src` and
`.claspignore` excludes everything outside `src/`, so `test/` never leaves the
repo.

## Writing a new one

`loadModule(srcPath, globals, exportNames)` returns the named functions from a
`.gs` file with `globals` injected. Private helpers (trailing `_`) are reachable
this way, since the module is evaluated as one scope.

```js
const { loadModule, makeLogger, makeDrive, makeGmail, makeChecker } =
  require('./helpers/apps-script');

const { check, done } = makeChecker();
const { gmail, sent } = makeGmail();
const files = { 'Member.txt': '...' };

const m = loadModule(MODULE, {
  Logger: makeLogger().logger,
  DriveApp: makeDrive(files),
  GmailApp: gmail,
  /* ...only the globals your module actually touches... */
}, ['someFunction_']);

check('sends nothing', sent.length, 0);
done();
```

Two conventions worth keeping:

- **Fixture headers must match the live extract.** `capwatch-fixtures.js` copies
  the real `Member.txt` and `Commanders.txt` headers verbatim. Modules resolve
  CAPWATCH columns by header name, so a fixture with an invented header would
  pass regardless of what the module reads, and prove nothing.
- **Stubs should throw on anything unhandled** rather than return something
  plausible. `Utilities.formatDate` throws on an unknown pattern for this reason:
  a stub that guesses turns a real bug into a passing test.
