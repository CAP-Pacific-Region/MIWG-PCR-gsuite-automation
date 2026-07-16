/**
 * Minimal Apps Script stubs, so .gs modules can be exercised under plain Node.
 *
 * WHY THIS EXISTS
 * Apps Script has no local test runner: the only way to run a .gs file for real
 * is to push it to a live tenant and press Run, which for this codebase means
 * touching production Drive, Gmail and Workspace. That is a bad place to discover
 * that a notification module mails the entire wing.
 *
 * A .gs file is just JavaScript whose globals happen to be injected by the
 * platform. So load the source, inject fakes for the globals it reaches for, and
 * the module's real logic runs unmodified under Node. What is verified is the
 * module's own code; what is faked is Google. That boundary is the point — these
 * tests prove the decisions (who gets mailed, and when), not that Google's APIs
 * behave.
 *
 * Nothing here is pushed to Apps Script: clasp's rootDir is ../src and
 * .claspignore excludes everything outside src/, so test/ never leaves the repo.
 */
const fs = require('fs');

/**
 * Loads a .gs module with the given globals injected, returning the named
 * functions from inside it.
 *
 * @param {string} srcPath - Absolute path to the .gs file
 * @param {Object} globals - Global name -> value to inject (Logger, DriveApp, ...)
 * @param {Array<string>} exportNames - Function names to return from the module
 * @returns {Object} Map of exportName -> function
 */
function loadModule(srcPath, globals, exportNames) {
  const names = Object.keys(globals);
  const factory = new Function(
    ...names,
    fs.readFileSync(srcPath, 'utf8') + '\n; return { ' + exportNames.join(', ') + ' };'
  );
  return factory(...names.map(n => globals[n]));
}

/** A Logger that records rather than prints. */
function makeLogger() {
  const calls = { info: [], warn: [], error: [] };
  return {
    logger: {
      info: (msg, ctx) => calls.info.push({ msg, ctx }),
      warn: (msg, ctx) => calls.warn.push({ msg, ctx }),
      error: (msg, ctx) => calls.error.push({ msg, ctx })
    },
    calls
  };
}

const Session = { getScriptTimeZone: () => 'America/Los_Angeles' };

/**
 * Parses CSV the way Utilities.parseCsv does for CAPWATCH files: comma
 * separated, double-quoted fields, "" as an escaped quote.
 */
function parseCsv(text) {
  return text.split(/\r?\n/).filter(line => line.length).map(line => {
    const out = [];
    let cur = '';
    let quoted = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (quoted) {
        if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
        else if (ch === '"') quoted = false;
        else cur += ch;
      } else if (ch === '"') quoted = true;
      else if (ch === ',') { out.push(cur); cur = ''; }
      else cur += ch;
    }
    out.push(cur);
    return out;
  });
}

const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

/**
 * Only the date patterns this codebase actually uses. Unknown patterns throw
 * rather than returning something plausible — a silently wrong date in a test
 * stub would make a real formatting bug look like a pass.
 */
const Utilities = {
  parseCsv: parseCsv,
  sleep: () => {},
  formatDate: (d, tz, pattern) => {
    const pad = n => String(n).padStart(2, '0');
    switch (pattern) {
      case 'yyyy-MM-dd': return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
      case 'yyyy': return String(d.getFullYear());
      case 'MM': return pad(d.getMonth() + 1);
      case 'd': return String(d.getDate());
      case 'd MMM': return `${d.getDate()} ${MONTHS[d.getMonth()]}`;
      case 'd MMM yyyy': return `${d.getDate()} ${MONTHS[d.getMonth()]} ${d.getFullYear()}`;
      default: throw new Error('apps-script stub: unhandled formatDate pattern ' + pattern);
    }
  }
};

/**
 * An in-memory DriveApp over a plain { filename: content } object. Mutations go
 * straight back into that object, so a test can assert on what was written.
 *
 * @param {Object} files - filename -> string content
 * @returns {Object} DriveApp-shaped stub
 */
function makeDrive(files) {
  return {
    getFolderById: () => ({
      getFilesByName: name => {
        const exists = Object.prototype.hasOwnProperty.call(files, name);
        let consumed = false;
        return {
          hasNext: () => exists && !consumed,
          next: () => {
            consumed = true;
            return {
              getBlob: () => ({ getDataAsString: () => files[name] }),
              setContent: content => { files[name] = content; },
              setTrashed: () => { delete files[name]; }
            };
          }
        };
      },
      createFile: (name, content) => { files[name] = content; }
    })
  };
}

/**
 * A GmailApp stub that records every send instead of sending. `failAddresses`
 * (optional) is a set of `from` addresses Gmail should reject with the same
 * "Invalid argument" error the live tenant raises for an unowned Send-As alias —
 * so a test can reproduce the wrong-identity failure.
 *
 * @param {Object} [opts] - { failFrom: string[] }
 * @returns {Object} { gmail, sent } where `sent` accumulates { to, subject, html, from }
 */
function makeGmail(opts) {
  const failFrom = (opts && opts.failFrom) || [];
  const sent = [];
  return {
    gmail: {
      sendEmail: (to, subject, body, options) => {
        const from = options && options.from;
        if (from && failFrom.indexOf(from) !== -1) {
          throw new Error('Invalid argument: ' + from);
        }
        sent.push({ to, subject, body, html: options && options.htmlBody, from: from });
      }
    },
    sent
  };
}

/** Tiny assertion helper. Returns { check, failures, section }. */
function makeChecker() {
  const state = { failures: 0 };
  return {
    section: title => console.log('\n' + title),
    check: (name, actual, expected) => {
      const a = JSON.stringify(actual);
      const e = JSON.stringify(expected);
      if (a === e) {
        console.log(`  PASS  ${name}`);
      } else {
        state.failures++;
        console.log(`  FAIL  ${name}\n        expected ${e}\n        actual   ${a}`);
      }
    },
    done: () => {
      console.log(state.failures === 0
        ? '\nAll assertions passed.\n'
        : `\n${state.failures} assertion(s) FAILED.\n`);
      process.exit(state.failures === 0 ? 0 : 1);
    }
  };
}

module.exports = {
  loadModule, makeLogger, makeDrive, makeGmail, makeChecker,
  Session, Utilities, parseCsv
};
