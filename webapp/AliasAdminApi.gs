/***********************************************
 * File: AliasAdminApi.gs
 * Description: Web app entry point and the server functions the browser calls.
 * Lets an authorized user add or remove a member on the "Secondary Aliases" tab
 * by CAPID, and provisions or revokes the matching directory alias to match.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-07-20
 * Changes: 1.0.0 — initial version.
 ***********************************************/

/**
 * SHEET CONTRACT WITH THE NIGHTLY MODULE — read before changing any write here.
 *
 * src/accounts-and-groups/SecondaryDomainAliases.gs owns columns C and D. It reads
 * the whole tab, builds a write-back array, and commits it POSITIONALLY in one
 * call: getRange(2, 3, dataRowCountAtReadTime, 2).setValues(...).
 *
 * That makes row POSITION a shared, unlocked data structure between two script
 * projects — and LockService locks are per-project, so the two cannot interlock.
 * If this app deleted a row while the nightly run held its array, every status
 * below the deletion would be written one row too high: alarming, misleading, and
 * silent.
 *
 * So this app never shifts rows.
 *   - Adding APPENDS (or reuses a blank row), which only ever makes the tab longer
 *     than the array — the extra row is simply left untouched by that run.
 *   - Removing CLEARS the row in place and leaves it blank. The nightly module
 *     skips blank column A and seeds its write-back with each row's current value,
 *     so a blank row stays blank.
 * Blank rows are then reused by the next add, which bounds their accumulation.
 * Deleting them by hand is safe when no run is in flight.
 *
 * Columns E and F are this app's own (CAPID, added-by). The nightly module reads
 * A/B and writes C/D only, so they are free — but they are OUTSIDE its write
 * range, which is why they must never become load-bearing for provisioning.
 */

const WEBAPP_COL_PRIMARY = 1;   // A
const WEBAPP_COL_OVERRIDE = 2;  // B
const WEBAPP_COL_STATUS = 3;    // C  (shared with the nightly module)
const WEBAPP_COL_LASTRUN = 4;   // D  (shared with the nightly module)
const WEBAPP_COL_CAPID = 5;     // E  (this app)
const WEBAPP_COL_ACTOR = 6;     // F  (this app)
const WEBAPP_COL_COUNT = 6;

/**
 * Serves the UI. Unauthorized callers get a plain refusal rather than a shell that
 * fails on first click — and the refusal names nothing they could use to escalate.
 */
function doGet() {
  const actor = resolveActor_();
  if (!isAliasAdmin_(actor)) {
    return HtmlService.createHtmlOutput(
      '<p style="font:14px/1.5 system-ui,sans-serif;padding:2rem">' +
      'You are not authorized to manage secondary aliases.<br>' +
      'Contact your wing IT director if you believe this is wrong.</p>'
    ).setTitle('Not authorized');
  }

  const t = HtmlService.createTemplateFromFile('Index');
  t.actor = actor;
  t.orgLabel = WEBAPP_CONFIG.ORG_LABEL;
  return t.evaluate()
    .setTitle(WEBAPP_CONFIG.ORG_LABEL + ' Secondary Aliases')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================================
// SERVER FUNCTIONS REACHABLE FROM THE BROWSER
// Every one of these must call requireAuthorized_() first. google.script.run can
// invoke ANY global function in the project, so an entry point that forgets the
// gate is reachable by any domain user who opens the page source.
// ============================================================================

/** Everything the page needs on load: config sanity, and the current tab contents. */
function apiGetState() {
  const actor = requireAuthorized_();
  const domain = webappNormalizeSecondaryDomain_(WEBAPP_CONFIG.SECONDARY_EMAIL_DOMAIN);
  return {
    actor: actor,
    orgLabel: WEBAPP_CONFIG.ORG_LABEL,
    secondaryDomain: domain,
    // Surfaced so a tenant that has not finished domain verification sees WHY
    // adds will fail, instead of collecting CONFLICT/ERROR rows and guessing.
    domainVerified: domain ? webappIsDomainVerified_(domain) : false,
    entries: apiListEntries_()
  };
}

/** Current tab contents, newest-looking rows last (sheet order). */
function apiListEntries() {
  requireAuthorized_();
  return apiListEntries_();
}

function apiListEntries_() {
  const sheet = webappGetAliasSheet_();
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const primary = String(webappCell_(rows[i], 0)).trim().toLowerCase();
    if (!primary) continue;   // a cleared row
    out.push({
      row: i + 1,
      primaryEmail: primary,
      aliasOverride: String(webappCell_(rows[i], 1)).trim().toLowerCase(),
      status: String(webappCell_(rows[i], 2)).trim(),
      lastRun: webappFormatDate_(webappCell_(rows[i], 3)),
      capid: String(webappCell_(rows[i], 4)).trim(),
      addedBy: String(webappCell_(rows[i], 5)).trim()
    });
  }
  return out;
}

/**
 * Looks a CAPID up in the directory without changing anything.
 *
 * Returns EVERY account carrying the CAPID, ranked, rather than silently picking
 * one. Duplicate accounts are common on these tenants and the plausible-looking
 * name is frequently the dead twin, so the choice belongs to the human.
 *
 * @param {string} capid
 * @returns {Object} { capid, accounts: [...], secondaryDomain }
 */
function apiLookupCapid(capid) {
  requireAuthorized_();
  const wanted = String(capid == null ? '' : capid).trim();
  if (!WEBAPP_CAPID_RE.test(wanted)) {
    throw new Error('"' + wanted + '" is not a CAPID (expected 5-7 digits).');
  }

  const domain = webappNormalizeSecondaryDomain_(WEBAPP_CONFIG.SECONDARY_EMAIL_DOMAIN);
  const onSheet = {};
  apiListEntries_().forEach(function (e) { onSheet[e.primaryEmail] = e; });

  const accounts = webappFindAccountsByCapid_(wanted).map(function (a) {
    const alias = webappDeriveSecondaryAlias_(a.email, domain);
    return {
      email: a.email,
      name: a.name,
      suspended: a.suspended,
      neverSignedIn: a.neverSignedIn,
      lastLogin: webappFormatDate_(a.lastLogin),
      orgUnitPath: a.orgUnitPath,
      derivedAlias: alias,
      hasAlias: alias ? a.aliases.indexOf(alias) !== -1 : false,
      onSheet: !!onSheet[a.email],
      sheetStatus: onSheet[a.email] ? onSheet[a.email].status : ''
    };
  });

  return { capid: wanted, accounts: accounts, secondaryDomain: domain };
}

/**
 * Adds an account to the tab and provisions its alias immediately.
 *
 * `primaryEmail` comes from the browser, so it is re-verified against the CAPID
 * server-side — a client could otherwise post any address it liked alongside a
 * CAPID it is allowed to manage.
 *
 * @returns {Object} { ok, primaryEmail, alias, status, entries }
 */
function apiAddByCapid(capid, primaryEmail) {
  const actor = requireAuthorized_();
  const wanted = String(capid == null ? '' : capid).trim();
  const email = String(primaryEmail || '').trim().toLowerCase();

  if (!WEBAPP_CAPID_RE.test(wanted)) throw new Error('"' + wanted + '" is not a CAPID (expected 5-7 digits).');

  const domain = webappNormalizeSecondaryDomain_(WEBAPP_CONFIG.SECONDARY_EMAIL_DOMAIN);
  if (!domain) throw new Error('No secondary domain is configured for this tenant (TENANT_SECONDARY_EMAIL_DOMAIN).');

  const account = webappGetAccount_(email);
  if (!account) throw new Error('No account exists at ' + email + '.');
  if (account.capids.indexOf(wanted) === -1) {
    throw new Error(email + ' does not carry CAPID ' + wanted + '. Look the CAPID up again.');
  }

  const alias = webappDeriveSecondaryAlias_(account.email, domain);
  if (!alias) throw new Error('Could not derive a secondary address for ' + account.email + '.');

  // Serializes concurrent admins within THIS project. It cannot serialize against
  // the nightly run in the other project — which is why nothing below shifts rows.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20 * 1000)) throw new Error('Another change is in progress. Try again in a moment.');

  try {
    const sheet = webappGetAliasSheet_();
    const existing = webappFindRowByEmail_(sheet, account.email);

    // Already listed: do not add a second row (two rows for one account make the
    // nightly status meaningless). Reconcile the account instead, which is what
    // the person actually wanted if the alias never got created.
    const result = account.aliases.indexOf(alias) !== -1
      ? { status: 'OK — already present: ' + alias, conflict: false }
      : webappAddAliasToAccount_(account.email, alias);

    const row = existing || webappClaimRow_(sheet);
    sheet.getRange(row, WEBAPP_COL_PRIMARY, 1, WEBAPP_COL_COUNT).setValues([[
      account.email, '', result.status, new Date(), wanted, actor
    ]]);

    webappAudit_(actor, existing ? 'RECONCILE' : 'ADD', wanted, account.email, alias, result.status);

    return {
      ok: !result.conflict && result.status.indexOf('ERROR') !== 0,
      alreadyListed: !!existing,
      primaryEmail: account.email,
      alias: alias,
      status: result.status,
      entries: apiListEntries_()
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Removes an account from the tab AND revokes the alias, so the address stops
 * delivering. Removing the row alone would leave a live address behind forever —
 * the nightly module only ever adds.
 *
 * @returns {Object} { ok, primaryEmail, alias, status, entries }
 */
function apiRemoveByEmail(primaryEmail) {
  const actor = requireAuthorized_();
  const email = String(primaryEmail || '').trim().toLowerCase();
  if (!email) throw new Error('No account was given to remove.');

  const domain = webappNormalizeSecondaryDomain_(WEBAPP_CONFIG.SECONDARY_EMAIL_DOMAIN);
  if (!domain) throw new Error('No secondary domain is configured for this tenant (TENANT_SECONDARY_EMAIL_DOMAIN).');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20 * 1000)) throw new Error('Another change is in progress. Try again in a moment.');

  try {
    const sheet = webappGetAliasSheet_();
    const row = webappFindRowByEmail_(sheet, email);
    if (!row) throw new Error(email + ' is not on the Secondary Aliases tab.');

    const values = sheet.getRange(row, WEBAPP_COL_PRIMARY, 1, WEBAPP_COL_COUNT).getValues()[0];
    const capid = String(webappCell_(values, 4)).trim();
    // Honor a column-B override: that, not the derived address, is the alias the
    // nightly module actually created for this row.
    const alias = String(webappCell_(values, 1)).trim().toLowerCase() ||
      webappDeriveSecondaryAlias_(email, domain);

    const account = webappGetAccount_(email);
    let status;
    if (!account) {
      // The account is gone (deleted, or moved to another wing). Nothing to revoke;
      // taking the row off is still exactly right.
      status = 'NOT PRESENT — no account at ' + email;
    } else {
      status = webappRemoveAliasFromAccount_(account, alias, domain).status;
      if (status.indexOf('REFUSED') === 0 || status.indexOf('ERROR') === 0) {
        // Leave the row in place: a half-done removal that looks complete is worse
        // than one that visibly failed and can be retried.
        webappAudit_(actor, 'REMOVE-FAILED', capid, email, alias, status);
        return { ok: false, primaryEmail: email, alias: alias, status: status, entries: apiListEntries_() };
      }
    }

    // Clear rather than deleteRow — see the SHEET CONTRACT note at the top.
    sheet.getRange(row, WEBAPP_COL_PRIMARY, 1, WEBAPP_COL_COUNT).clearContent();
    webappAudit_(actor, 'REMOVE', capid, email, alias, status);

    return { ok: true, primaryEmail: email, alias: alias, status: status, entries: apiListEntries_() };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// SHEET PLUMBING
// ============================================================================

function webappGetAliasSheet_() {
  const id = WEBAPP_CONFIG.AUTOMATION_SPREADSHEET_ID;
  if (!id) throw new Error('TENANT_AUTOMATION_SPREADSHEET_ID is not set on this script project.');

  const ss = SpreadsheetApp.openById(id);
  let sheet = ss.getSheetByName(WEBAPP_ALIAS_SHEET);
  if (!sheet) {
    // The tab is optional in the main project, so a tenant enabling this app may
    // not have one yet. Create it with the exact column contract rather than
    // making the admin build it by hand and risk a column-order mistake.
    sheet = ss.insertSheet(WEBAPP_ALIAS_SHEET);
    sheet.getRange(1, 1, 1, WEBAPP_COL_COUNT).setValues([[
      'Primary Email', 'Alias Email (override)', 'Status', 'Last Run', 'CAPID', 'Added By'
    ]]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    Logger.info('Created the Secondary Aliases tab', { spreadsheetId: id });
  }
  return sheet;
}

/** 1-based sheet row whose column A is this address, or 0. */
function webappFindRowByEmail_(sheet, email) {
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(webappCell_(rows[i], 0)).trim().toLowerCase() === email) return i + 1;
  }
  return 0;
}

/**
 * A row to write into: the first blank one left behind by a removal, else a new
 * row at the bottom. Reusing blanks keeps them from accumulating without ever
 * shifting a row that the nightly run may be holding a position for.
 */
function webappClaimRow_(sheet) {
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (!String(webappCell_(rows[i], 0)).trim()) return i + 1;
  }
  return Math.max(sheet.getLastRow(), 1) + 1;
}

/**
 * Append-only record of who changed what. The Status column tells you the state
 * of a row; only this tells you that someone removed an address last March.
 * Failures are logged too — a refused removal is exactly the kind of thing
 * someone will later swear never happened.
 */
function webappAudit_(actor, action, capid, primaryEmail, alias, result) {
  try {
    const ss = SpreadsheetApp.openById(WEBAPP_CONFIG.AUTOMATION_SPREADSHEET_ID);
    let sheet = ss.getSheetByName(WEBAPP_AUDIT_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(WEBAPP_AUDIT_SHEET);
      sheet.getRange(1, 1, 1, 6).setValues([[
        'Timestamp', 'Actor', 'Action', 'CAPID', 'Primary Email', 'Result'
      ]]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(), actor, action, capid, primaryEmail, result]);
  } catch (err) {
    // Never fail the operation because the log failed — the change already
    // happened, and throwing here would report a success as an error.
    Logger.error('Could not write the alias admin audit row', {
      actor: actor, action: action, user: primaryEmail, alias: alias, errorMessage: err.message
    });
  }
}

/** getDataRange() spans only populated columns, so trailing cells read undefined. */
function webappCell_(row, index) {
  const v = row[index];
  return (v === undefined || v === null) ? '' : v;
}

function webappFormatDate_(value) {
  if (!value) return '';
  const d = (value instanceof Date) ? value : new Date(value);
  if (isNaN(d.getTime()) || d.getTime() <= 0) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}
