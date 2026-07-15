/***********************************************
 * File: SecondaryDomainAliases.gs
 * Description: Adds a directory alias on a verified secondary domain to the accounts
 * listed in the "Secondary Aliases" tab, keeping the local part of each account's
 * primary address (e.g. jane.doe@cawgcap.org -> jane.doe@cawg.cap.gov). The tab is a
 * curated opt-in list, not the whole roster — only listed accounts are touched.
 * Safe to run unattended on a daily trigger.
 * Author: Noel Luneau
 * Version: 1.1.0
 * Date: 2026-07-14
 * Changes: Made the run trigger-safe — batch the sheet write-back into one call,
 *   latch CONFLICT rows so they report once instead of every night, and take a
 *   script lock so a scheduled run and a manual run cannot overlap.
 *   (1.0.0: initial version.)
 ***********************************************/

/**
 * Tab in the automation spreadsheet listing the accounts to alias.
 * Columns are positional, matching the house style of the `Aliases` tab:
 *   A  Primary Email  (required) — the existing account
 *   B  Alias Email    (optional) — override; blank derives localpart + secondary domain
 *   C  Status         (written back by this script)
 *   D  Last Run       (written back by this script)
 */
const SECONDARY_ALIAS_SHEET = 'Secondary Aliases';

/** Columns C and D are script-owned; A and B are yours. */
const SECONDARY_ALIAS_STATUS_COL = 3;

/**
 * Status prefix for an address that is already taken elsewhere in the tenant.
 * Rows carrying it are skipped quietly on later runs — see shouldSkipLatched_().
 */
const SECONDARY_ALIAS_CONFLICT_PREFIX = 'CONFLICT';

/**
 * Preview the run without touching any account. Writes a Status of
 * "DRY RUN — would add ..." for every row so the derived addresses can be
 * eyeballed against the roster before anything is created.
 *
 * Run this by hand. Never attach it to a trigger — it writes to the same Status
 * column the real run reads.
 */
function previewSecondaryDomainAliases() {
  processSecondaryDomainAliases_(true);
}

/**
 * Adds the secondary-domain alias to every account on the sheet.
 *
 * This is the trigger entry point. Idempotent and safe to run daily: an account
 * that already carries its alias is skipped, so a normal run over a settled list
 * makes no write calls at all.
 *
 * Schedule it AFTER updateAllMembers() (see ADMIN_GUIDE §8) so that an account
 * added to the tab the same morning already exists by the time this runs.
 */
function addSecondaryDomainAliases() {
  // Only one run at a time. Without this, the daily trigger firing while someone
  // is running it by hand would double-process rows and interleave sheet writes.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    Logger.warn('addSecondaryDomainAliases skipped: another run holds the lock.');
    return;
  }

  try {
    processSecondaryDomainAliases_(false);
  } finally {
    lock.releaseLock();
  }
}

function processSecondaryDomainAliases_(dryRun) {
  // Every tenant runs this same code. A blank secondary domain is the correct,
  // expected state for a tenant that has no parallel domain (cadets, pacific) —
  // so this is a normal no-op, NOT an error. Logging it at ERROR would put a
  // permanent false alarm in that tenant's logs every night.
  const domain = (CONFIG.SECONDARY_EMAIL_DOMAIN || '').trim();
  if (!domain) {
    Logger.info('Secondary-alias module not enabled for this tenant; skipping.', {
      hint: 'To enable, set TENANT_SECONDARY_EMAIL_DOMAIN in Script Properties, e.g. @cawg.cap.gov'
    });
    return;
  }

  // A directory alias can only be created on a domain this tenant has verified.
  // Checking once up front turns what would otherwise be one opaque HTTP 400 per
  // row into a single actionable message.
  if (!isDomainVerified_(domain)) {
    Logger.error('Secondary domain is not a verified domain in this tenant.', {
      domain: domain,
      hint: 'Admin console > Domains > Manage domains — add and verify it first. ' +
            'A subdomain of cap.gov also needs a DNS TXT record, which CAP National must publish.'
    });
    return;
  }

  // Also a normal no-op: the tab is optional and only the tenants that use this
  // feature need to create it. getSheetByName() returns null (it does not throw)
  // when the tab is absent.
  const sheet = SpreadsheetApp
    .openById(CONFIG.AUTOMATION_SPREADSHEET_ID)
    .getSheetByName(SECONDARY_ALIAS_SHEET);

  if (!sheet) {
    Logger.info('No secondary-alias tab in this tenant\'s spreadsheet; nothing to do.', {
      sheet: SECONDARY_ALIAS_SHEET
    });
    return;
  }

  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) {
    Logger.info('Secondary-alias tab has no data rows; nothing to do.');
    return;
  }

  const timestamp = new Date();
  const dataRowCount = rows.length - 1;

  // Accumulate column C/D values and write them back in ONE setValues() call at
  // the end. Two setValue() calls per row is a round-trip each — tolerable when
  // run by hand, wasteful every night. Seeded with each row's CURRENT values so
  // rows we skip keep what they already had rather than being blanked.
  const writeBack = [];
  for (let i = 1; i < rows.length; i++) {
    writeBack.push([secondaryAliasCell_(rows[i], 2), secondaryAliasCell_(rows[i], 3)]);
  }

  let added = 0;
  let alreadyPresent = 0;
  let latched = 0;
  let failed = 0;

  // Row 0 is the header; sheet rows are 1-based, so sheet row = i + 1.
  for (let i = 1; i < rows.length; i++) {
    const w = i - 1;  // index into writeBack
    const primaryEmail = String(secondaryAliasCell_(rows[i], 0)).trim().toLowerCase();
    if (!primaryEmail) continue;

    const override = String(secondaryAliasCell_(rows[i], 1)).trim().toLowerCase();
    const priorStatus = String(secondaryAliasCell_(rows[i], 2)).trim();
    const aliasEmail = override || deriveSecondaryAlias_(primaryEmail, domain);

    if (!aliasEmail) {
      writeBack[w] = ['ERROR — could not derive an alias address', timestamp];
      failed++;
      continue;
    }

    // A conflict is structural: the address belongs to another account or group,
    // and only a human can resolve it. Re-reporting it every night would train
    // people to ignore this log, so report once and stay quiet — until the row
    // changes (an edited column B yields a different alias, which un-latches it).
    if (shouldSkipLatched_(priorStatus, aliasEmail)) {
      latched++;
      continue;  // leaves Status AND Last Run untouched, preserving the first-seen date
    }

    let user;
    try {
      user = AdminDirectory.Users.get(primaryEmail, { fields: 'primaryEmail,aliases' });
    } catch (err) {
      // Left un-latched deliberately: unlike a conflict this may be transient (an
      // API blip) or pending (the row was added before the account exists), so a
      // later run should retry it.
      Logger.error('Account lookup failed', { user: primaryEmail, error: err.message });
      writeBack[w] = ['ERROR — account not found: ' + err.message, timestamp];
      failed++;
      continue;
    }

    const existing = (user.aliases || []).map(a => String(a).toLowerCase());
    if (existing.indexOf(aliasEmail) !== -1) {
      alreadyPresent++;
      writeBack[w] = ['OK — already present: ' + aliasEmail, timestamp];
      continue;
    }

    if (dryRun) {
      Logger.info('DRY RUN — would add alias', { user: primaryEmail, alias: aliasEmail });
      writeBack[w] = ['DRY RUN — would add ' + aliasEmail, timestamp];
      continue;
    }

    try {
      AdminDirectory.Users.Aliases.insert({ alias: aliasEmail }, primaryEmail);
      Logger.info('Secondary-domain alias added', { user: primaryEmail, alias: aliasEmail });
      writeBack[w] = ['ADDED — ' + aliasEmail, timestamp];
      added++;
    } catch (err) {
      // 409 means the address is taken somewhere else in the tenant — another
      // account, or a group. Unlike addAlias() in UpdateMembers.gs we do NOT
      // fall back to a numbered variant: the whole point is an address that
      // mirrors the primary, and jane.doe.1@ would not.
      const code = err.details?.code;
      const conflict = code === 409;
      const message = conflict
        ? SECONDARY_ALIAS_CONFLICT_PREFIX + ' — ' + aliasEmail +
          ' is already in use by another account or group'
        : 'ERROR — ' + err.message;

      Logger.error('Failed to add secondary-domain alias', {
        user: primaryEmail,
        attemptedAlias: aliasEmail,
        errorMessage: err.message,
        errorCode: code,
        willRetry: !conflict
      });
      writeBack[w] = [message, timestamp];
      failed++;
    }
  }

  sheet.getRange(2, SECONDARY_ALIAS_STATUS_COL, dataRowCount, 2).setValues(writeBack);

  Logger.info(dryRun ? 'Completed secondary-alias DRY RUN' : 'Completed secondary-alias run', {
    domain: domain,
    added: added,
    alreadyPresent: alreadyPresent,
    skippedLatchedConflicts: latched,
    failed: failed
  });
}

/**
 * True if this row already reported a conflict for this same alias address, and
 * so should be passed over silently.
 *
 * Comparing the recorded address against the one we resolved this run is what
 * makes an edit un-latch the row: change column B (or the account's primary) and
 * the addresses no longer match, so it is retried.
 */
function shouldSkipLatched_(priorStatus, aliasEmail) {
  if (priorStatus.indexOf(SECONDARY_ALIAS_CONFLICT_PREFIX) !== 0) return false;
  return priorStatus.toLowerCase().indexOf(aliasEmail) !== -1;
}

/** jane.doe@cawgcap.org + @cawg.cap.gov -> jane.doe@cawg.cap.gov */
function deriveSecondaryAlias_(primaryEmail, secondaryDomain) {
  const at = primaryEmail.indexOf('@');
  if (at <= 0) return '';
  return primaryEmail.slice(0, at) + secondaryDomain;
}

/**
 * True if the domain is present and verified on this tenant. Both primary and
 * secondary domains are returned by Domains.list; an unverified one is listed
 * but cannot carry aliases.
 */
function isDomainVerified_(secondaryDomain) {
  const bare = secondaryDomain.replace(/^@/, '').toLowerCase();
  try {
    const result = AdminDirectory.Domains.list('my_customer');
    return (result.domains || []).some(d =>
      String(d.domainName).toLowerCase() === bare && d.verified === true
    );
  } catch (err) {
    Logger.error('Could not list tenant domains', { error: err.message });
    return false;
  }
}

/**
 * Reads one cell defensively. getDataRange() only spans columns that contain
 * data, so a tab with nothing yet in C/D comes back two columns wide and any
 * index past the end is undefined rather than ''.
 */
function secondaryAliasCell_(row, index) {
  const v = row[index];
  return (v === undefined || v === null) ? '' : v;
}
