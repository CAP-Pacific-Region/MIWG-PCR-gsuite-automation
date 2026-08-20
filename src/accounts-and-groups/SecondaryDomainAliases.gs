/***********************************************
 * File: SecondaryDomainAliases.gs
 * Description: Adds a directory alias on a verified secondary domain to the accounts
 * listed in the "Secondary Aliases" tab, keeping the local part of each account's
 * primary address (e.g. jane.doe@cawgcap.org -> jane.doe@cawg.cap.gov). The tab is a
 * curated opt-in list, not the whole roster — only listed accounts are touched.
 * Safe to run unattended on a daily trigger.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.4.0
 * Date: 2026-08-19
 * Changes: 1.4.0 — when an alias is newly ADDED, make the address usable and tell
 *   the member. (1) configureSecondaryAliasSendAs_ turns on Gmail "Send mail as"
 *   for the member via the existing service-account impersonation, so the address
 *   is ready to SEND from, not just receive — the member has no setup steps to
 *   follow and fail. Admin accounts (which cannot be impersonated for settings) and
 *   any failure fall back to manual steps. (2) The member is emailed: whether we
 *   set sending up for them or they must, plus the CAPR 120-1 §6.9 reminder that a
 *   .gov address is for official CAP business only — no commercial or fundraising
 *   use. Both fire ONLY on a real insert, so a settled list does nothing (an
 *   already-present alias never reaches this path). Both are best-effort and never
 *   throw — the alias is already created and must not read as failed. Opt-out via
 *   Script Property SECONDARY_ALIAS_NOTIFY=false. New template
 *   SecondaryAliasWelcomeEmail.html. Uses gmail.send/script.send_mail and
 *   gmail.settings.sharing, all already in the manifest.
 *   1.3.0 — added findSecondaryDomainAliasBlockers(), a READ-ONLY diagnostic
 *   for the silent failure this module's latching creates. The secondary domain is meant
 *   to carry ALIASES on primary-domain accounts, but a separate ACCOUNT sitting on
 *   first.last@<secondary> occupies that exact address, so Users.Aliases.insert 409s and
 *   the row latches — reported once, silent thereafter (by design, so the log stays
 *   readable). The affected member then has no secondary alias indefinitely, with one
 *   stale CONFLICT row as the only trace. The diagnostic pairs each secondary-domain
 *   account with the primary-domain account it is denying, so the blockers can be cleared
 *   deliberately. Found on seniors: four such accounts, all original-population artifacts.
 *   1.2.1 — Accept TENANT_SECONDARY_EMAIL_DOMAIN with or without a leading '@'. It
 *   was set bare on the seniors tenant, and the derive step concatenates, so every
 *   address came out as 'jane.doecawg.cap.gov' — no '@' at all. Caught by the first
 *   live preview.
 *   (1.2.0: let previewSecondaryDomainAliases() run before the secondary domain is
 *   verified (warn instead of bailing), so the list can be built and validated in
 *   advance; the real run still refuses.
 *   1.1.0: made the run trigger-safe — batch the sheet write-back into one call,
 *   latch CONFLICT rows so they report once instead of every night, and take a
 *   script lock so a scheduled run and a manual run cannot overlap.
 *   1.0.0: initial version.)
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
 * Deliberately runs even when the secondary domain is not verified yet (it only
 * warns), so the list can be built and checked in advance of the domain going
 * live. It still confirms every listed account exists and resolves the address
 * each would get.
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
  const configured = (CONFIG.SECONDARY_EMAIL_DOMAIN || '').trim();
  if (!configured) {
    Logger.info('Secondary-alias module not enabled for this tenant; skipping.', {
      hint: 'To enable, set TENANT_SECONDARY_EMAIL_DOMAIN in Script Properties, e.g. @cawg.cap.gov'
    });
    return;
  }

  // The property is typed by hand per tenant, so accept it with or without the
  // leading '@' rather than trusting the format. Getting this wrong is silent and
  // total: deriveSecondaryAlias_() concatenates, so a bare 'cawg.cap.gov' yields
  // 'jane.doecawg.cap.gov' for every row — no '@' at all.
  const domain = normalizeSecondaryDomain_(configured);
  if (!domain) {
    Logger.error('TENANT_SECONDARY_EMAIL_DOMAIN is not a usable domain.', {
      configured: configured,
      hint: 'Expected something like "@cawg.cap.gov" or "cawg.cap.gov".'
    });
    return;
  }

  // A directory alias can only be created on a domain this tenant has verified;
  // Aliases.insert rejects anything else. Checking once up front turns what would
  // otherwise be one opaque HTTP 400 per row into a single actionable message.
  //
  // This blocks the real run only. A preview writes nothing to any account, so it
  // stays useful while the domain is still pending verification — which is exactly
  // when you want to build and check the list.
  if (!isDomainVerified_(domain)) {
    const hint = 'Admin console > Domains > Manage domains — add and verify it first. ' +
      'A subdomain of cap.gov also needs a DNS TXT record, which CAP National must publish.';

    if (!dryRun) {
      Logger.error('Secondary domain is not a verified domain in this tenant.', {
        domain: domain,
        hint: hint
      });
      return;
    }

    Logger.warn('Secondary domain is not verified yet — previewing anyway. ' +
      'Addresses below are what WOULD be created; no alias can actually be added until ' +
      'the domain is verified.', { domain: domain, hint: hint });
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
  let notified = 0;

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
      user = AdminDirectory.Users.get(primaryEmail, { fields: 'primaryEmail,aliases,name,isAdmin' });
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

      // Set up "Send mail as" for the member so they have nothing manual to do —
      // the address is ready to send from, not just receive. Best-effort: on an
      // admin account (which cannot be impersonated for settings) or any failure
      // this returns a non-'configured' outcome and the email falls back to the
      // manual Gmail steps. Never throws; the alias is already created.
      const sendAs = configureSecondaryAliasSendAs_(
        primaryEmail, aliasEmail, user.name ? user.name.fullName : '', !!user.isAdmin);

      // Tell the member their new address exists, how (or that we already set up)
      // sending from it, and that a .gov address is official-use-only
      // (CAPR 120-1 §6.9). Sent ONLY on a real insert, so a settled list makes no
      // mail. Best-effort: a mail failure must not fail the run or lose the alias.
      notified += maybeSendSecondaryAliasWelcome_(
        primaryEmail, aliasEmail, user.name ? user.name.fullName : '', sendAs) ? 1 : 0;
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
    failed: failed,
    welcomeEmailsSent: notified
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

/**
 * Canonicalizes the configured secondary domain to a single leading-'@' form, so
 * both "@cawg.cap.gov" and "cawg.cap.gov" work. Returns '' for anything that
 * cannot be a domain, which the caller reports rather than building addresses from.
 */
function normalizeSecondaryDomain_(configured) {
  const bare = String(configured || '').trim().toLowerCase().replace(/^@+/, '');
  // A stray interior '@' ("user@cawg.cap.gov") or a dotless value ("cawg") is a
  // typo, not a domain — refuse instead of minting 34 bad addresses from it.
  if (!bare || bare.indexOf('@') !== -1 || bare.indexOf('.') === -1) return '';
  return '@' + bare;
}

/**
 * jane.doe@cawgcap.org + @cawg.cap.gov -> jane.doe@cawg.cap.gov
 * secondaryDomain must already be normalized (leading '@').
 */
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

/**
 * Template name (folder + basename, no extension). A file at
 * src/accounts-and-groups/SecondaryAliasWelcomeEmail.html deploys under Apps
 * Script as this exact name, and HtmlService needs it verbatim — see ADMIN_GUIDE.
 */
const SECONDARY_ALIAS_WELCOME_TEMPLATE = 'accounts-and-groups/SecondaryAliasWelcomeEmail';

/**
 * True unless an admin has switched member notification off. The address is minted
 * either way; this only governs whether the member is told. Opt-OUT (default on),
 * via Script Property SECONDARY_ALIAS_NOTIFY, so a deliberate bulk backfill can be
 * run quietly — although in practice a settled list sends nothing anyway, since
 * only a genuine insert reaches the send path.
 */
function secondaryAliasNotifyEnabled_() {
  const raw = String(
    PropertiesService.getScriptProperties().getProperty('SECONDARY_ALIAS_NOTIFY') || ''
  ).trim().toLowerCase();
  return !(raw === 'false' || raw === '0' || raw === 'no' || raw === 'off');
}

/**
 * Turns on Gmail "Send mail as" for the member so the new address is ready to
 * SEND from, not just receive — the whole point being that the member should not
 * follow setup steps only to find the address unusable.
 *
 * Reuses the service-account impersonation the send-as name sync already relies on
 * (getImpersonatedToken_ + the sendAs REST endpoint in UpdateMembers.gs). Because
 * the address is on a domain this tenant owns, Gmail auto-accepts it: no
 * verification email, no SMTP.
 *
 * BEST-EFFORT, and NEVER throws — the alias is already created before this runs.
 * The outcome only steers what the member is told: a 'configured'/'exists' result
 * lets the email say "already set up," anything else falls back to manual steps.
 *
 * @returns {'configured'|'exists'|'skipped-admin'|'unavailable'|'failed'}
 */
function configureSecondaryAliasSendAs_(primaryEmail, aliasEmail, fullName, isAdmin) {
  // Google refuses to impersonate an admin for user settings (the sendAs call
  // 403s), and addAliasesFromSheet() skips admins for the same reason. Rare on
  // this tab; hand them the manual steps instead of logging a scary failure.
  if (isAdmin) {
    Logger.info('Auto Send-As skipped for an admin account; manual steps will be sent', {
      user: primaryEmail, alias: aliasEmail
    });
    return 'skipped-admin';
  }

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SA_IMPERSONATION_EMAIL') || !props.getProperty('SA_PRIVATE_KEY')) {
    // Not an error: a tenant may not have set up impersonation. The member simply
    // gets the manual steps.
    Logger.warn('Auto Send-As unavailable — no impersonation credentials; sending manual steps', {
      user: primaryEmail
    });
    return 'unavailable';
  }

  const scope = 'https://www.googleapis.com/auth/gmail.settings.basic ' +
                'https://www.googleapis.com/auth/gmail.settings.sharing';
  try {
    const token = getImpersonatedToken_(primaryEmail, scope);
    const resp = UrlFetchApp.fetch(
      'https://gmail.googleapis.com/gmail/v1/users/me/settings/sendAs', {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({
          sendAsEmail: aliasEmail,
          displayName: String(fullName || '').trim(),
          treatAsAlias: true
        }),
        muteHttpExceptions: true
      });

    const code = resp.getResponseCode();
    if (code >= 200 && code < 300) {
      Logger.info('Send-As configured for secondary alias', { user: primaryEmail, alias: aliasEmail });
      return 'configured';
    }
    if (code === 409) {
      // Already a send-as identity on the account — nothing to do, still "ready".
      Logger.info('Send-As already present for secondary alias', { user: primaryEmail, alias: aliasEmail });
      return 'exists';
    }
    Logger.warn('Send-As setup failed for secondary alias (member can add it manually)', {
      user: primaryEmail, alias: aliasEmail, code: code, response: resp.getContentText()
    });
    return 'failed';
  } catch (err) {
    Logger.warn('Send-As setup errored for secondary alias (member can add it manually)', {
      user: primaryEmail, alias: aliasEmail, errorMessage: err.message
    });
    return 'failed';
  }
}

/** True once the send-as identity exists, so the email can skip the manual steps. */
function secondaryAliasSendAsReady_(outcome) {
  return outcome === 'configured' || outcome === 'exists';
}

/**
 * Emails the member that their secondary-domain address now exists: how to send
 * from it in Gmail, and the CAPR 120-1 §6.9 reminder that a .gov address is for
 * official CAP business only (no commercial or fundraising use).
 *
 * BEST-EFFORT BY CONTRACT. The caller has already created the alias; a mail
 * failure here must neither throw nor be reported as an alias failure, or a flaky
 * mail send would make a successful provisioning look broken and get "fixed". So
 * every failure is swallowed to a warning and returns false.
 *
 * @returns {boolean} true iff a message was actually sent.
 */
function maybeSendSecondaryAliasWelcome_(primaryEmail, aliasEmail, fullName, sendAsOutcome) {
  if (!secondaryAliasNotifyEnabled_()) return false;

  // Without a from-identity the mail would go out as whatever account holds the
  // trigger, which is not what a member should see. Treat a missing sender as
  // "notifications not configured" rather than guessing.
  const sender = String(CONFIG.AUTOMATION_SENDER_EMAIL || '').trim();
  if (!sender) {
    Logger.warn('Secondary-alias welcome not sent: no TENANT_AUTOMATION_SENDER_EMAIL configured', {
      user: primaryEmail, alias: aliasEmail
    });
    return false;
  }

  try {
    const html = renderSecondaryAliasWelcome_(primaryEmail, aliasEmail, fullName, sendAsOutcome);
    const subject = 'Your new ' + CONFIG.ORG_LABEL + ' email address: ' + aliasEmail;
    const options = {
      htmlBody: html,
      from: sender,
      name: SENDER_NAME
    };
    const replyTo = String(CONFIG.ITSUPPORT_EMAIL || '').trim();
    if (replyTo) options.replyTo = replyTo;

    // Same retrying send the retention mail uses, so a transient Gmail quota blip
    // does not drop the notice.
    executeWithRetry(function () {
      GmailApp.sendEmail(primaryEmail, subject, htmlToPlainText_(html), options);
    });

    Logger.info('Secondary-alias welcome sent', { user: primaryEmail, alias: aliasEmail });
    return true;
  } catch (err) {
    Logger.warn('Secondary-alias welcome failed to send (alias was still created)', {
      user: primaryEmail, alias: aliasEmail, errorMessage: err.message
    });
    return false;
  }
}

/**
 * Builds the "sending from your new address" block of the email, which differs by
 * whether we already turned on Send-As for the member.
 *
 * The alias address is inlined here rather than left as a {{placeholder}}: this
 * fragment is substituted INTO the template as a replacement value, and a global
 * replace does not re-scan what it just inserted, so any {{...}} left inside would
 * ship to the member raw.
 */
function secondaryAliasSendAsSectionHtml_(primaryEmail, aliasEmail, sendAsOutcome) {
  const addr = '<strong>' + aliasEmail + '</strong>';
  const primary = '<strong>' + primaryEmail + '</strong>';

  // How to pick which address a single message goes out from. Shown in both
  // variants — it is the same action whether we set the address up or the member
  // did, and it is the step people most often miss.
  const switching = '<p><strong>Choosing which address a message is sent from.</strong> ' +
    'You decide per message, right in the compose window. Click <strong>Compose</strong>, ' +
    'then click the <strong>From</strong> line near the top of the new message — a menu drops ' +
    'down listing both ' + primary + ' and ' + addr + '. Pick whichever you want this message ' +
    'to come from. Your primary address, ' + primary + ', stays the default; switching to ' +
    addr + ' affects only the message you are writing.</p>' +
    '<p class="muted">On phones, tap the <strong>From</strong> field in the Gmail app the same ' +
    'way to switch addresses before sending.</p>';

  const timing = '<p class="muted">A brand-new address can take up to a day for Google to ' +
    'finish registering. If it is not selectable yet, check back later that day — nothing ' +
    'is wrong and you do not need to do anything.</p>';

  if (secondaryAliasSendAsReady_(sendAsOutcome)) {
    return '<p><strong>Sending mail from your new address — already set up.</strong> ' +
      'We have turned on sending from ' + addr + ' for you, so there is nothing to install.</p>' +
      switching +
      timing;
  }

  // Fallback: the member sets it up themselves (admin accounts, or if the
  // automatic setup did not succeed).
  return '<p><strong>Sending mail from your new address.</strong> Incoming mail already ' +
    'arrives without any setup. To also <em>send</em> as ' + addr + ', add it to Gmail once:</p>' +
    '<ol class="steps">' +
      '<li>In Gmail, click the <strong>gear icon</strong> (top right) &rarr; ' +
        '<strong>See all settings</strong>.</li>' +
      '<li>Open the <strong>Accounts</strong> tab.</li>' +
      '<li>Under <strong>&ldquo;Send mail as&rdquo;</strong>, click ' +
        '<strong>Add another email address</strong>.</li>' +
      '<li>Enter your name and ' + addr + '. Leave <strong>&ldquo;Treat as an alias&rdquo;</strong> ' +
        'checked, then click <strong>Next Step</strong> and <strong>Add</strong>.</li>' +
      '<li>Because this address is on a domain your account already owns, there is no ' +
        'confirmation email to wait for and no separate password or SMTP server to configure. ' +
        '(If Gmail says it cannot verify the address, it has not finished registering yet — ' +
        'try again later that day.)</li>' +
    '</ol>' +
    switching +
    timing;
}

/**
 * Renders SecondaryAliasWelcomeEmail.html with this member's and tenant's values.
 * Uses a replacer FUNCTION, not a string, so a value containing '$&' cannot
 * corrupt the output; an unrecognized {{placeholder}} is left visible so a typo
 * surfaces in the mail rather than silently blanking.
 */
function renderSecondaryAliasWelcome_(primaryEmail, aliasEmail, fullName, sendAsOutcome) {
  const itSupportUrl = String(
    PropertiesService.getScriptProperties().getProperty('TENANT_ITSUPPORT_URL') || ''
  ).trim() || 'https://support.pcrcap.org';

  const fields = {
    fullName: String(fullName || '').trim() || 'Member',
    primaryEmail: primaryEmail,
    aliasEmail: aliasEmail,
    wingName: CONFIG.WING_NAME,
    orgLabel: CONFIG.ORG_LABEL,
    itSupportUrl: itSupportUrl,
    // The send-from instructions differ depending on whether we already set up
    // "Send mail as" for them — see secondaryAliasSendAsSectionHtml_.
    sendAsSection: secondaryAliasSendAsSectionHtml_(primaryEmail, aliasEmail, sendAsOutcome),
    signature: '<strong>' + CONFIG.ORG_LABEL + ' IT Team</strong><br>' +
      CONFIG.WING_NAME + ' Civil Air Patrol'
  };

  return HtmlService
    .createHtmlOutputFromFile(SECONDARY_ALIAS_WELCOME_TEMPLATE)
    .getContent()
    .replace(/{{(\w+)}}/g, function (match, key) {
      return Object.prototype.hasOwnProperty.call(fields, key) ? fields[key] : match;
    });
}

/**
 * Minimal HTML-to-text for the plain-text alternative part. Not a general
 * converter — just enough that a client with images/HTML disabled still gets a
 * readable body instead of raw markup. The HTML part carries the real formatting.
 */
function htmlToPlainText_(html) {
  // NOT a security sanitizer, and deliberately NOT a regex HTML filter (those are
  // both unreliable and flagged by scanners). It renders our OWN template to the
  // text/plain part. Strategy: take only the <body> — found by plain string search,
  // not a filtering regex — so the <head>/<style> CSS is excluded outright rather
  // than stripped. Then drop tags and any stray angle bracket, so no partial tag
  // survives, and decode the few entities in ONE pass so &amp; -> & cannot re-form
  // an entity for a later pass.
  var s = bodyInnerHtml_(String(html));
  s = s.replace(/<(?:br|\/p|\/li|\/h1|\/div)>/gi, '\n')
       .replace(/<[^>]*>/g, '')
       .replace(/[<>]/g, '');
  var ENT = { amp: '&', bull: '•', rarr: '->', ldquo: '"', rdquo: '"', nbsp: ' ' };
  s = s.replace(/&(amp|bull|rarr|ldquo|rdquo|nbsp);/g, function (m, name) { return ENT[name]; });
  return s.replace(/\n{3,}/g, '\n\n').replace(/[ \t]+\n/g, '\n').trim();
}

/**
 * The content between <body ...> and </body>, located with plain indexOf (no
 * regex), so the head/style block is excluded from the plaintext. Falls back to
 * the whole string if there is no body element.
 */
function bodyInnerHtml_(s) {
  var lower = s.toLowerCase();
  var open = lower.indexOf('<body');
  if (open === -1) return s;
  var contentStart = s.indexOf('>', open);
  if (contentStart === -1) return s;
  var close = lower.indexOf('</body>', contentStart);
  return s.slice(contentStart + 1, close === -1 ? s.length : close);
}

/**
 * READ-ONLY DIAGNOSTIC — finds secondary-domain addresses that exist as their own
 * ACCOUNT instead of being available as an alias.
 *
 * WHY THIS EXISTS: the secondary domain is meant to carry ALIASES on members'
 * primary-domain accounts. A separate account sitting on `first.last@<secondary>`
 * occupies that exact address, so Users.Aliases.insert 409s. processSecondaryDomainAliases_
 * then LATCHES the conflict (reports once, silent thereafter, by design so the log
 * stays readable) — which means the affected members have quietly had no
 * secondary-domain alias since whenever that account appeared, with one stale
 * CONFLICT row as the only evidence.
 *
 * These blockers are typically original-population artifacts: never signed into,
 * often suspended, and carrying no CAPID (so scanAccountsWithoutCapid lists them too).
 *
 * Writes NOTHING. Uses admin.directory.user.readonly, already in appsscript.json.
 *
 * @returns {Object} { secondaryDomain, blockers: [...], strays: [...], scannedUsers }
 */
function findSecondaryDomainAliasBlockers() {
  const start = new Date();
  const domain = normalizeSecondaryDomain_(CONFIG.SECONDARY_EMAIL_DOMAIN || '');
  if (!domain) {
    Logger.info('No secondary domain configured for this tenant; nothing to check.', {
      hint: 'Set TENANT_SECONDARY_EMAIL_DOMAIN in Script Properties, e.g. @cawg.cap.gov'
    });
    return { secondaryDomain: '', blockers: [], strays: [], scannedUsers: 0 };
  }
  Logger.info('Secondary-domain alias blocker scan starting (read-only)', { domain: domain });

  // Every account in the tenant, indexed by primaryEmail, plus the set of addresses
  // already held as aliases (an alias also occupies an address).
  const byEmail = {};
  const aliasOwner = {};
  let scannedUsers = 0;
  let pageToken = null;

  do {
    const page = AdminDirectory.Users.list({
      customer: 'my_customer',
      maxResults: 500,
      projection: 'full',
      pageToken: pageToken
    });
    (page.users || []).forEach(function (u) {
      scannedUsers++;
      const email = String(u.primaryEmail || '').toLowerCase();
      byEmail[email] = u;
      (u.aliases || []).forEach(function (a) {
        aliasOwner[String(a).toLowerCase()] = email;
      });
    });
    pageToken = page.nextPageToken;
  } while (pageToken);

  const suffix = domain.toLowerCase();          // '@cawg.cap.gov'
  const blockers = [];                          // stray blocks a real primary account
  const strays = [];                            // on the secondary domain, no primary counterpart

  Object.keys(byEmail).forEach(function (email) {
    if (email.slice(-suffix.length) !== suffix) return;   // not on the secondary domain
    const u = byEmail[email];
    const local = email.slice(0, email.length - suffix.length);

    // Which primary-domain account WOULD want this address as its alias?
    const primaryEmail = local + String(CONFIG.EMAIL_DOMAIN || '').toLowerCase();
    const primary = byEmail[primaryEmail];

    const info = {
      strayAccount: u.primaryEmail,
      created: u.creationTime || null,
      suspended: !!u.suspended,
      neverSignedIn: !u.lastLoginTime || new Date(u.lastLoginTime).getTime() <= 0,
      lastLogin: u.lastLoginTime || null,
      orgUnitPath: u.orgUnitPath || '',
      blocksAliasFor: primary ? primary.primaryEmail : null,
      primaryAlreadyHasAlias: !!aliasOwner[email]
    };
    if (primary) blockers.push(info); else strays.push(info);
  });

  blockers.sort(function (a, b) { return String(a.strayAccount).localeCompare(String(b.strayAccount)); });
  strays.sort(function (a, b) { return String(a.strayAccount).localeCompare(String(b.strayAccount)); });

  Logger.info('===== SECONDARY-DOMAIN ALIAS BLOCKERS =====' +
    '\nSecondary domain ................ ' + domain +
    '\nScanned users ................... ' + scannedUsers +
    '\nBLOCKERS (alias is blocked) ..... ' + blockers.length +
    '   <-- each denies a real account its alias' +
    '\nStrays (no primary counterpart) . ' + strays.length);

  function render_(title, list, note) {
    if (!list.length) return;
    const lines = ['===== ' + title + ' (' + list.length + ') ====='];
    if (note) lines.push(note, '');
    list.forEach(function (b) {
      lines.push('    ' + b.strayAccount +
        '\n        created=' + (b.created || '?') +
        '  ' + (b.suspended ? 'SUSPENDED' : 'active') +
        '  ' + (b.neverSignedIn ? 'never-signed-in' : 'last-login=' + b.lastLogin) +
        '  ou=' + b.orgUnitPath +
        (b.blocksAliasFor ? '\n        BLOCKS the secondary alias for: ' + b.blocksAliasFor : '') +
        (b.primaryAlreadyHasAlias ? '\n        (address is ALSO held as an alias — inconsistent, investigate)' : ''));
    });
    Logger.info(lines.join('\n'));
  }

  render_('BLOCKERS — a real primary account is being denied its alias', blockers,
    'Each stray below occupies the exact address the alias module wants to add to the\n' +
    'primary account named under it. Users.Aliases.insert 409s and the row LATCHES, so it\n' +
    'is reported once and then stays silent. If the stray is never-signed-in it holds no\n' +
    'mail: retire or delete it, then RE-RUN the alias module. NOTE: shouldSkipLatched_ only\n' +
    'un-latches when the row\'s alias address CHANGES, so clearing the blocker alone will\n' +
    'not retry it — nudge column B on that row to force a retry.');
  render_('STRAYS — on the secondary domain with no primary-domain counterpart', strays,
    'These block nothing today. Likely original-population artifacts; confirm before removing.');

  const result = {
    secondaryDomain: domain,
    scannedUsers: scannedUsers,
    blockers: blockers,
    strays: strays,
    durationMs: new Date() - start
  };
  Logger.info('Secondary-domain alias blocker scan complete', {
    blockers: blockers.length,
    strays: strays.length,
    durationMs: result.durationMs
  });
  return result;
}
