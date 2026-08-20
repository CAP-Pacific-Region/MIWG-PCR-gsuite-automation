/***********************************************
 * File: Notify.gs
 * Description: When this web app creates a secondary-domain alias, make the address
 * immediately usable and tell the member — the same treatment the nightly
 * addSecondaryDomainAliases() run gives a hand-added row, done here at the moment
 * of the click so nobody waits for the next morning.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-08-19
 * Changes: 1.0.0 — initial version.
 ***********************************************/

/**
 * DUPLICATED LOGIC — keep in step with the nightly module.
 *
 * These are ports of the member-notification code in
 * src/accounts-and-groups/SecondaryDomainAliases.gs (1.4.0) and the impersonation
 * helper getImpersonatedToken_ in src/accounts-and-groups/UpdateMembers.gs. They
 * are copied, not shared, because this is a separate script project (see Config.gs
 * for why). The nightly run only emails on a genuine INSERT — but an alias this app
 * creates is 'already present' by the time the trigger runs, so the trigger would
 * never notify a member added here. That is precisely why the same logic has to
 * live in this project too. test/SecondaryAliasWebAppNotify.test.js pins it.
 */

/** Template basename (this project has no folders, unlike src/). */
const WEBAPP_WELCOME_TEMPLATE = 'SecondaryAliasWelcomeEmail';

/**
 * Turns on Gmail "Send mail as" for the member so the new address is ready to
 * SEND from, not just receive. Mirrors configureSecondaryAliasSendAs_ in the
 * nightly module: reuses service-account impersonation, auto-accepted because the
 * address is on a domain this tenant owns. BEST-EFFORT, never throws.
 *
 * @returns {'configured'|'exists'|'skipped-admin'|'unavailable'|'failed'}
 */
function webappConfigureSendAs_(primaryEmail, aliasEmail, fullName, isAdmin) {
  // Google refuses to impersonate an admin for user settings (the sendAs call
  // 403s). Rare on this list; hand them the manual steps instead.
  if (isAdmin) {
    Logger.info('Auto Send-As skipped for an admin account; manual steps will be sent', {
      user: primaryEmail, alias: aliasEmail
    });
    return 'skipped-admin';
  }

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SA_IMPERSONATION_EMAIL') || !props.getProperty('SA_PRIVATE_KEY')) {
    Logger.warn('Auto Send-As unavailable — no impersonation credentials; sending manual steps', {
      user: primaryEmail
    });
    return 'unavailable';
  }

  const scope = 'https://www.googleapis.com/auth/gmail.settings.basic ' +
                'https://www.googleapis.com/auth/gmail.settings.sharing';
  try {
    const token = webappGetImpersonatedToken_(primaryEmail, scope);
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
function webappSendAsReady_(outcome) {
  return outcome === 'configured' || outcome === 'exists';
}

/**
 * Emails the member that their address exists, how (or that we already set up)
 * sending from it, and the CAPR 120-1 §6.9 official-use reminder. BEST-EFFORT,
 * never throws — the alias is already created.
 *
 * @returns {boolean} true iff a message was sent.
 */
function webappMaybeSendWelcome_(primaryEmail, aliasEmail, fullName, sendAsOutcome) {
  const sender = String(WEBAPP_CONFIG.AUTOMATION_SENDER_EMAIL || '').trim();
  if (!sender) {
    Logger.warn('Welcome email not sent: no TENANT_AUTOMATION_SENDER_EMAIL configured', {
      user: primaryEmail, alias: aliasEmail
    });
    return false;
  }

  try {
    const html = webappRenderWelcome_(primaryEmail, aliasEmail, fullName, sendAsOutcome);
    const subject = 'Your new ' + WEBAPP_CONFIG.ORG_LABEL + ' email address: ' + aliasEmail;
    const options = { htmlBody: html, from: sender, name: WEBAPP_CONFIG.SENDER_NAME };
    const replyTo = String(WEBAPP_CONFIG.ITSUPPORT_EMAIL || '').trim();
    if (replyTo) options.replyTo = replyTo;

    GmailApp.sendEmail(primaryEmail, subject, webappHtmlToPlainText_(html), options);
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
 * The "sending from your new address" block, which differs by whether we already
 * turned on Send-As. The alias/primary are inlined (not left as {{placeholders}}):
 * this fragment is substituted INTO the template and a global replace does not
 * re-scan what it just inserted.
 */
function webappSendAsSectionHtml_(primaryEmail, aliasEmail, sendAsOutcome) {
  const addr = '<strong>' + aliasEmail + '</strong>';
  const primary = '<strong>' + primaryEmail + '</strong>';

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

  if (webappSendAsReady_(sendAsOutcome)) {
    return '<p><strong>Sending mail from your new address — already set up.</strong> ' +
      'We have turned on sending from ' + addr + ' for you, so there is nothing to install.</p>' +
      switching + timing;
  }

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
    switching + timing;
}

/** Renders SecondaryAliasWelcomeEmail.html. Replacer FUNCTION so a '$&' cannot corrupt. */
function webappRenderWelcome_(primaryEmail, aliasEmail, fullName, sendAsOutcome) {
  const itSupportUrl = String(
    PropertiesService.getScriptProperties().getProperty('TENANT_ITSUPPORT_URL') || ''
  ).trim() || 'https://support.pcrcap.org';

  const fields = {
    fullName: String(fullName || '').trim() || 'Member',
    primaryEmail: primaryEmail,
    aliasEmail: aliasEmail,
    wingName: WEBAPP_CONFIG.WING_NAME,
    orgLabel: WEBAPP_CONFIG.ORG_LABEL,
    itSupportUrl: itSupportUrl,
    sendAsSection: webappSendAsSectionHtml_(primaryEmail, aliasEmail, sendAsOutcome),
    signature: '<strong>' + WEBAPP_CONFIG.ORG_LABEL + ' IT Team</strong><br>' +
      WEBAPP_CONFIG.WING_NAME + ' Civil Air Patrol'
  };

  return HtmlService
    .createHtmlOutputFromFile(WEBAPP_WELCOME_TEMPLATE)
    .getContent()
    .replace(/{{(\w+)}}/g, function (match, key) {
      return Object.prototype.hasOwnProperty.call(fields, key) ? fields[key] : match;
    });
}

/** Minimal HTML->text for the plain-text alternative part. */
function webappHtmlToPlainText_(html) {
  // NOT a security sanitizer — it renders our OWN template to the text/plain
  // alternative part. It is still written to leave no markup or half-tags: whole
  // non-text blocks go first, then every complete tag, then any stray angle
  // bracket, so an unclosed "<script"/"<style" cannot survive. Entities are then
  // decoded in a SINGLE pass, so decoding &amp; -> & cannot re-form an entity for
  // a later pass to decode again.
  var s = String(html)
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<head[\s\S]*?<\/head>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<(?:br|\/p|\/li|\/h1|\/div)>/gi, '\n')
    .replace(/<[^>]*>/g, '')
    .replace(/[<>]/g, '');
  var ENT = { amp: '&', bull: '•', rarr: '->', ldquo: '"', rdquo: '"', nbsp: ' ' };
  s = s.replace(/&(amp|bull|rarr|ldquo|rdquo|nbsp);/g, function (m, name) { return ENT[name]; });
  return s.replace(/\n{3,}/g, '\n\n').replace(/[ \t]+\n/g, '\n').trim();
}

/**
 * Mints an OAuth access token that impersonates `userToImpersonate`, via the
 * tenant's domain-wide-delegated service account. Ported verbatim from
 * getImpersonatedToken_ in src/accounts-and-groups/UpdateMembers.gs — keep in step.
 */
function webappGetImpersonatedToken_(userToImpersonate, scope) {
  const props = PropertiesService.getScriptProperties();
  const SERVICE_ACCOUNT_EMAIL = props.getProperty('SA_IMPERSONATION_EMAIL');
  let PRIVATE_KEY = props.getProperty('SA_PRIVATE_KEY');

  if (!SERVICE_ACCOUNT_EMAIL || !PRIVATE_KEY) {
    throw new Error('Missing service account credentials in Script Properties (SA_IMPERSONATION_EMAIL / SA_PRIVATE_KEY).');
  }
  PRIVATE_KEY = PRIVATE_KEY.replace(/\\n/g, '\n');

  const now = Math.floor(Date.now() / 1000);
  const claimSet = {
    iss: SERVICE_ACCOUNT_EMAIL,
    sub: userToImpersonate,
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now,
    scope: scope
  };
  const header = { alg: 'RS256', typ: 'JWT' };
  const toSign =
    Utilities.base64EncodeWebSafe(JSON.stringify(header)) + '.' +
    Utilities.base64EncodeWebSafe(JSON.stringify(claimSet));
  const signature = Utilities.computeRsaSha256Signature(toSign, PRIVATE_KEY);
  const jwt = toSign + '.' + Utilities.base64EncodeWebSafe(signature);

  const response = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: jwt },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Token exchange failed (' + code + '): ' + body);
  }
  return JSON.parse(body).access_token;
}
