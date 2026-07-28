/***********************************************
 * File: GmailSignature.gs
 * Description: The only writes in this project. Sets the caller's Gmail signature
 * on their own org-owned Send-As identities, via service-account impersonation.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-07-28
 * Changes: 1.0.0 — initial version.
 ***********************************************/

/**
 * WHY IMPERSONATION AND NOT THE CALLER'S OWN GRANT
 *
 * Gmail settings have no admin-on-behalf-of endpoint: the only way to write a
 * user's signature is to hold a token for that user. Deploying this app as
 * USER_ACCESSING would give it one — but the app also has to read the CAPWATCH
 * extract, which members cannot see and must not be given, so the project runs as
 * the deployer and mints a per-user token the same way
 * updateSignatureForAllAliases() in src/ does.
 *
 * The cost is that SA_IMPERSONATION_EMAIL / SA_PRIVATE_KEY must be present in
 * THIS project's Script Properties as well as the main project's: Script
 * Properties are per project and there is no shared store. That is a second copy
 * of a credential that can act as any user in the tenant, so:
 *   - it lives ONLY in Script Properties, never in this repo (see the leaked-key
 *     history — a key in git is a key to rotate, not to edit out), and
 *   - the subject is always the AUTHENTICATED caller, established in Auth.gs and
 *     passed down from SignatureApi.gs. No function here accepts an account from
 *     the browser, and none should ever be given one.
 */

/** Scopes the impersonated token is minted for. Nothing here reads mail. */
const SIG_GMAIL_SCOPES =
  'https://www.googleapis.com/auth/gmail.settings.basic ' +
  'https://www.googleapis.com/auth/gmail.settings.sharing';

/**
 * An OAuth2 access token for one user, via domain-wide delegation.
 * Ported from getImpersonatedToken_() in src/accounts-and-groups/UpdateMembers.gs.
 *
 * @param {string} userToImpersonate
 * @param {string} scope - space-separated
 * @returns {string} access token
 */
function sigImpersonatedToken_(userToImpersonate, scope) {
  const props = PropertiesService.getScriptProperties();
  const serviceAccount = props.getProperty('SA_IMPERSONATION_EMAIL');
  let privateKey = props.getProperty('SA_PRIVATE_KEY');

  if (!serviceAccount || !privateKey) {
    throw new Error('Missing service account credentials in Script Properties ' +
      '(SA_IMPERSONATION_EMAIL / SA_PRIVATE_KEY).');
  }

  // If the key was stored with literal "\n", convert to real newlines.
  privateKey = privateKey.replace(/\\n/g, '\n');

  const now = Math.floor(Date.now() / 1000);
  const claimSet = {
    iss: serviceAccount,
    sub: userToImpersonate,
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now,
    scope: scope
  };

  const header = { alg: 'RS256', typ: 'JWT' };
  const toSign =
    `${Utilities.base64EncodeWebSafe(JSON.stringify(header))}.` +
    `${Utilities.base64EncodeWebSafe(JSON.stringify(claimSet))}`;
  const jwt = `${toSign}.${Utilities.base64EncodeWebSafe(Utilities.computeRsaSha256Signature(toSign, privateKey))}`;

  const response = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: jwt },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code < 200 || code >= 300) {
    Logger.error('Impersonation token request failed', { user: userToImpersonate, code: code });
    throw new Error('Google would not issue a token for your account (HTTP ' + code + ').');
  }

  const token = JSON.parse(body).access_token;
  if (!token) throw new Error('Google returned no access token for your account.');
  return token;
}

/**
 * The caller's Send-As identities, split into the ones this app may write and the
 * ones it must not.
 *
 * THE ORG-OWNED SPLIT IS NOT OPTIONAL. Members add their personal accounts as
 * Send-As identities, and stamping a CAP signature onto someone's private mail
 * would be a real intrusion — the same reason updateSignatureForAllAliases() in
 * src/ carries this check. isOnATenantDomain_() compares the whole domain, so a
 * lookalike (@cawgcap.org.example.com) is refused rather than matched as a suffix.
 *
 * @param {string} userEmail - the AUTHENTICATED caller; never a client value
 * @returns {{token: string, base: string, orgOwned: Array<string>, personal: Array<string>}}
 */
function sigSendAsSnapshot_(userEmail) {
  const token = sigImpersonatedToken_(userEmail, SIG_GMAIL_SCOPES);
  const base = 'https://gmail.googleapis.com/gmail/v1/users/' +
    encodeURIComponent(userEmail) + '/settings/sendAs';

  const listResponse = UrlFetchApp.fetch(base, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  const listCode = listResponse.getResponseCode();
  if (listCode < 200 || listCode >= 300) {
    Logger.error('Failed to list send-as identities', {
      user: userEmail, code: listCode, response: listResponse.getContentText()
    });
    throw new Error('Your Gmail send-as addresses could not be read (HTTP ' + listCode + ').');
  }

  const snapshot = { token: token, base: base, orgOwned: [], personal: [] };
  (JSON.parse(listResponse.getContentText()).sendAs || []).forEach(function (identity) {
    const address = String(identity.sendAsEmail || '').toLowerCase();
    if (!address) return;
    if (isOnATenantDomain_(address)) snapshot.orgOwned.push(address);
    else snapshot.personal.push(address);
  });
  return snapshot;
}

/**
 * Writes `signatureHtml` to every org-owned Send-As identity on the caller's own
 * mailbox, and reports what happened per address.
 *
 * @param {string} userEmail - the AUTHENTICATED caller; never a client value
 * @param {string} signatureHtml
 * @returns {{updated: Array<string>, skipped: Array<string>, failed: Array<Object>}}
 */
function sigApplyToSendAsIdentities_(userEmail, signatureHtml) {
  const snapshot = sigSendAsSnapshot_(userEmail);
  const token = snapshot.token;
  const base = snapshot.base;
  const result = { updated: [], skipped: snapshot.personal.slice(), failed: [] };

  snapshot.orgOwned.forEach(function (address) {
    const response = UrlFetchApp.fetch(base + '/' + encodeURIComponent(address), {
      method: 'patch',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ signature: signatureHtml }),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    if (code >= 200 && code < 300) {
      result.updated.push(address);
    } else {
      Logger.error('Failed to update signature for a send-as identity', {
        user: userEmail, alias: address, code: code, response: response.getContentText()
      });
      result.failed.push({ address: address, code: code });
    }
  });

  Logger.info('Signature applied from the self-service app', {
    user: userEmail,
    updated: result.updated.length,
    skippedNonOrg: result.skipped.length,
    failed: result.failed.length
  });

  return result;
}
