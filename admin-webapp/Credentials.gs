/***********************************************
 * File: Credentials.gs
 * Description: Temporary passwords, the welcome email, and the audit ledger that
 * records a send. All three are ports of src/ — see the notes below on what that
 * costs and how the copies are held together.
 *
 * NOT named WelcomeEmail.gs, which is what it wants to be called: Apps Script
 * addresses files WITHOUT their extension, so WelcomeEmail.gs and the
 * WelcomeEmail.html template beside it would be one name and the push is
 * rejected outright ("A file with this name already exists"). The HTML keeps the
 * name, because matching src/recruiting-and-retention/WelcomeEmail.html is what
 * makes the byte-for-byte copy obvious.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-08-16
 * Changes: 1.0.0 — initial version.
 ***********************************************/

/**
 * THREE COPIES, AND WHY EACH IS ACCEPTABLE
 *
 * 1. WelcomeEmail.html is a byte-for-byte copy of
 *    src/recruiting-and-retention/WelcomeEmail.html. A member must not be able to
 *    tell whether their credentials came from provisioning or from a help desk,
 *    and two templates that drift would make the wing's welcome mail depend on
 *    which route created the account. test/AdminWebApp.test.js compares the two
 *    files and fails on any byte of difference — so updating the wing's welcome
 *    email means updating both, which the test will insist on.
 *
 * 2. admGenerateTempPassword_() is a port of generateTempPassword_(). Same
 *    complexity guarantees, same ambiguous-character exclusions. A weaker
 *    generator here would be a real hole, since these passwords are read aloud
 *    over the phone as often as they are mailed.
 *
 * 3. The LEDGER is not a copy but SHARED STATE: the same WelcomeEmailLedger.txt
 *    in the same CAPWATCH folder that src/accounts-and-groups/WelcomeEmailAudit.gs
 *    reads and writes. That is the point — a welcome email sent from this app has
 *    to stop the monthly audit reporting that member as MISSED. Which means the
 *    FORMAT here must match exactly, including the version number, and the test
 *    pins that too.
 */

const ADM_WELCOME_LEDGER_FILE_NAME = 'WelcomeEmailLedger.txt';

/**
 * Bumped in lockstep with WELCOME_AUDIT_CONFIG.LEDGER_VERSION in src/. A
 * mismatch makes the main project's loader refuse to run rather than guess, so
 * these two numbers moving apart is a real outage, not a cosmetic drift.
 */
const ADM_WELCOME_LEDGER_VERSION = 1;

/**
 * A temporary password. `changePasswordAtNextLogin` is always set alongside it,
 * so the value is only valid until first sign-in.
 *
 * Guarantees at least one lowercase letter, one uppercase letter, one digit and
 * one special character, so it satisfies any complexity policy regardless of the
 * letters drawn. Ported from generateTempPassword_() in UpdateMembers.gs.
 *
 * @returns {string} a random 14-character password
 */
function admGenerateTempPassword_() {
  const hex = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  const letters = 'abcdefghijkmnpqrstuvwxyz'; // no ambiguous l/o
  let body = '';
  for (let i = 0; i < 20; i += 2) {
    const byte = parseInt(hex.substr(i, 2), 16);
    const ch = letters.charAt(byte % letters.length);
    body += (i % 4 === 0) ? ch.toUpperCase() : ch;
  }

  const specials = '!@#$%^&*';
  const lower = letters.charAt(parseInt(hex.substr(20, 2), 16) % letters.length);
  const upper = letters.charAt(parseInt(hex.substr(22, 2), 16) % letters.length).toUpperCase();
  const digit = String(parseInt(hex.substr(24, 2), 16) % 10);
  const special = specials.charAt(parseInt(hex.substr(26, 2), 16) % specials.length);
  return body + lower + upper + digit + special;
}

/**
 * Sets a new password on an account.
 *
 * Separated from everything that decides WHETHER to: this function has no
 * opinion, so the guards live in exactly one place (Actions.gs) and cannot be
 * skipped by a second caller that forgot them.
 *
 * @param {string} email
 * @param {string} password
 * @returns {void}
 * @throws {Error} when the directory refuses the update
 */
function admSetPassword_(email, password) {
  AdminDirectory.Users.update({
    password: password,
    changePasswordAtNextLogin: true
  }, email);
}

/**
 * Sends the welcome email carrying a temporary password.
 *
 * @param {Object} member - a record from admBuildMemberRecord_()
 * @param {string} email - the account the credentials are for
 * @param {string} tempPassword
 * @param {Array<string>} recipients - vetted by admWelcomeEligibility_(); NEVER
 *   the account being reset, and never an address on a tenant domain
 * @returns {void}
 */
function admSendWelcomeEmail_(member, email, tempPassword, recipients) {
  const html = HtmlService.createTemplateFromFile('WelcomeEmail').getRawContent();

  const merged = html
    .replace(/{{WING}}/g, ADMIN_CONFIG.WING)
    .replace(/{{firstName}}/g, member.firstName)
    .replace(/{{lastName}}/g, member.lastName)
    .replace(/{{email}}/g, email)
    .replace(/{{password}}/g, tempPassword)
    .replace(/{{ITSUPPORT_EMAIL}}/g, ADMIN_CONFIG.SUPPORT_EMAIL)
    .replace(/{{DOMAIN}}/g, ADMIN_CONFIG.DOMAIN)
    .replace(/{{rank}}/g, member.rank || '')
    .replace(/{{primaryEmail}}/g, member.email || '')
    .replace(/{{secondaryEmail}}/g, member.secondaryEmail || '');

  const options = {
    // The vetted list, NOT [member.email, member.secondaryEmail] as src/ uses.
    // src/ sends at account creation, when no address of ours exists to mail into
    // by mistake; here the account may be years old and its own address may well
    // be the member's CAPWATCH primary. Mailing credentials to the mailbox they
    // unlock is the silent failure this whole module is shaped around.
    to: recipients.join(','),
    subject: 'New Workspace Account – ' +
      (member.rank ? member.rank + ' ' : '') + member.firstName + ' ' + member.lastName,
    htmlBody: merged
  };
  if (ADMIN_CONFIG.SUPPORT_EMAIL) options.cc = ADMIN_CONFIG.SUPPORT_EMAIL;

  MailApp.sendEmail(options);
}

// ============================================================================
// LEDGER — shared state with src/accounts-and-groups/WelcomeEmailAudit.gs
// ============================================================================

/**
 * Loads the ledger. Ported from welcomeLedgerLoad_().
 *
 * A missing file is normal before the first seed and reads as empty. A file that
 * exists but cannot be parsed THROWS rather than being re-baselined — silently
 * continuing would either bury real findings or invent them, and this app must
 * not be the thing that quietly resets a wing-wide audit.
 *
 * @returns {{seededAt: string, sent: Object}}
 */
function admWelcomeLedgerLoad_() {
  const folder = DriveApp.getFolderById(ADMIN_CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files = folder.getFilesByName(ADM_WELCOME_LEDGER_FILE_NAME);
  if (!files.hasNext()) return { seededAt: '', sent: {} };

  const content = files.next().getBlob().getDataAsString();
  if (!content) return { seededAt: '', sent: {} };

  let parsed;
  try {
    parsed = JSON.parse(content);
  } catch (e) {
    throw new Error('Cannot parse ' + ADM_WELCOME_LEDGER_FILE_NAME + ' — the welcome-email ' +
      'audit ledger is corrupt. Fix it in Drive before sending from here.');
  }
  if (!parsed || parsed.version !== ADM_WELCOME_LEDGER_VERSION) {
    throw new Error(ADM_WELCOME_LEDGER_FILE_NAME + ' has version ' +
      (parsed ? parsed.version : 'none') + ', expected ' + ADM_WELCOME_LEDGER_VERSION +
      '. This app and the main project must agree on the ledger format.');
  }
  return { seededAt: parsed.seededAt || '', sent: parsed.sent || {} };
}

/**
 * Records that a welcome email was sent to this CAPID, so the monthly audit in
 * src/ stops reporting the member as MISSED.
 *
 * THE CALLER MUST NOT LET A FAILURE HERE ESCAPE: the email has already gone out,
 * and a bookkeeping problem must never turn a successful send into a failed one.
 * The cost of a lost entry is one false MISSED row in the next audit.
 *
 * @param {string|number} capid
 * @returns {void}
 */
function admWelcomeLedgerRecordSent_(capid) {
  const ledger = admWelcomeLedgerLoad_();
  ledger.sent[String(capid).trim()] = {
    on: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    by: 'send'
  };

  const folder = DriveApp.getFolderById(ADMIN_CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const files = folder.getFilesByName(ADM_WELCOME_LEDGER_FILE_NAME);
  const content = JSON.stringify({
    version: ADM_WELCOME_LEDGER_VERSION,
    seededAt: ledger.seededAt || '',
    written: new Date().toISOString(),
    sent: ledger.sent
  });

  if (files.hasNext()) files.next().setContent(content);
  else folder.createFile(ADM_WELCOME_LEDGER_FILE_NAME, content);
}

/**
 * What the ledger says about one CAPID, for the member card. Never throws — an
 * unreadable ledger is worth a line on the card, not a failed lookup.
 *
 * @param {string} capid
 * @returns {{state: string, on: string, by: string, note: string}}
 */
function admWelcomeLedgerStatus_(capid) {
  try {
    const ledger = admWelcomeLedgerLoad_();
    const entry = ledger.sent[String(capid).trim()];
    if (entry) {
      return {
        state: entry.by === 'seed' ? 'assumed-welcomed' : 'welcomed',
        on: entry.on || '',
        by: entry.by || '',
        note: entry.by === 'seed'
          ? 'Assumed from login history when the ledger was seeded, not from a recorded send.'
          : ''
      };
    }
    return {
      state: ledger.seededAt ? 'no-record' : 'unseeded',
      on: '', by: '',
      note: ledger.seededAt
        ? 'No welcome email has been recorded for this CAPID.'
        : 'The ledger has not been seeded yet, so nothing can be concluded from it.'
    };
  } catch (err) {
    Logger.warn('Welcome-email ledger could not be read', { capsn: String(capid), errorMessage: err.message });
    return { state: 'unavailable', on: '', by: '', note: err.message };
  }
}
