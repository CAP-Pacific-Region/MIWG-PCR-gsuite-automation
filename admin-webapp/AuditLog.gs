/***********************************************
 * File: AuditLog.gs
 * Description: Records who did what to whom. Every action in Actions.gs writes
 * one row here, whether it succeeded, was refused, or failed.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.0.0
 * Date: 2026-08-16
 * Changes: 1.0.0 — initial version.
 ***********************************************/

/**
 * WHY THIS EXISTS, GIVEN THAT APPS SCRIPT ALREADY LOGS EVERYTHING
 *
 * The execution log records the DEPLOYER for every run — because that is who the
 * app executes as. From Stackdriver's point of view one person did all of it,
 * which makes the platform's own log useless for the one question that matters
 * about a delegated tool: which admin reset that member's password?
 *
 * So the actor is written down explicitly. Two destinations, deliberately:
 *
 *   - the execution log, always. The record of last resort, and the only one
 *     that survives a misconfigured spreadsheet id.
 *   - a sheet, when TENANT_AUTOMATION_SPREADSHEET_ID is set. This is the one
 *     people will actually read: the wing IT staff who need to review help-desk
 *     activity have a Google account, not Cloud Logging access.
 *
 * WHAT IS NEVER WRITTEN: the temporary password. It appears in exactly two
 * places — the admin's browser and, when mailed, the member's inbox — and a log
 * row would outlive both. `passwordShown: true` records that one was issued
 * without recording what it was.
 */

const ADM_AUDIT_HEADER = [
  'Timestamp', 'Admin', 'Authorized via', 'Action', 'Target account',
  'CAPID', 'Outcome', 'Detail'
];

/**
 * Writes one audit row. NEVER THROWS: an action that succeeded must not be
 * reported to the admin as failed because a spreadsheet was unavailable, and an
 * action that was refused must not turn into an error page. A failure to log is
 * itself logged, which is the best that can be done from inside the logger.
 *
 * @param {Object} actor - from requireAdmin_()
 * @param {string} action - stable slug, e.g. 'reset-password'
 * @param {Object} details
 * @param {string} [details.target] - the account acted on
 * @param {string} [details.capid]
 * @param {string} details.outcome - 'ok' | 'refused' | 'failed'
 * @param {string} [details.detail] - a sentence; no secrets
 * @returns {void}
 */
function admAudit_(actor, action, details) {
  const row = {
    at: new Date(),
    admin: (actor && actor.email) || '(unknown)',
    via: (actor && actor.via) || 'none',
    action: action,
    target: details.target || '',
    capid: details.capid || '',
    outcome: details.outcome || 'ok',
    detail: details.detail || ''
  };

  const logLine = 'Admin web app action';
  const payload = {
    actor: row.admin, via: row.via, action: row.action,
    target: row.target, capsn: row.capid, outcome: row.outcome, detail: row.detail
  };
  if (row.outcome === 'ok') Logger.info(logLine, payload);
  else Logger.warn(logLine, payload);

  if (!ADMIN_CONFIG.AUDIT_SPREADSHEET_ID) return;

  try {
    const sheet = admAuditSheet_();
    if (sheet) {
      sheet.appendRow([
        row.at, row.admin, row.via, row.action, row.target, row.capid, row.outcome, row.detail
      ]);
    }
  } catch (err) {
    Logger.error('Audit row could not be written to the sheet; it is in this log only', {
      spreadsheetId: ADMIN_CONFIG.AUDIT_SPREADSHEET_ID,
      sheetName: ADMIN_CONFIG.AUDIT_SHEET_NAME,
      errorMessage: err.message
    });
  }
}

/**
 * The audit tab, created with its header the first time an action is taken.
 *
 * Creating it lazily rather than at setup is what keeps the deployment checklist
 * short — there is no "remember to add a tab" step to forget, and a tenant that
 * points at a fresh spreadsheet works immediately.
 *
 * @returns {Object|null} a Sheet, or null when the spreadsheet cannot be opened
 */
function admAuditSheet_() {
  const ss = SpreadsheetApp.openById(ADMIN_CONFIG.AUDIT_SPREADSHEET_ID);
  const name = ADMIN_CONFIG.AUDIT_SHEET_NAME;
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  // Two admins acting at the same moment must not create two tabs with the same
  // name (or race one another's header write). Whoever loses the lock re-reads.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    Logger.warn('Could not take the lock to create the audit tab', { errorMessage: err.message });
    return ss.getSheetByName(name);
  }

  try {
    sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(ADM_AUDIT_HEADER);
      sheet.setFrozenRows(1);
      Logger.info('Created the admin web app audit tab', { sheetName: name });
    }
  } finally {
    lock.releaseLock();
  }
  return sheet;
}
