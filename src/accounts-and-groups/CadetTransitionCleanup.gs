/**
 * Cadet → senior transition: closing the old account.
 *
 * Version: 2.0.0
 * Date: 2026-08-03
 * Changes: 2.0.0 — the close no longer DELETES the cadet account. Two live
 *   attempts proved the old design impossible: a Group cannot take an address
 *   while Google still reserves it, and Google reserves a former primary address
 *   for ~20 days after a delete (recovery tombstone) and for longer than an
 *   execution can wait after a rename — both 409'd "Entity already exists". Any
 *   delete-now design therefore means a bounce window of a day to weeks. The
 *   account is now PARKED instead: kept live, forwarding to the senior address,
 *   so the address never stops existing and nothing ever bounces. It costs a
 *   seat until expiry, which is the deliberate trade. Auto-forward is requested
 *   but is best-effort (cross-domain forwarding needs the member to click a
 *   Google confirmation); the daily catch-up sweep is what guarantees delivery.
 *   expireParkedAccounts() deletes the account once the forwarding window ends.
 *   1.1.0 — added remindPendingTransitionCloses(): a daily email to IT when
 *   accounts pass grace and are due for the manual close (or stuck past it).
 *   1.0.0 — initial release.
 *
 * The end of the lifecycle. Runs on the CADETS tenant
 * (TRANSITION_CONFIG.ROLE === 'source'). See CadetTransition.gs for detection
 * and state, CadetTransitionMigrate.gs for the mail move.
 *
 * This is the only part of the feature that destroys data, and the destruction
 * is permanent: Archived User licenses are not provisioned on this edition, so a
 * deleted mailbox has no archive behind it and no undo. Everything here is
 * therefore built to refuse rather than proceed when anything is unclear.
 *
 * WHAT THE CLOSE DOES, and why it does not delete:
 *
 *   1. Catch up any mail that arrived since the migration.
 *   2. Verify the destination account really exists.
 *   3. Park the cadet account: keep it LIVE, forwarding to the senior address.
 *
 * Deleting was tried and cannot work. A Group can only take the address once
 * Google releases it, and Google reserves a former primary address well beyond
 * one execution — ~20 days after a delete, and >75s (docs say up to 24h) after a
 * rename. Both were measured failing on 2026-08-03. So a delete-now design
 * guarantees a bounce window; parking gives none, at the cost of a seat.
 *
 * The seat is reclaimed by expireParkedAccounts() at the end of the forwarding
 * window, by which point nobody is writing to the old address.
 */

/** Continuation-safe cap: how many accounts to close in one execution. */
const CLEANUP_MAX_PER_RUN_ = 10;

// ============================================================================
// DELETION
// ============================================================================

/**
 * Closes cadet accounts whose transition is complete and whose grace has run out.
 *
 * @param {boolean} [dryRun=true] - report what would happen, change nothing.
 *   Defaults to a dry run because the alternative is irreversible.
 * @returns {{closed: number, skipped: number, failed: number}}
 */
function closeCompletedTransitions(dryRun) {
  if (TRANSITION_CONFIG.ROLE !== 'source') {
    Logger.info('Transition cleanup skipped — not the source tenant');
    return { closed: 0, skipped: 0, failed: 0 };
  }
  // A dry run only reads, so it needs no lock and should work even while another
  // run holds it. A real close deletes accounts — serialize it.
  if (dryRun !== false) return closeCompletedTransitions_(true);
  return withTransitionLock_(() => closeCompletedTransitions_(false),
    { closed: 0, skipped: 0, failed: 0 });
}

function closeCompletedTransitions_(isDry) {
  const rows = readTransitions_();
  const now = new Date();
  let closed = 0;
  let skipped = 0;
  let failed = 0;

  console.log(isDry ? '=== DRY RUN — nothing will be deleted ===' : '=== LIVE — deleting accounts ===');
  console.log('');

  for (const capid in rows) {
    const row = rows[capid];
    const reason = whyNotCloseable_(row, now);

    if (reason) {
      console.log(`${capid} | ${row.Name} | skip — ${reason}`);
      skipped++;
      continue;
    }

    if (closed >= CLEANUP_MAX_PER_RUN_) {
      console.log(`${capid} | ${row.Name} | deferred — per-run cap reached`);
      skipped++;
      continue;
    }

    if (isDry) {
      console.log(`${capid} | ${row.Name} | WOULD close ${row.CadetEmail}, ` +
        `forward to ${row.SeniorEmail} for ${TRANSITION_CONFIG.FORWARD_GROUP_MONTHS} months`);
      closed++;
      continue;
    }

    try {
      closeOneTransition_(row);
      closed++;
    } catch (e) {
      setTransitionField_(row._rowNumber, 'Notes',
        `Close failed ${new Date().toISOString()}: ${e && e.message ? e.message : String(e)}`);
      Logger.error('Transition close failed', {
        capid: capid,
        cadetEmail: row.CadetEmail,
        errorMessage: e && e.message ? e.message : String(e)
      });
      failed++;
    }
  }

  console.log('');
  console.log(isDry
    ? `DRY RUN — ${closed} would be closed, ${skipped} skipped. Pass false to do it.`
    : `Closed ${closed}, skipped ${skipped}, failed ${failed}.`);

  return { closed: closed, skipped: skipped, failed: failed };
}

/**
 * Every reason a row must not be closed. Returns '' when it is safe.
 *
 * Written as a list of refusals rather than a permission check on purpose: the
 * default for an unrecognised state is "do not delete". Anything unclear here
 * costs a suspended seat until a human looks; getting it wrong costs a mailbox
 * that cannot be recovered.
 *
 * @param {Object} row
 * @param {Date} now
 * @returns {string} reason to refuse, or '' to proceed
 */
function whyNotCloseable_(row, now) {
  if (row.MigrationStatus !== TRANSITION_CONFIG.STATUS.COMPLETE) {
    return 'migration is ' + (row.MigrationStatus || 'unset') + ', not COMPLETE';
  }
  if (!row.SeniorEmail) {
    return 'no destination account — nothing to forward to';
  }
  if (!row.MigratedDate) {
    return 'COMPLETE but no MigratedDate — cannot tell what was migrated when';
  }
  if (row.ForwardGroupCreated) {
    return 'already closed on ' + row.ForwardGroupCreated;
  }

  // Mail known to exist in the source and not in the destination — a failed
  // catch-up, or messages too large to fetch. Deleting now destroys precisely
  // the mail we already know did not make it across.
  if (String(row.Notes || '').indexOf('DO NOT DELETE') > -1) {
    return 'mail is known to be unmigrated (see Notes) — handle it, then clear the note';
  }

  if (TRANSITION_CONFIG.REQUIRE_MIGRATION_BEFORE_DELETE && isBlankField_(row.MessagesMigrated)) {
    return 'MessagesMigrated is empty — no evidence anything was actually moved';
  }

  // A row whose close already got as far as deleting the account, but left no
  // forwarding group, must NOT be run through the close again: the account is
  // gone, so the catch-up sweep would fail against a deleted mailbox. It needs
  // repairFailedTransitionCloses() instead (after the account is restored).
  if (needsForwardRepair_(row)) {
    return 'previous close failed after deleting — restore the account, then run ' +
      'repairFailedTransitionCloses(false). Do NOT re-close.';
  }

  // Deleting the account destroys the member's Drive with it, and this function
  // knew nothing about Drive until now: mail was migrated, the row read COMPLETE,
  // and closing would have silently taken gigabytes of files that were never
  // copied. Cross-tenant ownership transfer does not exist and sharing does not
  // survive the owner, so the files must be copied first — see
  // CadetTransitionDrive.gs. Measured 2026-07-15: 43.5GB across 6511 files for
  // four members, one holding 36.7GB.
  //
  // Refuses until DriveMigrated says otherwise. 0 is a legitimate answer for a
  // member with an empty Drive and must be treated as handled; blank means nobody
  // has looked. isBlankField_ is the whole point — a plain `|| ''` or `!value`
  // check coalesces 0 to empty and blocks the close forever, which it did.
  if (isBlankField_(row.DriveMigrated)) {
    return 'Drive not handled — closing would destroy it. Copy it, or set ' +
      'DriveMigrated=0 if there is nothing to copy';
  }

  // Same reasoning as Drive: personal contacts die with the account. Blank means
  // nobody looked; 0 is a deliberate "nothing to copy". EVERY member so far has
  // 0 saved contacts, so the falsy-zero bug here blocked all of them.
  if (isBlankField_(row.ContactsMigrated)) {
    return 'Contacts not handled — closing would destroy them. Copy them, or set ' +
      'ContactsMigrated=0 if there is nothing to copy';
  }

  const deleteAfter = row.DeleteAfter ? new Date(row.DeleteAfter) : null;
  if (!deleteAfter) return 'no DeleteAfter set';
  if (now < deleteAfter) {
    const days = Math.ceil((deleteAfter - now) / 86400000);
    return `${days}d of grace remaining`;
  }

  // Telling someone their account is going, then deleting it, is the whole
  // bargain. A member who was never told has had no chance to warn anyone.
  if (!row.NotifiedDate) {
    return 'member has not been notified — run notifyCompletedTransitions(false) first';
  }

  return '';
}

/**
 * Emails IT when transitioned accounts have passed their grace and need the
 * manual close — the nudge that closes the loop, since deletion is deliberately
 * not automated.
 *
 * Read-only (no lock needed). Two buckets, both keyed off whyNotCloseable_ so
 * they match exactly what a real close would do:
 *   - READY: grace elapsed and every guard passes → run closeCompletedTransitions(false).
 *   - STUCK: grace elapsed but something still blocks it (a DO NOT DELETE hold,
 *     an un-notified member) → needs a human to resolve the block first.
 *
 * Sends only when there is something to report, so it goes quiet on its own once
 * the closes are done. Meant to run daily (armTransitionTriggers installs it).
 *
 * @returns {{ready: number, stuck: number}}
 */
function remindPendingTransitionCloses() {
  if (TRANSITION_CONFIG.ROLE !== 'source') return { ready: 0, stuck: 0 };

  const rows = readTransitions_();
  const now = new Date();
  const ready = [];
  const stuck = [];

  for (const capid in rows) {
    const row = rows[capid];
    if (row.MigrationStatus !== TRANSITION_CONFIG.STATUS.COMPLETE) continue;
    if (row.ForwardGroupCreated) continue;                 // already closed

    const deleteAfter = row.DeleteAfter ? new Date(row.DeleteAfter) : null;
    if (!deleteAfter || now < deleteAfter) continue;       // grace not up yet

    const reason = whyNotCloseable_(row, now);
    if (reason === '') ready.push(row);
    else stuck.push({ row: row, reason: reason });
  }

  if (!ready.length && !stuck.length) {
    Logger.info('No transition accounts past grace — no reminder sent');
    return { ready: 0, stuck: 0 };
  }

  const lines = [];
  lines.push('Transitioned cadet accounts have passed their deletion grace period.');
  lines.push('');

  if (ready.length) {
    lines.push('READY TO DELETE (' + ready.length + ') — run closeCompletedTransitions(false)');
    lines.push('on the cadets Apps Script project, signed in as ' + AUTOMATION_SENDER_EMAIL + ':');
    ready.forEach(function (r) {
      lines.push('  ' + r.CAPID + '  ' + r.Name + '  ' +
        r.CadetEmail + ' -> ' + r.SeniorEmail);
    });
    lines.push('');
  }

  if (stuck.length) {
    lines.push('PAST GRACE BUT BLOCKED (' + stuck.length + ') — needs a human before it can close:');
    stuck.forEach(function (s) {
      lines.push('  ' + s.row.CAPID + '  ' + s.row.Name + '  — ' + s.reason);
    });
    lines.push('');
  }

  lines.push('Deletion is permanent — no archive, no undo. Review before running.');
  lines.push('closeCompletedTransitions(true) shows the full picture first.');

  const subject = '[Action] ' + ready.length + ' cadet account(s) ready to delete' +
    (stuck.length ? ', ' + stuck.length + ' stuck' : '');

  executeWithRetry(function () {
    GmailApp.sendEmail(ITSUPPORT_EMAIL, subject, lines.join('\n'), {
      replyTo: ITSUPPORT_EMAIL,
      from: AUTOMATION_SENDER_EMAIL,
      name: SENDER_NAME
    });
  });

  Logger.info('Pending-close reminder sent', {
    to: ITSUPPORT_EMAIL, ready: ready.length, stuck: stuck.length
  });
  return { ready: ready.length, stuck: stuck.length };
}

/**
 * Closes one account: catch up, verify, delete, forward.
 *
 * @param {Object} row
 */
function closeOneTransition_(row) {
  Logger.info('Closing transition', {
    capid: row.CAPID, cadetEmail: row.CadetEmail, seniorEmail: row.SeniorEmail
  });

  // 1. Final catch-up. The source is live and receiving right up to this moment.
  const sweep = catchUpOneTransition_(row);
  if (!sweep.ok) {
    throw new Error('Final catch-up failed, refusing to delete: ' + sweep.error);
  }
  if (sweep.imported > 0) {
    Logger.info('Final sweep moved late mail', {
      capid: row.CAPID, imported: sweep.imported
    });
  }

  // 2. Prove the destination is real before touching the source. A typo or a
  // stale SeniorEmail would otherwise strand mail with nowhere to go.
  assertDestinationExists_(row.SeniorEmail);

  // 3. PARK the account rather than deleting it.
  //
  // The original design deleted the account and put a Group at the freed
  // address. That cannot work: Google reserves a former primary address after it
  // stops being one — ~20 days for a deleted account (recovery tombstone), and
  // still >75s for a renamed one (measured 2026-08-03, both attempts 409'd).
  // The group can never be created in the same run that frees the address, so
  // any delete-now design means a bounce window of a day to weeks.
  //
  // Instead the account stays LIVE and becomes the forwarding mechanism itself:
  // zero bounce gap, because the address never stops existing. It costs a seat
  // until expiry, which is the deliberate trade (see expireParkedAccounts).
  const expires = new Date();
  expires.setMonth(expires.getMonth() + TRANSITION_CONFIG.FORWARD_GROUP_MONTHS);

  const fwd = parkAccountForForwarding_(row.CadetEmail, row.SeniorEmail);

  setTransitionField_(row._rowNumber, 'ForwardGroupCreated', new Date().toISOString());
  setTransitionField_(row._rowNumber, 'ForwardGroupExpires', expires.toISOString());
  setTransitionField_(row._rowNumber, 'Notes',
    `Parked ${new Date().toISOString()}; account kept live and forwarding to ` +
    `${row.SeniorEmail} until ${expires.toISOString()}. Auto-forward: ${fwd.status}.` +
    (fwd.status === 'pending'
      ? ' MEMBER MUST CLICK the confirmation Google emailed to their senior address;' +
        ' until then the daily sweep carries mail across instead.'
      : ''));

  Logger.info('Transition parked', {
    capid: row.CAPID,
    forwarding: row.CadetEmail + ' -> ' + row.SeniorEmail,
    autoForward: fwd.status,
    expires: expires.toISOString()
  });
}

/**
 * Turns the cadet account into the forwarding mechanism, keeping it live.
 *
 * Requests Gmail auto-forwarding to the senior address. Cross-domain forwarding
 * needs the TARGET to confirm — Google emails the senior address a verification
 * link — so this usually returns 'pending' rather than 'accepted', and the
 * forward does not carry mail until the member clicks it.
 *
 * That is why it is a best-effort layer, not the guarantee: the daily catch-up
 * sweep moves anything that arrives regardless, verified or not. Auto-forward
 * just makes delivery instant instead of sweep-interval-delayed once confirmed.
 *
 * NOTE: needs gmail.settings.basic + gmail.settings.sharing on the cadets local
 * SA's DWD grant. Without them this returns 'unavailable' and the sweep alone
 * does the work — degraded, not broken.
 *
 * @param {string} cadetEmail
 * @param {string} seniorEmail
 * @returns {{status: string}} 'accepted' | 'pending' | 'unavailable'
 */
function parkAccountForForwarding_(cadetEmail, seniorEmail) {
  try {
    const token = getImpersonatedToken_(cadetEmail,
      'https://www.googleapis.com/auth/gmail.settings.sharing ' +
      'https://www.googleapis.com/auth/gmail.settings.basic');

    const created = gmailSettingsFetch_(
      'https://gmail.googleapis.com/gmail/v1/users/me/settings/forwardingAddresses',
      { method: 'post', token: token, payload: JSON.stringify({ forwardingEmail: seniorEmail }) },
      'create forwarding address'
    );

    const status = String((created && created.verificationStatus) || 'pending').toLowerCase();

    if (status === 'accepted') {
      gmailSettingsFetch_(
        'https://gmail.googleapis.com/gmail/v1/users/me/settings/autoForwarding',
        { method: 'put', token: token, payload: JSON.stringify({
            enabled: true, emailAddress: seniorEmail, disposition: 'leaveInInbox'
          }) },
        'enable auto-forwarding'
      );
      Logger.info('Auto-forwarding enabled', { from: cadetEmail, to: seniorEmail });
      return { status: 'accepted' };
    }

    Logger.info('Forwarding address created but awaiting the member\'s confirmation', {
      from: cadetEmail, to: seniorEmail
    });
    return { status: 'pending' };

  } catch (e) {
    // Missing scope, or Gmail refusing the cross-domain forward. Not fatal: the
    // sweep is what actually guarantees delivery.
    Logger.warn('Auto-forward unavailable — the daily sweep will carry mail instead', {
      from: cadetEmail, to: seniorEmail,
      errorMessage: (e && e.message) ? e.message : String(e)
    });
    return { status: 'unavailable' };
  }
}

/**
 * One Gmail settings call. Separate from gmailFetch_ (which lives in the migrate
 * module and is tuned for message bodies) because these are small, rare calls
 * where a clear error matters more than throughput.
 *
 * @param {string} url
 * @param {{method: string, token: string, payload: string=}} opts
 * @param {string} what
 * @returns {Object}
 */
function gmailSettingsFetch_(url, opts, what) {
  const params = {
    method: opts.method,
    headers: { Authorization: 'Bearer ' + opts.token },
    muteHttpExceptions: true
  };
  if (opts.payload) {
    params.contentType = 'application/json';
    params.payload = opts.payload;
  }
  const resp = UrlFetchApp.fetch(url, params);
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error(`Gmail ${what} failed (${code}): ${text}`);
  }
  return text ? JSON.parse(text) : {};
}

/**
 * Confirms the destination account exists and is usable.
 *
 * @param {string} seniorEmail
 */
function assertDestinationExists_(seniorEmail) {
  const cfg = getCrossTenantConfig_();
  const peer = peerCapidToEmail_();
  const known = Object.keys(peer).some(c => peer[c] === String(seniorEmail).toLowerCase());

  if (!known) {
    throw new Error(`Destination ${seniorEmail} is not in the ${cfg.peerDomain} ` +
      `directory — refusing to delete the source`);
  }
}

/**
 * Catch-up for a single row. Same machinery as catchUpTransitionMail(), scoped.
 *
 * @param {Object} row
 * @returns {{ok: boolean, imported: number, error: string}}
 */
function catchUpOneTransition_(row) {
  const since = Math.floor(new Date(row.MigratedDate).getTime() / 1000) - 60;
  const before = Number(row.MessagesMigrated) || 0;

  try {
    const result = migrateOneTransition_(row, new Date(), {
      query: 'after:' + since,
      notify: false
    });
    return { ok: true, imported: Math.max(0, result.imported - before), error: '' };

  } catch (e) {
    return { ok: false, imported: 0, error: e && e.message ? e.message : String(e) };
  }
}

// ============================================================================
// FORWARDING GROUP
// ============================================================================

/**
 * Creates a Group at the freed cadet address that forwards to the senior one.
 *
 * A group costs no license seat, which is the entire reason this is a group and
 * not a retained account — retention would keep consuming one of the 2000.
 *
 * TWO settings both matter, and they govern opposite directions:
 *  - allowExternalMembers — lets the group DELIVER to the senior address. The
 *    forward target is on the SENIOR tenant, i.e. external to this cadets
 *    domain, so without this the member is either rejected on insert or gets no
 *    mail. This is the setting that actually makes the forward work.
 *  - whoCanPostMessage ANYONE_CAN_POST — lets outsiders who still have the old
 *    address SEND to the group. A default group rejects external senders.
 *
 * Order matters: settings are applied BEFORE the member is added, because on a
 * domain that restricts external members the Members.insert fails until
 * allowExternalMembers is set.
 *
 * DEPENDS ON DOMAIN POLICY: the cadets domain must permit external group
 * members (Admin console → Groups → Sharing settings). allowExternalMembers on
 * the group cannot override a domain that forbids them outright. Verified
 * 2026-07-17 with testForwardingGroup() — a real message reached the senior
 * address.
 *
 * The creation entry point is createForwardingGroupWhenFree_ below. There is no
 * plain create-at-this-address helper on purpose: every real caller is taking
 * over an address a user just held, which needs the retry-until-free behavior.
 */

/**
 * Repairs rows whose close deleted the account but failed to create the
 * forwarding group — the 2026-08-03 batch, stranded by the delete-then-create
 * 409 (see freeAddressAndForward_).
 *
 * PREREQUISITE: restore the affected accounts first, in the Admin console
 * (Users → recently deleted → Recover), within Google's ~20-day window. This
 * function needs the account to exist so it can rename it off the address.
 * previewTransitionRepair() lists exactly who needs restoring.
 *
 * Does NOT re-migrate: their mail, Drive and contacts are already across and the
 * rows still record it. It only vacates the address, stands up the forwarding
 * group, and deletes the renamed shell — the tail of a close that never
 * completed. It deliberately skips the catch-up sweep too: mail that arrived
 * after the original delete went nowhere (the account was gone), so there is
 * nothing new in the source to collect.
 *
 * @param {boolean} [dryRun=true]
 * @returns {{repaired: number, missing: number, failed: number}}
 */
function repairFailedTransitionCloses(dryRun) {
  if (TRANSITION_CONFIG.ROLE !== 'source') return { repaired: 0, missing: 0, failed: 0 };
  const isDry = dryRun !== false;
  if (isDry) return previewTransitionRepair();

  return withTransitionLock_(function () {
    const rows = readTransitions_();
    let repaired = 0, missing = 0, failed = 0;

    for (const capid in rows) {
      const row = rows[capid];
      if (!needsForwardRepair_(row)) continue;

      // The account must be restored first — we cannot rename what isn't there.
      if (!userExists_(row.CadetEmail)) {
        console.log(`${capid} | ${row.Name} | NOT RESTORED — recover ${row.CadetEmail} in the Admin console first`);
        missing++;
        continue;
      }

      try {
        const expires = new Date();
        expires.setMonth(expires.getMonth() + TRANSITION_CONFIG.FORWARD_GROUP_MONTHS);

        const fwd = parkAccountForForwarding_(row.CadetEmail, row.SeniorEmail);

        setTransitionField_(row._rowNumber, 'ForwardGroupCreated', new Date().toISOString());
        setTransitionField_(row._rowNumber, 'ForwardGroupExpires', expires.toISOString());
        setTransitionField_(row._rowNumber, 'Notes',
          `Close repaired ${new Date().toISOString()}; account kept live and forwarding to ` +
          `${row.SeniorEmail} until ${expires.toISOString()}. Auto-forward: ${fwd.status}.` +
          (fwd.status === 'pending'
            ? ' MEMBER MUST CLICK the confirmation Google emailed to their senior address;' +
              ' until then the daily sweep carries mail across instead.'
            : ''));

        console.log(`${capid} | ${row.Name} | repaired — ${row.CadetEmail} kept live, forwarding to ${row.SeniorEmail} (auto-forward: ${fwd.status})`);
        repaired++;

      } catch (e) {
        const msg = (e && e.message) ? e.message : String(e);
        setTransitionField_(row._rowNumber, 'Notes',
          `Close repair FAILED ${new Date().toISOString()}: ${msg}`);
        Logger.error('Transition close repair failed', { capid: capid, errorMessage: msg });
        console.log(`${capid} | ${row.Name} | FAILED — ${msg}`);
        failed++;
      }
    }

    console.log('');
    console.log(`Repaired ${repaired}, awaiting restore ${missing}, failed ${failed}.`);
    return { repaired: repaired, missing: missing, failed: failed };
  }, { repaired: 0, missing: 0, failed: 0 });
}

/** Read-only: who needs repair, and whether they have been restored yet. */
function previewTransitionRepair() {
  const rows = readTransitions_();
  let needs = 0, ready = 0;

  console.log('=== Rows whose close deleted the account but left no forwarding group ===');
  for (const capid in rows) {
    const row = rows[capid];
    if (!needsForwardRepair_(row)) continue;
    needs++;
    const exists = userExists_(row.CadetEmail);
    if (exists) ready++;
    console.log(`${capid} | ${row.Name} | ${row.CadetEmail} → ${row.SeniorEmail} | ` +
      (exists ? 'RESTORED, ready to repair' : 'still deleted — recover it in the Admin console'));
  }

  console.log('');
  if (!needs) {
    console.log('Nothing to repair.');
  } else {
    console.log(`${needs} row(s) need repair; ${ready} restored and ready.`);
    console.log('Restore any missing accounts (Users → recently deleted → Recover),');
    console.log('then run repairFailedTransitionCloses(false).');
  }
  return { repaired: 0, missing: needs - ready, failed: 0 };
}

/**
 * A row that migrated and was closed-ish, but has no forwarding group: the
 * signature of the failed batch.
 *
 * @param {Object} row
 * @returns {boolean}
 */
function needsForwardRepair_(row) {
  if (row.MigrationStatus !== TRANSITION_CONFIG.STATUS.COMPLETE) return false;
  if (!row.SeniorEmail) return false;
  if (row.ForwardGroupCreated) return false;   // already has one
  return String(row.Notes || '').indexOf('Close failed') > -1 ||
         String(row.Notes || '').indexOf('Close repair FAILED') > -1;
}

/** True if a user currently exists (a restored or live account). */
function userExists_(email) {
  try {
    AdminDirectory.Users.get(email, { projection: 'basic', fields: 'primaryEmail' });
    return true;
  } catch (e) {
    return false;
  }
}

// ============================================================================
// EXPIRY
// ============================================================================

/**
 * Removes forwarding groups whose 12 months are up.
 *
 * Without this they accumulate forever. They cost no seat, but they do hold
 * their addresses hostage — a future member with the same name cannot be given
 * an address a stale group still owns.
 *
 * @param {boolean} [dryRun=true]
 * @returns {{removed: number}}
 */
function expireParkedAccounts(dryRun) {
  if (TRANSITION_CONFIG.ROLE !== 'source') return { removed: 0 };
  if (dryRun !== false) return expireParkedAccounts_(true);
  return withTransitionLock_(() => expireParkedAccounts_(false), { removed: 0 });
}

function expireParkedAccounts_(isDry) {
  const rows = readTransitions_();
  const now = new Date();
  let removed = 0;

  for (const capid in rows) {
    const row = rows[capid];
    if (!row.ForwardGroupCreated || !row.ForwardGroupExpires) continue;

    const expires = new Date(row.ForwardGroupExpires);
    if (now < expires) continue;

    if (isDry) {
      console.log(`${capid} | ${row.Name} | WOULD delete parked account ${row.CadetEmail}`);
      removed++;
      continue;
    }

    try {
      // Final sweep before the mailbox goes: anything that arrived since the
      // last one would otherwise die with the account.
      const sweep = catchUpOneTransition_(row);
      if (!sweep.ok) {
        Logger.error('Refusing to delete parked account — final sweep failed', {
          capid: capid, errorMessage: sweep.error
        });
        setTransitionField_(row._rowNumber, 'Notes',
          `Expiry blocked ${new Date().toISOString()} — final sweep failed: ${sweep.error}`);
        continue;
      }

      executeWithRetry(() => AdminDirectory.Users.remove(row.CadetEmail));
      setTransitionField_(row._rowNumber, 'ForwardGroupExpires', '');
      setTransitionField_(row._rowNumber, 'Notes',
        `Parked account deleted ${new Date().toISOString()} — forwarding window ended` +
        (sweep.imported ? `; final sweep moved ${sweep.imported} message(s)` : ''));
      Logger.info('Parked account expired and deleted', {
        address: row.CadetEmail, finalSweep: sweep.imported
      });
      removed++;

    } catch (e) {
      Logger.error('Could not delete parked account', {
        address: row.CadetEmail,
        errorMessage: e && e.message ? e.message : String(e)
      });
    }
  }

  console.log(isDry
    ? `DRY RUN — ${removed} parked account(s) would be deleted. Pass false to do it.`
    : `Deleted ${removed} parked account(s).`);

  return { removed: removed };
}

