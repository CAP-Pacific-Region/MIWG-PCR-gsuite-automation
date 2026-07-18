/**
 * Cadet → senior transition: closing the old account.
 *
 * Version: 1.2.0
 * Date: 2026-07-17
 * Changes: 1.2.0 — the forwarding group now sets allowExternalMembers=true
 *   (the destination is on the peer/senior tenant, external to this domain, so
 *   without it the group can't deliver the forward) and applies settings BEFORE
 *   adding the member. Added testForwardingGroup() to prove the mechanism —
 *   incl. the domain's external-member policy — with a throwaway group before a
 *   real close, which deletes first and forwards second, depends on it.
 *   1.1.0 — added remindPendingTransitionCloses(): a daily email to IT when
 *   accounts pass grace and are due for the manual close (or stuck past it).
 *   1.0.0 — initial release. Catches up late mail, deletes the cadet account,
 *   then forwards its address to the senior mailbox. The only step that destroys
 *   data — kept manual, never triggered.
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
 * ORDER MATTERS, and not for tidiness:
 *
 *   1. Catch up any mail that arrived since the migration. The source account is
 *      live and still receiving — it is not suspended, because suspension is not
 *      running on this tenant — so a snapshot taken days ago is already stale.
 *      Skipping this silently destroys everything that landed in between.
 *   2. Delete the cadet account.
 *   3. Create a Group at the freed address, forwarding to the senior account.
 *
 * Steps 2 and 3 cannot be reordered: a Google Group cannot share an address with
 * a User, so the address must be freed before the group can take it. That
 * ordering leaves a window where mail to the old address bounces — small, but
 * real, and the reason the completion email tells members to update their
 * contacts rather than trusting the forward.
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

  // 2. Prove the destination is real before destroying the source. A typo or a
  // stale SeniorEmail would otherwise delete the only copy and forward the
  // address into the void.
  assertDestinationExists_(row.SeniorEmail);

  // 3. Delete. Irreversible from here.
  executeWithRetry(() => AdminDirectory.Users.remove(row.CadetEmail));
  Logger.info('Cadet account deleted', {
    capid: row.CAPID, email: row.CadetEmail
  });

  // 4. Take the freed address as a forwarding group. Only possible now: a Group
  // cannot coexist with a User on the same address.
  const expires = new Date();
  expires.setMonth(expires.getMonth() + TRANSITION_CONFIG.FORWARD_GROUP_MONTHS);

  createForwardingGroup_(row.CadetEmail, row.SeniorEmail, row.Name);

  setTransitionField_(row._rowNumber, 'ForwardGroupCreated', new Date().toISOString());
  setTransitionField_(row._rowNumber, 'ForwardGroupExpires', expires.toISOString());
  setTransitionField_(row._rowNumber, 'Notes',
    `Closed ${new Date().toISOString()}; forwards until ${expires.toISOString()}`);

  Logger.info('Transition closed', {
    capid: row.CAPID,
    forwarding: row.CadetEmail + ' -> ' + row.SeniorEmail,
    expires: expires.toISOString()
  });
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
 * allowExternalMembers is set — and by this point the cadet account is already
 * deleted, so a throw here leaves a half-built forward.
 *
 * DEPENDS ON DOMAIN POLICY: the cadets domain must permit external group
 * members (Admin console → Groups → Sharing settings). allowExternalMembers on
 * the group cannot override a domain that forbids them outright. Verify with a
 * throwaway group before relying on this — see testForwardingGroup().
 *
 * @param {string} oldAddress - the freed cadet address
 * @param {string} forwardTo - the senior address
 * @param {string} name - member name, for the group description
 */
function createForwardingGroup_(oldAddress, forwardTo, name) {
  const group = executeWithRetry(() =>
    AdminDirectory.Groups.insert({
      email: oldAddress,
      name: name + ' (forwarding)',
      description: 'Auto-created by the cadet transition automation. Forwards to ' +
        forwardTo + '. Safe to delete once members have updated their contacts.'
    })
  );

  // Settings first — allowExternalMembers must be on before the cross-tenant
  // member can be added.
  applyForwardingGroupSettings_(oldAddress);

  executeWithRetry(() =>
    AdminDirectory.Members.insert({ email: forwardTo, role: 'MEMBER' }, oldAddress)
  );

  Logger.info('Forwarding group created', {
    address: oldAddress, forwardsTo: forwardTo, groupId: group.id
  });
}

/**
 * Applies the settings a forwarding group needs: external members allowed
 * (deliver to the senior tenant) and anyone-can-post (outsiders can reach it).
 *
 * @param {string} groupEmail
 */
function applyForwardingGroupSettings_(groupEmail) {
  executeWithRetry(() =>
    AdminGroupsSettings.Groups.patch({
      allowExternalMembers: 'true',        // deliver to the cross-tenant senior address
      whoCanPostMessage: 'ANYONE_CAN_POST', // outsiders with the old address can send
      whoCanJoin: 'INVITED_CAN_JOIN',
      whoCanViewMembership: 'ALL_MANAGERS_CAN_VIEW',
      messageModerationLevel: 'MODERATE_NONE',
      spamModerationLevel: 'ALLOW',
      includeInGlobalAddressList: false,
      archiveOnly: false
    }, groupEmail)
  );
}

/**
 * Proves the forwarding-group mechanism end to end with a THROWAWAY group,
 * before a real close depends on it.
 *
 * The real close deletes the account first and only then builds the group, so a
 * domain that forbids external members would fail the forward AFTER an
 * irreversible delete. This exercises the exact three steps createForwardingGroup_
 * does — create group, allow-external + anyone-can-post settings, add a
 * cross-tenant member — against a scratch address, and reports which step fails
 * if any. It does NOT delete anyone.
 *
 * After it runs clean, send a real email to `testAddress` and confirm it lands
 * at `externalMember` — that is the only way to prove actual delivery. Then run
 * `deleteTestForwardingGroup(testAddress)` to clean up.
 *
 * @param {string} testAddress - a throwaway address on THIS (cadets) domain,
 *   e.g. zz-forward-test@cawgcadets.org (must not already exist)
 * @param {string} externalMember - a real senior address on the peer tenant,
 *   e.g. your own @cawgcap.org, to receive the forward
 */
function testForwardingGroup(testAddress, externalMember) {
  if (!testAddress || !externalMember) {
    console.log('Usage: testForwardingGroup("zz-forward-test@cawgcadets.org", "you@cawgcap.org")');
    return;
  }
  console.log('Testing forwarding-group creation (no accounts touched)...');

  try {
    const g = AdminDirectory.Groups.insert({
      email: testAddress, name: 'Forwarding test',
      description: 'Throwaway — delete after testing.'
    });
    console.log('1. group created: ' + g.id);
  } catch (e) {
    console.log('1. FAILED to create group: ' + (e && e.message)); return;
  }

  try {
    applyForwardingGroupSettings_(testAddress);
    console.log('2. settings applied (allowExternalMembers=true, ANYONE_CAN_POST)');
  } catch (e) {
    console.log('2. FAILED to apply settings: ' + (e && e.message));
    console.log('   → the domain may block these group settings. Fix before relying on the forward.');
    return;
  }

  try {
    AdminDirectory.Members.insert({ email: externalMember, role: 'MEMBER' }, testAddress);
    console.log('3. external member added: ' + externalMember);
  } catch (e) {
    console.log('3. FAILED to add external member: ' + (e && e.message));
    console.log('   → the cadets domain likely forbids external group members');
    console.log('     (Admin console → Groups → Sharing settings). This MUST be allowed,');
    console.log('     or the real close will delete an account and then fail to forward it.');
    return;
  }

  console.log('');
  console.log('All three steps passed. Now send a test email to ' + testAddress);
  console.log('and confirm it arrives at ' + externalMember + ' — that proves delivery.');
  console.log('Then clean up: deleteTestForwardingGroup("' + testAddress + '")');
}

/** Removes a throwaway forwarding-test group. */
function deleteTestForwardingGroup(testAddress) {
  AdminDirectory.Groups.remove(testAddress);
  console.log('Deleted test group ' + testAddress);
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
function expireForwardingGroups(dryRun) {
  if (TRANSITION_CONFIG.ROLE !== 'source') return { removed: 0 };
  if (dryRun !== false) return expireForwardingGroups_(true);
  return withTransitionLock_(() => expireForwardingGroups_(false), { removed: 0 });
}

function expireForwardingGroups_(isDry) {
  const rows = readTransitions_();
  const now = new Date();
  let removed = 0;

  for (const capid in rows) {
    const row = rows[capid];
    if (!row.ForwardGroupCreated || !row.ForwardGroupExpires) continue;

    const expires = new Date(row.ForwardGroupExpires);
    if (now < expires) continue;

    if (isDry) {
      console.log(`${capid} | ${row.Name} | WOULD remove forwarding group ${row.CadetEmail}`);
      removed++;
      continue;
    }

    try {
      executeWithRetry(() => AdminDirectory.Groups.remove(row.CadetEmail));
      setTransitionField_(row._rowNumber, 'ForwardGroupExpires', '');
      setTransitionField_(row._rowNumber, 'Notes',
        `Forwarding group removed ${new Date().toISOString()}`);
      Logger.info('Forwarding group expired', { address: row.CadetEmail });
      removed++;
    } catch (e) {
      Logger.error('Could not remove forwarding group', {
        address: row.CadetEmail,
        errorMessage: e && e.message ? e.message : String(e)
      });
    }
  }

  console.log(isDry
    ? `DRY RUN — ${removed} groups would be removed. Pass false to do it.`
    : `Removed ${removed} forwarding groups.`);

  return { removed: removed };
}
