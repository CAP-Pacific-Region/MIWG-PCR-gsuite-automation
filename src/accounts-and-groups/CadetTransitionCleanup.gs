/**
 * Cadet → senior transition: closing the old account.
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

  const isDry = dryRun !== false;
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

  if (TRANSITION_CONFIG.REQUIRE_MIGRATION_BEFORE_DELETE && !row.MessagesMigrated) {
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
  // Refuses until DriveMigrated says otherwise. '0' is a legitimate answer for a
  // member with an empty Drive and must be set deliberately; blank means nobody
  // has looked, which is not the same thing.
  if (String(row.DriveMigrated || '') === '') {
    return 'Drive not handled — closing would destroy it. Copy it, or set ' +
      'DriveMigrated=0 if there is nothing to copy';
  }

  // Same reasoning as Drive: personal contacts die with the account. Blank means
  // nobody looked; 0 is a deliberate "nothing to copy".
  if (String(row.ContactsMigrated || '') === '') {
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
 * ANYONE_CAN_POST is required and deliberate: the whole point is that outsiders
 * who still have the old address can reach the member. A default group rejects
 * external senders, which would forward nothing and bounce everyone.
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

  executeWithRetry(() =>
    AdminDirectory.Members.insert({ email: forwardTo, role: 'MEMBER' }, oldAddress)
  );

  // Without this the group rejects external mail and forwards nothing.
  try {
    executeWithRetry(() =>
      AdminGroupsSettings.Groups.patch({
        whoCanPostMessage: 'ANYONE_CAN_POST',
        whoCanJoin: 'INVITED_CAN_JOIN',
        whoCanViewMembership: 'ALL_MANAGERS_CAN_VIEW',
        messageModerationLevel: 'MODERATE_NONE',
        spamModerationLevel: 'ALLOW',
        includeInGlobalAddressList: false,
        archiveOnly: false
      }, oldAddress)
    );
  } catch (e) {
    // The group exists and forwards to a member either way; settings are what
    // let strangers reach it. Loud, but not worth unwinding a deletion over.
    Logger.error('Forwarding group created but settings failed — external mail ' +
      'will be REJECTED until whoCanPostMessage is fixed by hand', {
      group: oldAddress,
      errorMessage: e && e.message ? e.message : String(e)
    });
  }

  Logger.info('Forwarding group created', {
    address: oldAddress, forwardsTo: forwardTo, groupId: group.id
  });
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

  const isDry = dryRun !== false;
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
