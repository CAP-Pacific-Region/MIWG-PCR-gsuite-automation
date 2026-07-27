/**
 * SquadronGroups.gs — which group settings the sync is allowed to reconcile.
 *
 * This one setting has been wrong in both directions. As a log-only stub it
 * applied nothing, so allowExternalMembers stayed false and cross-tenant nesting
 * failed silently. Made to apply the callers' whole settings block (v1.2.8) it
 * would have pushed the cadet receive lists from ANYONE_CAN_POST down to
 * ALL_MEMBERS_CAN_POST and re-broken the same delivery from the other side.
 *
 * So the managed key list is the thing under test, not an implementation detail:
 * every key in it is reconciled on every sync across all three tenants, and every
 * key out of it is left to whatever console/GAM set. These assertions pin both
 * halves, and pin that the enforced posting value only ever widens access.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'squadron-groups', 'SquadronGroups.gs');
const { section, check, done } = makeChecker();

/**
 * Loads the module with a fake Groups Settings API.
 *
 * @param {Object} live - what AdminGroupsSettings.Groups.get returns
 * @param {Object} [opts] - { dryRun: boolean, available: boolean }
 * @returns {Object} { apply, patches, logs }
 */
function load(live, opts) {
  const o = opts || {};
  const patches = [];
  const { logger, calls: logs } = makeLogger();

  const AdminGroupsSettings = o.available === false ? {} : {
    Groups: {
      get: () => live,
      patch: (patch, email) => { patches.push({ email, patch }); return {}; }
    }
  };

  const m = loadModule(MODULE, {
    Logger: logger,
    CONFIG: { WING: 'CA', EMAIL_DOMAIN: '@example.org' },
    DRY_RUN: !!o.dryRun,
    AdminGroupsSettings: AdminGroupsSettings,
    executeWithRetry: fn => fn(),
    parseFile: () => []
  }, ['applyGroupSettings']);

  return { apply: m.applyGroupSettings, patches, logs };
}

/** The settings block updateDistributionLists() passes for every managed list. */
const CALLER_SETTINGS = {
  whoCanJoin: 'INVITED_CAN_JOIN',
  whoCanViewMembership: 'ALL_MEMBERS_CAN_VIEW',
  whoCanViewGroup: 'ALL_MEMBERS_CAN_VIEW',
  whoCanPostMessage: 'ANYONE_CAN_POST',
  allowExternalMembers: 'true',
  whoCanContactOwner: 'ALL_MEMBERS_CAN_CONTACT',
  messageModerationLevel: 'MODERATE_NONE',
  spamModerationLevel: 'MODERATE',
  enableCollaborativeInbox: 'true',
  includeInGlobalAddressList: 'true',
  replyTo: 'REPLY_TO_SENDER'
};

/** A cadet-side all-hands list as it sits today: internal senders only. */
const INTERNAL_ONLY = {
  whoCanPostMessage: 'ALL_IN_DOMAIN_CAN_POST',
  allowExternalMembers: 'false',
  spamModerationLevel: 'MODERATE',
  whoCanViewMembership: 'ALL_IN_DOMAIN_CAN_VIEW',
  enableCollaborativeInbox: 'false'
};

// ---------------------------------------------------------------------------
section('1. The reported failure: a senior cannot post to the cadet all-hands');
{
  const { apply, patches } = load(INTERNAL_ONLY);
  apply('ca.all@cadets.example.org', CALLER_SETTINGS);

  check('one patch call', patches.length, 1);
  check('posting opens to the other tenant',
    patches[0].patch.whoCanPostMessage, 'ANYONE_CAN_POST');
  check('external membership is still fixed too',
    patches[0].patch.allowExternalMembers, 'true');
  check('the group it was addressed to',
    patches[0].email, 'ca.all@cadets.example.org');
}

// ---------------------------------------------------------------------------
section('2. Only the delivery-governing keys are touched');
{
  const { apply, patches } = load(INTERNAL_ONLY);
  apply('ca101.cadets@cadets.example.org', CALLER_SETTINGS);

  check('visibility is left to console/GAM',
    'whoCanViewMembership' in patches[0].patch, false);
  check('collaborative inbox is left alone',
    'enableCollaborativeInbox' in patches[0].patch, false);
  check('so is everything else the caller passes',
    Object.keys(patches[0].patch).sort(),
    ['allowExternalMembers', 'whoCanPostMessage']);
}

// ---------------------------------------------------------------------------
section('3. Openness never arrives without moderation');
{
  const unmoderated = Object.assign({}, INTERNAL_ONLY, { spamModerationLevel: 'ALLOW' });
  const { apply, patches } = load(unmoderated);
  apply('ca101.all@cadets.example.org', CALLER_SETTINGS);

  check('spam moderation is restored with the open posting',
    patches[0].patch.spamModerationLevel, 'MODERATE');
}

// ---------------------------------------------------------------------------
section('4. A correct group is not written to');
{
  const correct = {
    whoCanPostMessage: 'ANYONE_CAN_POST',
    allowExternalMembers: 'true',
    spamModerationLevel: 'MODERATE'
  };
  const { apply, patches, logs } = load(correct);
  apply('ca101.parents@cadets.example.org', CALLER_SETTINGS);

  check('no API write', patches.length, 0);
  check('and it says why', logs.info.some(l => /already correct/i.test(l.msg)), true);
}

// ---------------------------------------------------------------------------
section('5. Partial drift patches only what drifted');
{
  const drifted = {
    whoCanPostMessage: 'ANYONE_CAN_POST',
    allowExternalMembers: 'false',
    spamModerationLevel: 'MODERATE'
  };
  const { apply, patches } = load(drifted);
  apply('ca101.all@cadets.example.org', CALLER_SETTINGS);

  check('one key only', Object.keys(patches[0].patch), ['allowExternalMembers']);
}

// ---------------------------------------------------------------------------
section('6. Guards');
{
  const dry = load(INTERNAL_ONLY, { dryRun: true });
  dry.apply('ca101.all@cadets.example.org', CALLER_SETTINGS);
  check('dry run writes nothing', dry.patches.length, 0);
  check('dry run still reports the intent',
    dry.logs.info.some(l => /Dry-Run/.test(l.msg)), true);

  const noApi = load(INTERNAL_ONLY, { available: false });
  noApi.apply('ca101.all@cadets.example.org', CALLER_SETTINGS);
  check('a missing advanced service warns rather than throwing',
    noApi.logs.warn.some(l => /not available/i.test(l.msg)), true);

  const blank = load(INTERNAL_ONLY);
  blank.apply('', CALLER_SETTINGS);
  check('no group address, no call', blank.patches.length, 0);

  const empty = load(INTERNAL_ONLY);
  empty.apply('ca101.all@cadets.example.org', { whoCanViewGroup: 'ALL_MEMBERS_CAN_VIEW' });
  check('a settings block with no managed key is a no-op', empty.patches.length, 0);
}

done();
