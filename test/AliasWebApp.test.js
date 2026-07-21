/**
 * webapp/ — the Secondary Alias admin web app.
 *
 * Three things are worth pinning here, and they are the three that fail silently
 * in production:
 *
 *   1. The address arithmetic is DUPLICATED from src/ (the web app is a separate
 *      script project). If the two copies ever disagree about what a member's
 *      alias is, the symptom is a sheet row whose status flips every night, with
 *      nothing in either log saying why. So both copies are run over the same
 *      inputs and required to agree.
 *   2. The removal guards. This is the only code in the codebase that deletes a
 *      directory alias; a bug here silently strips a working address.
 *   3. The authorization gate fails CLOSED — no group configured, no caller, or a
 *      Directory outage must all deny, never allow.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const { section, check, done } = makeChecker();

const WEBAPP = f => path.join(__dirname, '..', 'webapp', f);
const SRC = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'SecondaryDomainAliases.gs');

const WEBAPP_CAPID_RE = /^\d{5,7}$/;
const WEBAPP_CONFLICT_PREFIX = 'CONFLICT';

// ---------------------------------------------------------------------------
section('1. address arithmetic — the webapp copy must match src/ exactly');
{
  const src = loadModule(SRC, { Logger: makeLogger().logger, AdminDirectory: {}, CONFIG: {} },
    ['normalizeSecondaryDomain_', 'deriveSecondaryAlias_']);
  const allow = () => 'it@cawgcap.org';
  const web = loadModule(WEBAPP('Directory.gs'),
    { Logger: makeLogger().logger, AdminDirectory: {}, WEBAPP_CAPID_RE, WEBAPP_CONFLICT_PREFIX,
      requireAuthorized_: allow },
    ['webappNormalizeSecondaryDomain_', 'webappDeriveSecondaryAlias_', 'webappProvisioningCapidsFromUser_',
     'webappRankAccounts_', 'webappRemoveAliasFromAccount_']);

  // Every form the property has actually been set to, plus the ones that would
  // quietly mint garbage addresses. '@'-less was the real bug on seniors: the
  // derive step concatenates, so a bare domain yielded 'jane.doecawg.cap.gov'.
  const domainInputs = [
    '@cawg.cap.gov', 'cawg.cap.gov', '  CAWG.Cap.Gov  ', '@@cawg.cap.gov',
    '', '   ', 'cawg', 'user@cawg.cap.gov', null, undefined
  ];
  domainInputs.forEach(function (input) {
    check('normalize ' + JSON.stringify(input),
      web.webappNormalizeSecondaryDomain_(input),
      src.normalizeSecondaryDomain_(input));
  });

  check('a dotless value is refused, not turned into 34 bad addresses',
    web.webappNormalizeSecondaryDomain_('cawg'), '');
  check('an address is refused where a domain was expected',
    web.webappNormalizeSecondaryDomain_('user@cawg.cap.gov'), '');

  const aliasInputs = [
    ['jane.doe@cawgcap.org', '@cawg.cap.gov'],
    ['JANE.DOE@cawgcap.org', '@cawg.cap.gov'],
    ['lu-ann.fernandez@cawgcap.org', '@cawg.cap.gov'],
    ['sam.roe.2@cawgcap.org', '@cawg.cap.gov']
  ];
  aliasInputs.forEach(function (pair) {
    check('derive ' + pair[0],
      web.webappDeriveSecondaryAlias_(pair[0], pair[1]),
      src.deriveSecondaryAlias_(pair[0].toLowerCase(), pair[1]));
  });

  // Only the webapp copy is asked about these: src/ never sees a malformed address
  // because it is read from a curated column, whereas the webapp derives from
  // whatever the directory returns.
  check('no localpart yields nothing rather than a bare domain',
    web.webappDeriveSecondaryAlias_('@cawgcap.org', '@cawg.cap.gov'), '');
  check('no domain configured yields nothing',
    web.webappDeriveSecondaryAlias_('jane.doe@cawgcap.org', ''), '');

  // -------------------------------------------------------------------------
  section('2. webappRemoveAliasFromAccount_ — the guards on the only destructive call');
  {
    const account = {
      email: 'jane.doe@cawgcap.org',
      aliases: ['jane.doe@cawg.cap.gov', 'jdoe@cawgcap.org']
    };
    let removed = [];
    const web2 = loadModule(WEBAPP('Directory.gs'), {
      Logger: makeLogger().logger,
      AdminDirectory: { Users: { Aliases: { remove: (u, a) => removed.push([u, a]) } } },
      WEBAPP_CAPID_RE, WEBAPP_CONFLICT_PREFIX, requireAuthorized_: allow
    }, ['webappRemoveAliasFromAccount_']);

    const r1 = web2.webappRemoveAliasFromAccount_(account, 'jane.doe@cawg.cap.gov', '@cawg.cap.gov');
    check('the happy path revokes', [r1.removed, removed.length], [true, 1]);
    check('and revokes exactly the address asked for', removed[0], ['jane.doe@cawgcap.org', 'jane.doe@cawg.cap.gov']);

    // THE important one: this app must never strip an address on the PRIMARY
    // domain. It did not create those and has no way to know what they are for.
    removed = [];
    const r2 = web2.webappRemoveAliasFromAccount_(account, 'jdoe@cawgcap.org', '@cawg.cap.gov');
    check('refuses an alias outside the secondary domain', r2.removed, false);
    check('and calls nothing', removed.length, 0);
    check('and says why', r2.status.indexOf('REFUSED') === 0, true);

    removed = [];
    const r3 = web2.webappRemoveAliasFromAccount_(account, 'jane.doe@cawgcap.org', '@cawg.cap.gov');
    check('refuses the account primary address', [r3.removed, removed.length], [false, 0]);

    removed = [];
    const r4 = web2.webappRemoveAliasFromAccount_(account, 'someone.else@cawg.cap.gov', '@cawg.cap.gov');
    check('refuses an address the account does not hold', [r4.removed, removed.length], [false, 0]);
    check('reported as NOT PRESENT, not an error — a stale row is normal',
      r4.status.indexOf('NOT PRESENT') === 0, true);

    removed = [];
    const r5 = web2.webappRemoveAliasFromAccount_(account, 'jane.doe@cawg.cap.gov', '');
    check('an unconfigured secondary domain revokes nothing (fails closed)',
      [r5.removed, removed.length], [false, 0]);

    // A suffix check written as indexOf() would match mid-string; this is the
    // lookalike domain that would slip through.
    removed = [];
    const r6 = web2.webappRemoveAliasFromAccount_(
      { email: 'jane.doe@cawgcap.org', aliases: ['jane.doe@cawg.cap.gov.evil.com'] },
      'jane.doe@cawg.cap.gov.evil.com', '@cawg.cap.gov');
    check('a lookalike domain that merely CONTAINS the secondary domain is refused',
      [r6.removed, removed.length], [false, 0]);

    // Defense in depth. The destructive call takes a caller-supplied account
    // object, so it re-checks authorization itself rather than trusting that
    // Apps Script keeps trailing-underscore functions out of google.script.run.
    removed = [];
    const denied = loadModule(WEBAPP('Directory.gs'), {
      Logger: makeLogger().logger,
      AdminDirectory: { Users: { Aliases: { remove: (u, a) => removed.push([u, a]) } } },
      WEBAPP_CAPID_RE, WEBAPP_CONFLICT_PREFIX,
      requireAuthorized_: () => { throw new Error('You are not authorized to manage secondary aliases.'); }
    }, ['webappRemoveAliasFromAccount_', 'webappAddAliasToAccount_']);

    let removeThrew = false;
    try { denied.webappRemoveAliasFromAccount_(account, 'jane.doe@cawg.cap.gov', '@cawg.cap.gov'); }
    catch (e) { removeThrew = true; }
    check('an unauthorized caller cannot revoke, even reaching the helper directly',
      [removeThrew, removed.length], [true, 0]);

    let addThrew = false;
    try { denied.webappAddAliasToAccount_('jane.doe@cawgcap.org', 'jane.doe@cawg.cap.gov'); }
    catch (e) { addThrew = true; }
    check('nor add', addThrew, true);
  }

  // -------------------------------------------------------------------------
  section('3. CAPID carriers — retired markers must not resolve to an account');
  {
    check('an organization externalId counts',
      web.webappProvisioningCapidsFromUser_({ externalIds: [{ type: 'organization', value: '123456' }] }),
      ['123456']);
    check('the retired-duplicate marker does NOT',
      web.webappProvisioningCapidsFromUser_({
        externalIds: [{ type: 'custom', customType: 'duplicate_retired_capid', value: '123456' }]
      }), []);
    check('a 4-digit value is not a CAPID',
      web.webappProvisioningCapidsFromUser_({ employeeId: '1234' }), []);
  }

  // -------------------------------------------------------------------------
  section('4. webappRankAccounts_ — the account in use is offered first');
  {
    // Synthetic names. The shape is the real one from the cadets tenant: the
    // canonical-looking twin is the dead one.
    const ranked = web.webappRankAccounts_([
      { email: 'sam.roe@x.org', suspended: false, created: '2026-01-23T00:00:00Z', neverSignedIn: true },
      { email: 'sam.roe.2@x.org', suspended: false, created: '2025-11-24T00:00:00Z', lastLogin: '2026-07-15T00:00:00Z', neverSignedIn: false }
    ]);
    check('the signed-into account ranks above the canonical-but-dead twin',
      ranked[0].email, 'sam.roe.2@x.org');
    check('both are still offered — the human picks', ranked.length, 2);
  }
}

// ---------------------------------------------------------------------------
section('5. isAliasAdmin_ — fails closed on every uncertainty');
{
  function auth(config, adminDirectory) {
    return loadModule(WEBAPP('Auth.gs'), {
      Logger: makeLogger().logger,
      Session: { getActiveUser: () => ({ getEmail: () => 'jane.doe@cawgcap.org' }) },
      WEBAPP_CONFIG: config,
      AdminDirectory: adminDirectory
    }, ['isAliasAdmin_', 'resolveActor_', 'requireAuthorized_']);
  }

  const yes = { Members: { hasMember: () => ({ isMember: true }) } };
  const no = { Members: { hasMember: () => ({ isMember: false }) } };
  const boom = { Members: { hasMember: () => { throw new Error('Resource Not Found: groupKey'); } } };

  check('a member of the group is allowed',
    auth({ ADMIN_GROUP: 'admins@x.org' }, yes).isAliasAdmin_('jane.doe@cawgcap.org'), true);
  check('a non-member is denied',
    auth({ ADMIN_GROUP: 'admins@x.org' }, no).isAliasAdmin_('jane.doe@cawgcap.org'), false);

  // Each of these has been a real outage mode in some other system: the config is
  // missing, the identity is missing, or Google is having a bad day. None may open.
  check('NO GROUP CONFIGURED denies everyone (an unconfigured tenant is locked, not open)',
    auth({ ADMIN_GROUP: '' }, yes).isAliasAdmin_('jane.doe@cawgcap.org'), false);
  check('an empty caller is denied even when the group check would say yes',
    auth({ ADMIN_GROUP: 'admins@x.org' }, yes).isAliasAdmin_(''), false);
  check('a Directory error denies rather than allows',
    auth({ ADMIN_GROUP: 'admins@x.org' }, boom).isAliasAdmin_('jane.doe@cawgcap.org'), false);

  // A malformed API response must not be read as truthy-enough.
  check('a response without isMember:true is denied',
    auth({ ADMIN_GROUP: 'admins@x.org' }, { Members: { hasMember: () => ({}) } })
      .isAliasAdmin_('jane.doe@cawgcap.org'), false);

  let threw = false;
  try {
    auth({ ADMIN_GROUP: 'admins@x.org' }, no).requireAuthorized_();
  } catch (e) {
    threw = true;
    check('the refusal does not name the group (nothing to escalate against)',
      e.message.indexOf('admins@x.org'), -1);
  }
  check('requireAuthorized_ throws rather than returning a flag a caller can forget', threw, true);

  const okActor = auth({ ADMIN_GROUP: 'admins@x.org' }, yes).requireAuthorized_();
  check('and returns the actor for attribution', okActor, 'jane.doe@cawgcap.org');
}

// ---------------------------------------------------------------------------
section('6. row placement — the web app never shifts a row under the nightly run');
{
  const api = loadModule(WEBAPP('AliasAdminApi.gs'), { Logger: makeLogger().logger },
    ['webappFindRowByEmail_', 'webappClaimRow_', 'webappCell_']);

  // A tab with a blank row in the middle, left behind by a removal.
  const values = [
    ['Primary Email', 'Alias Email (override)', 'Status', 'Last Run', 'CAPID', 'Added By'],
    ['jane.doe@cawgcap.org', '', 'ADDED — jane.doe@cawg.cap.gov', '', '123456', 'it@cawgcap.org'],
    ['', '', '', '', '', ''],
    ['sam.roe@cawgcap.org', '', 'OK — already present', '', '234567', 'it@cawgcap.org']
  ];
  const sheet = { getDataRange: () => ({ getValues: () => values }), getLastRow: () => values.length };

  check('finds a listed account by address (case-insensitively)',
    api.webappFindRowByEmail_(sheet, 'jane.doe@cawgcap.org'), 2);
  check('an unlisted account is 0, not a false row', api.webappFindRowByEmail_(sheet, 'nobody@cawgcap.org'), 0);

  // The point of claimRow_: adds refill the hole instead of growing the tab, and
  // crucially never insert a row that would shift the ones below it.
  check('an add reuses the blank row left by a removal', api.webappClaimRow_(sheet), 3);

  const full = { getDataRange: () => ({ getValues: () => values.filter(r => r[0]) }), getLastRow: () => 3 };
  check('with no blanks it appends at the bottom', api.webappClaimRow_(full), 4);

  check('a short row reads as empty, not undefined', api.webappCell_(['a'], 5), '');
}

done();
