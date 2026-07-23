/**
 * RecoveryEmailNotify.gs — the compliance predicates, the suppression window,
 * and recipient addressing.
 *
 * These are the decisions that determine whether the module reports the right
 * people, for the right issues, at the right interval, to the right inbox. Each
 * is checked directly here; RecoveryEmailNotify.run.test.js then drives the
 * whole run.
 *
 * Run: npm test
 */
const path = require('path');
const {
  loadModule, makeLogger, makeChecker, Session, Utilities
} = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'notifications', 'RecoveryEmailNotify.gs');
const { section, check, done } = makeChecker();

const CONFIG = {
  DOMAIN: 'cawgcap.org',
  SECONDARY_EMAIL_DOMAIN: '@cawg.cap.gov',
  EMAIL_DOMAIN: '@cawgcap.org',
  COMMAND_EMAIL_DOMAIN: '@cawgcap.org',
  ORG_LABEL: 'CAWG'
};

// Evaluation "today". The account fixtures below are dated relative to this.
const NOW = new Date('2026-07-01T12:00:00Z');
const EPOCH = '1970-01-01T00:00:00.000Z';

/**
 * @param {Object} [overrides] - Globals to replace
 * @returns {Object} Exported internals
 */
function load(overrides) {
  return loadModule(MODULE, Object.assign({
    Logger: makeLogger().logger,
    Session: Session,
    Utilities: Utilities,
    CONFIG: CONFIG,
    ITSUPPORT_EMAIL: 'it.support@cawgcap.org',
    parseFile: () => [],
    createEmailMap: () => ({}),
    getActiveUsers: () => []
  }, overrides || {}), [
    'rcEvaluateMembers_', 'rcIsTenantDomainEmail_', 'rcIsSuppressed_',
    'rcDeriveCommandEmail_', 'rcResolveRecipientEmail_', 'rcBuildRecipientDirectory_',
    'rcBuildDigestHtml_', 'rcSelectAddressees_', 'rcBuildCommandDirectoryMap_',
    'rcBuildAccountMap_', 'rcHasLoggedIn_'
  ]);
}

/**
 * A member as getMembers() hands it over, i.e. AFTER UpdateMembers.addContactInfo
 * has derived recoveryEmail. `recovery` is that derived value — the same field
 * that ends up on the real Workspace account.
 *
 * @param {Object} spec - { capid, primary, recovery, type }
 * @returns {Object} Member object
 */
function member(spec) {
  const m = {
    capsn: spec.capid,
    firstName: 'First' + spec.capid,
    lastName: 'Last' + spec.capid,
    rank: spec.type === 'CADET' ? 'C/SSgt' : 'Capt',
    type: spec.type || 'SENIOR',
    orgid: spec.orgid || '100',
    orgName: 'Alpha Squadron',
    charter: 'PCR-CA-070',
    primaryEmailValue: spec.primary,
    secondaryEmail: spec.secondary
  };
  // A member merged from the ManualMembers sheet never passes through
  // addContactInfo, so recoveryEmail is absent entirely rather than ''.
  if (!spec.manual) m.recoveryEmail = spec.recovery === undefined ? '' : spec.recovery;
  return m;
}

/** A member whose eServices email record is fully compliant. */
function emailOk(capid) {
  return member({ capid, primary: 'm' + capid + '@cawgcap.org', recovery: 'm' + capid + '@gmail.com' });
}

/**
 * A directory account as rcBuildAccountMap_ holds it.
 *
 * @param {Object} spec - { capid, twoSv, lastLogin, created }
 * @returns {Object} Account object
 */
function account(spec) {
  return {
    capid: spec.capid,
    email: 'm' + spec.capid + '@cawgcap.org',
    isEnrolledIn2Sv: spec.twoSv !== false,
    lastLoginTime: spec.lastLogin === undefined ? '2026-06-20T08:00:00.000Z' : spec.lastLogin,
    creationTime: spec.created === undefined ? '2024-01-05T08:00:00.000Z' : spec.created
  };
}

/** Evaluates with the account conditions in play. */
function evaluate(mod, members, accounts) {
  const byCapid = {};
  (accounts || []).forEach(a => { byCapid[String(a.capid)] = a; });
  return mod.rcEvaluateMembers_(members, byCapid, NOW);
}

// ---------------------------------------------------------------------------
section('1. The wing-intended setup with a healthy account is compliant');
{
  const mod = load();
  // CAP address in PRIMARY, personal address in SECONDARY (so recoveryEmail
  // derived to that personal address), account in use with 2SV on.
  const r = evaluate(mod, { '1': emailOk('1') }, [account({ capid: '1' })]);

  check('nobody is flagged', Object.keys(r.flagged).length, 0);
  check('counted compliant', r.compliant, 1);
  check('and evaluated', r.evaluated, 1);
}

// ---------------------------------------------------------------------------
section('2. A personal address in PRIMARY is flagged on primary only');
{
  const mod = load();
  // They CAN recover (personal primary feeds recoveryEmail), so the only thing
  // wrong is which slot it sits in.
  const r = evaluate(mod, {
    '2': member({ capid: '2', primary: 'a.member@gmail.com', recovery: 'a.member@gmail.com' })
  }, [account({ capid: '2' })]);

  check('flagged', Object.keys(r.flagged), ['2']);
  check('for primary only', r.flagged['2'].reasons, 'PRIMARY');
  check('recovery is not flagged', r.flagged['2'].flagRecovery, false);
  check('under the EMAIL category', r.flagged['2'].categories, ['EMAIL']);
}

// ---------------------------------------------------------------------------
section('3. No personal address anywhere is flagged on recovery only');
{
  const mod = load();
  // CAP address in PRIMARY (correct), but nothing personal to recover to.
  const r = evaluate(mod, {
    '3': member({ capid: '3', primary: 'a.member@cawgcap.org', recovery: '' })
  }, [account({ capid: '3' })]);

  check('flagged for recovery only', r.flagged['3'].reasons, 'RECOVERY');
  check('primary is not flagged', r.flagged['3'].flagPrimary, false);
}

// ---------------------------------------------------------------------------
section('4. Both email slots wrong is one row, one EMAIL category');
{
  const mod = load();
  const r = evaluate(mod, {
    '4': member({ capid: '4', primary: 'someone@cawg.cap.gov', recovery: '' })
  }, [account({ capid: '4' })]);

  // Primary IS a tenant address here (secondary domain), so only recovery fails.
  check('the secondary domain counts as a CAP address', r.flagged['4'].reasons, 'RECOVERY');

  const r2 = evaluate(mod, {
    '5': member({ capid: '5', primary: '', recovery: '' })
  }, [account({ capid: '5' })]);
  check('no primary at all + no recovery = both reasons', r2.flagged['5'].reasons, 'PRIMARY+RECOVERY');
  check('reported once, not twice', Object.keys(r2.flagged).length, 1);
  check('and both share the one EMAIL category', r2.flagged['5'].categories, ['EMAIL']);
}

// ---------------------------------------------------------------------------
section('5. A cadet covered by a parent email is not chased for recovery');
{
  // Cadets are evaluated on the CADETS tenant, so their CAP address is on the
  // cadet domain — load that tenant's config rather than the senior one.
  const mod = load({
    CONFIG: Object.assign({}, CONFIG, {
      DOMAIN: 'cawgcadets.org',
      EMAIL_DOMAIN: '@cawgcadets.org',
      SECONDARY_EMAIL_DOMAIN: ''
    })
  });
  // UpdateMembers derives recoveryEmail as
  //   firstPersonalEmail_([secondary, primary]) || parentEmail || ''
  // so a cadet with only a parent address on file still has a working recovery
  // path. Flagging them would bury the real gaps under most of the cadet wing.
  const r = evaluate(mod, {
    '6': member({
      capid: '6', type: 'CADET',
      primary: 'a.cadet@cawgcadets.org',
      secondary: undefined,
      recovery: 'parent@example.com'
    })
  }, [{ capid: '6', email: 'a.cadet@cawgcadets.org', isEnrolledIn2Sv: true, lastLoginTime: '2026-06-20T08:00:00.000Z', creationTime: '2024-01-05T08:00:00.000Z' }]);

  check('not flagged at all', Object.keys(r.flagged).length, 0);
  check('counted compliant', r.compliant, 1);
}

// ---------------------------------------------------------------------------
section('6. Manual-sheet members are skipped, not reported');
{
  const mod = load();
  // Merged after addContactInfo, so they carry no recoveryEmail at all. Read
  // naively every one of them looks like it has neither a CAP primary nor a
  // recovery address — an artefact of where they came from, not something a
  // commander can fix.
  const r = evaluate(mod, {
    '7': member({ capid: '7', manual: true }),
    '8': emailOk('8')
  }, [account({ capid: '8' })]);

  check('the manual member is not flagged', Object.keys(r.flagged).length, 0);
  check('and is counted as skipped', r.skipped, 1);
  check('the real member is still evaluated', r.evaluated, 1);
}

// ---------------------------------------------------------------------------
section('7. An in-use account without 2SV is flagged; enrolled or unused are not');
{
  const mod = load();
  const r = evaluate(mod, {
    '10': emailOk('10'),  // signed in, no 2SV
    '11': emailOk('11'),  // signed in, 2SV on
    '12': emailOk('12')   // never signed in, account 10 days old
  }, [
    account({ capid: '10', twoSv: false }),
    account({ capid: '11' }),
    account({ capid: '12', twoSv: false, lastLogin: EPOCH, created: '2026-06-21T08:00:00.000Z' })
  ]);

  check('the unenrolled in-use account is flagged', r.flagged['10'].reasons, '2SV');
  check('under its own category', r.flagged['10'].categories, ['TWOSV']);
  check('the enrolled one is not', r.flagged['11'], undefined);
  // A 10-day-old unused account trips neither condition: you cannot enroll an
  // account you have never entered, and it is inside the sign-in grace period.
  check('the fresh unused one is not flagged for anything', r.flagged['12'], undefined);
  check('only the one member is flagged', Object.keys(r.flagged), ['10']);
}

// ---------------------------------------------------------------------------
section('8. Never signed in 60+ days after creation is flagged; younger accounts wait');
{
  const mod = load();
  const r = evaluate(mod, {
    '20': emailOk('20'),  // created 2026-04-01, never used — 91 days at NOW
    '21': emailOk('21'),  // created 2026-05-15, never used — 47 days at NOW
    '22': emailOk('22'),  // no directory account at all
    '23': emailOk('23')   // never used but no creationTime to judge from
  }, [
    account({ capid: '20', twoSv: false, lastLogin: EPOCH, created: '2026-04-01T08:00:00.000Z' }),
    account({ capid: '21', twoSv: false, lastLogin: EPOCH, created: '2026-05-15T08:00:00.000Z' }),
    account({ capid: '23', twoSv: false, lastLogin: '', created: '' })
  ]);

  check('the old unused account is flagged', r.flagged['20'].reasons, 'LOGIN');
  check('under its own category', r.flagged['20'].categories, ['LOGIN']);
  check('with the creation date for the digest', r.flagged['20'].accountCreated, '2026-04-01');
  // NOT also flagged for 2SV: one row saying "has never signed in" already
  // implies the rest.
  check('and not also for 2SV', r.flagged['20'].flagTwoSv, false);
  check('the young unused account waits out its grace period', r.flagged['21'], undefined);
  check('no account means no account conditions', r.flagged['22'], undefined);
  check('no creation date means no judgement', r.flagged['23'], undefined);
}

// ---------------------------------------------------------------------------
section('9. Email and account issues combine into one row with both categories');
{
  const mod = load();
  const r = evaluate(mod, {
    '30': member({ capid: '30', primary: 'a.member@gmail.com', recovery: 'a.member@gmail.com' })
  }, [account({ capid: '30', twoSv: false })]);

  check('one flagged member', Object.keys(r.flagged), ['30']);
  check('carrying both reasons', r.flagged['30'].reasons, 'PRIMARY+2SV');
  check('and both categories', r.flagged['30'].categories, ['EMAIL', 'TWOSV']);
}

// ---------------------------------------------------------------------------
section('10. Duplicate accounts resolve to the one actually in use');
{
  const { rcBuildAccountMap_, rcHasLoggedIn_ } = load();
  // The known duplicate-account problem: one CAPID, several accounts. The 2SV
  // and sign-in facts must describe the account the member actually uses, not
  // an abandoned twin — judging the dead twin would flag people who are fine.
  const map = rcBuildAccountMap_([
    { capid: '40', email: 'dead.twin@cawgcap.org', isEnrolledIn2Sv: false, lastLoginTime: '2024-02-01T08:00:00.000Z', creationTime: '2023-01-01T08:00:00.000Z' },
    { capid: '40', email: 'live.account@cawgcap.org', isEnrolledIn2Sv: true, lastLoginTime: '2026-06-25T08:00:00.000Z', creationTime: '2024-01-01T08:00:00.000Z' },
    { capid: '', email: 'no.capid@cawgcap.org' }
  ]);

  check('the most recently used account wins', map['40'].email, 'live.account@cawgcap.org');
  check('accounts without a CAPID are ignored', Object.keys(map), ['40']);

  check('an epoch lastLoginTime reads as never signed in',
    rcHasLoggedIn_({ lastLoginTime: EPOCH }), false);
  check('an absent one too', rcHasLoggedIn_({ lastLoginTime: '' }), false);
  check('a real sign-in reads as one',
    rcHasLoggedIn_({ lastLoginTime: '2026-06-25T08:00:00.000Z' }), true);
}

// ---------------------------------------------------------------------------
section('11. Tenant-domain matching tolerates the configured "@" and case');
{
  const { rcIsTenantDomainEmail_ } = load();

  check('primary domain', rcIsTenantDomainEmail_('x@cawgcap.org'), true);
  check('secondary domain, configured with a leading @',
    rcIsTenantDomainEmail_('x@cawg.cap.gov'), true);
  check('case-insensitive', rcIsTenantDomainEmail_('X@CAWGCAP.ORG'), true);
  check('a personal address is not a CAP address',
    rcIsTenantDomainEmail_('x@gmail.com'), false);
  check('a lookalike subdomain is not a CAP address',
    rcIsTenantDomainEmail_('x@mail.cawgcap.org'), false);
  check('an absent address is not a CAP address', rcIsTenantDomainEmail_(''), false);
  check('so is undefined', rcIsTenantDomainEmail_(undefined), false);
  check('a malformed address is not a CAP address',
    rcIsTenantDomainEmail_('not-an-address'), false);
}

// ---------------------------------------------------------------------------
section('12. The three-month window opens on the day, not a day early');
{
  const { rcIsSuppressed_ } = load();

  check('same day — suppressed', rcIsSuppressed_('2026-01-15', '2026-01-15'), true);
  check('one month later — still suppressed', rcIsSuppressed_('2026-01-15', '2026-02-15'), true);
  check('a day before three months — still suppressed',
    rcIsSuppressed_('2026-01-15', '2026-04-14'), true);
  check('exactly three months — reportable again',
    rcIsSuppressed_('2026-01-15', '2026-04-15'), false);
  check('later still — reportable', rcIsSuppressed_('2026-01-15', '2026-07-01'), false);
  check('across a year boundary', rcIsSuppressed_('2025-12-01', '2026-03-01'), false);
  check('and just inside it', rcIsSuppressed_('2025-12-01', '2026-02-28'), true);
  check('no recorded date means nothing to suppress',
    rcIsSuppressed_(undefined, '2026-01-15'), false);
}

// ---------------------------------------------------------------------------
section('13. Command staff are addressed at their CAP account');
{
  const { rcDeriveCommandEmail_, rcResolveRecipientEmail_ } = load();

  check('first.last on the command domain',
    rcDeriveCommandEmail_({ firstName: 'Ada', lastName: 'Alpha' }), 'ada.alpha@cawgcap.org');
  check('lowercased and space-stripped',
    rcDeriveCommandEmail_({ firstName: 'Mary Jo', lastName: 'Van Dyke' }),
    'maryjo.vandyke@cawgcap.org');
  check('an unusable name derives nothing',
    rcDeriveCommandEmail_({ firstName: '', lastName: 'Alpha' }), null);

  // The CAPWATCH primary is a last resort, never a preference: a wrong or
  // personal primary is the very thing this module exists to report.
  check('falls back to the CAPWATCH primary only when the name is unusable',
    rcResolveRecipientEmail_({ firstName: '', lastName: '' }, '900001',
      { '900001': 'fallback@example.com' }), 'fallback@example.com');
  check('otherwise the derived account wins over the CAPWATCH primary',
    rcResolveRecipientEmail_({ firstName: 'Ada', lastName: 'Alpha' }, '900001',
      { '900001': 'personal@gmail.com' }), 'ada.alpha@cawgcap.org');
  check('and with neither, there is nobody to mail',
    rcResolveRecipientEmail_({}, '900009', {}), null);
}

// ---------------------------------------------------------------------------
section('14. On a cadet tenant, command staff are addressed on the SENIOR domain');
{
  // The reason CONFIG.COMMAND_EMAIL_DOMAIN exists at all: a cadet unit's
  // commander is a senior with no cadet-domain account, so deriving on the
  // tenant's own domain would address an account that does not exist.
  const { rcDeriveCommandEmail_ } = load({
    CONFIG: Object.assign({}, CONFIG, {
      DOMAIN: 'cawgcadets.org',
      EMAIL_DOMAIN: '@cawgcadets.org',
      SECONDARY_EMAIL_DOMAIN: '',
      COMMAND_EMAIL_DOMAIN: '@cawgcap.org'
    })
  });

  check('addressed on the senior domain, not the cadet one',
    rcDeriveCommandEmail_({ firstName: 'Ada', lastName: 'Alpha' }), 'ada.alpha@cawgcap.org');
}

// ---------------------------------------------------------------------------
section('15. The recipient directory collects CC, personnel officers and deputies');
{
  // Commanders.txt: ORGID=0, CAPID=4, NameLast=8, NameFirst=9, Rank=12
  const COMMANDERS = [
    ['100', 'PCR', 'CA', '070', '900001', '', '', '', 'Alpha', 'Ada', '', '', 'Maj']
  ];
  // DutyPosition.txt: CAPID=0, Duty=1, Level=3, Asst=4, ORGID=7
  const DUTIES = [
    ['900002', 'Personnel Officer', '', 'UNIT', '0', '', '', '100'],
    ['900003', 'Personnel Officer', '', 'UNIT', '1', '', '', '100'],
    ['900004', 'Deputy Commander', '', 'UNIT', '0', '', '', '100'],
    ['900005', 'Safety Officer', '', 'UNIT', '0', '', '', '100'],
    ['900006', 'Personnel Officer', '', 'UNIT', '0', '', '', '200']
  ];
  const MEMBER_ROWS = [
    ['900002', '', 'Papa', 'Pat', '', '', '', '', '', '', '', '100', '', '', 'Capt'],
    ['900003', '', 'Quebec', 'Quinn', '', '', '', '', '', '', '', '100', '', '', '1st Lt'],
    ['900004', '', 'Delta', 'Dana', '', '', '', '', '', '', '', '100', '', '', 'Maj'],
    ['900005', '', 'Sierra', 'Sam', '', '', '', '', '', '', '', '100', '', '', 'Capt'],
    ['900006', '', 'Romeo', 'Rio', '', '', '', '', '', '', '', '200', '', '', 'Capt']
  ];

  const { rcBuildRecipientDirectory_ } = load({
    parseFile: name => {
      if (name === 'Commanders') return COMMANDERS;
      if (name === 'DutyPosition') return DUTIES;
      if (name === 'Member') return MEMBER_ROWS;
      return [];
    }
  });

  const byOrg = rcBuildRecipientDirectory_([]);
  const org100 = byOrg['100'].map(r => r.email).sort();

  check('the commander is included',
    org100.indexOf('ada.alpha@cawgcap.org') > -1, true);
  check('so is the personnel officer', org100.indexOf('pat.papa@cawgcap.org') > -1, true);
  check('and the ASSISTANT personnel officer',
    org100.indexOf('quinn.quebec@cawgcap.org') > -1, true);
  check('and the deputy commander', org100.indexOf('dana.delta@cawgcap.org') > -1, true);
  check('unrelated duties are not', org100.indexOf('sam.sierra@cawgcap.org'), -1);
  check('four recipients in all', org100.length, 4);

  check('the assistant is labelled as one',
    byOrg['100'].filter(r => r.capid === '900003')[0].role,
    'Personnel Officer (Assistant)');

  // "Direct command" means the member's own unit: another squadron's personnel
  // officer must never be copied on this one's members.
  check('another unit\'s staff land under that unit',
    byOrg['200'].map(r => r.email), ['rio.romeo@cawgcap.org']);
}

// ---------------------------------------------------------------------------
section('16. The digest explains each issue present, and only those');
{
  const { rcBuildDigestHtml_ } = load();
  const html = rcBuildDigestHtml_(
    { rank: 'Maj', lastName: 'Alpha', email: 'ada.alpha@cawgcap.org' },
    [
      { capid: '111', name: 'Capt A Member', type: 'SENIOR', flagPrimary: true, flagRecovery: false, categories: ['EMAIL'] },
      { capid: '222', name: '2d Lt B Member', type: 'SENIOR', flagPrimary: false, flagRecovery: true, categories: ['EMAIL'] },
      { capid: '333', name: '1st Lt C Member', type: 'SENIOR', flagTwoSv: true, categories: ['TWOSV'] },
      { capid: '444', name: 'SM D Member', type: 'SENIOR', flagLogin: true, accountCreated: '2026-04-01', categories: ['LOGIN'] }
    ]
  );

  check('addresses the commander', html.indexOf('Maj Alpha,') > -1, true);
  check('names the primary-email issue',
    html.indexOf('CAP email is not their PRIMARY email') > -1, true);
  check('names the missing-recovery issue',
    html.indexOf('No personal (non-CAP) email on file') > -1, true);
  check('names the 2SV issue',
    html.indexOf('2-Step Verification is not enabled on their account') > -1, true);
  check('names the never-signed-in issue, with the creation date',
    html.indexOf('Has never signed in to their account (created 2026-04-01)') > -1, true);
  check('lists all four members',
    ['111', '222', '333', '444'].every(id => new RegExp('<td>' + id + '</td>').test(html)), true);
  check('tells them to contact the member',
    html.indexOf('contact these members') > -1, true);
  // Asserted as the exact anchor target, not a bare substring: stricter (the
  // URL must be where a reader will click it, not merely somewhere in the
  // body), and it keeps CodeQL from reading an indexOf-on-a-URL as an
  // incomplete sanitization check.
  check('and that they may fix the email record themselves',
    /<a href="https:\/\/www\.capnhq\.gov\/CAP\.PersonnelInfo\.Web\/">/.test(html), true);
  check('points 2SV enrollment at the right place',
    /<a href="https:\/\/myaccount\.google\.com\/signinoptions\/two-step-verification">/.test(html), true);
  check('and locked-out members at the support portal for an exemption',
    /<a href="https:\/\/support\.pcrcap\.org">/.test(html), true);
  check('explains the three-month quiet period',
    html.indexOf('not be reported again for 3 months') > -1, true);
  check('uses the configured wing label, not a hard-coded one',
    html.indexOf('CAWG account') > -1, true);

  // A digest that is all 2SV must not open with a lecture about eServices email
  // slots — the guidance blocks follow the issues actually present.
  const twoSvOnly = rcBuildDigestHtml_(
    { rank: 'Maj', lastName: 'Alpha' },
    [{ capid: '333', name: '1st Lt C Member', type: 'SENIOR', flagTwoSv: true, categories: ['TWOSV'] }]
  );
  check('a 2SV-only digest offers 2SV guidance',
    twoSvOnly.indexOf('two-step-verification') > -1, true);
  check('and no eServices guidance',
    twoSvOnly.indexOf('CAP.PersonnelInfo.Web') === -1, true);
}

// ---------------------------------------------------------------------------
section('17. A suppressed issue stays out of the digest row');
{
  const { rcBuildDigestHtml_ } = load();
  // The member currently trips EMAIL and TWOSV, but EMAIL is inside its window:
  // the run sets reportCategories to what it is actually reporting, and the
  // digest must follow that — re-listing the suppressed issue every month is
  // exactly what the window exists to prevent.
  const html = rcBuildDigestHtml_(
    { rank: 'Maj', lastName: 'Alpha' },
    [{
      capid: '555', name: 'Capt E Member', type: 'SENIOR',
      flagPrimary: true, flagTwoSv: true,
      categories: ['EMAIL', 'TWOSV'], reportCategories: ['TWOSV']
    }]
  );

  check('the reported issue is shown',
    html.indexOf('2-Step Verification is not enabled') > -1, true);
  check('the suppressed one is not',
    html.indexOf('CAP email is not their PRIMARY email'), -1);
  check('nor its guidance block', html.indexOf('CAP.PersonnelInfo.Web'), -1);
}

// ---------------------------------------------------------------------------
section('18. CAPWATCH-sourced text is escaped');
{
  const { rcBuildDigestHtml_ } = load();
  const html = rcBuildDigestHtml_(
    { rank: 'Maj', lastName: 'Alpha' },
    [{ capid: '1', name: '<script>alert(1)</script>', type: 'SENIOR', flagPrimary: true, flagRecovery: false, categories: ['EMAIL'] }]
  );

  check('no raw script tag survives', html.indexOf('<script>'), -1);
  check('it is escaped instead', html.indexOf('&lt;script&gt;') > -1, true);
}

// ---------------------------------------------------------------------------
section('19. One person holding several duties is mailed once');
{
  const { rcSelectAddressees_ } = load();
  // Not hypothetical: on the first live run one small unit's commander held four
  // of these duties and appeared four times in the raw list, and another held
  // three. The raw duty list repeats them; the run must not. This is the
  // reduction the PREVIEW also runs, so what an operator checks is what is sent.
  const selected = rcSelectAddressees_([
    { email: 'kelly.kilo@cawgcap.org', role: 'Commander' },
    { email: 'val.victor@cawgcap.org', role: 'Personnel Officer' },
    { email: 'kelly.kilo@cawgcap.org', role: 'Deputy Commander' },
    { email: 'KELLY.KILO@cawgcap.org', role: 'Personnel Officer (Assistant)' }
  ]);

  check('the commander is the addressee', selected.addressee.email, 'kelly.kilo@cawgcap.org');
  check('and is not also copied', selected.cc, ['val.victor@cawgcap.org']);
  check('case-insensitive dedupe', selected.cc.length, 1);
}

// ---------------------------------------------------------------------------
section('20. With no commander, the remaining staff still get it');
{
  const { rcSelectAddressees_ } = load();
  const selected = rcSelectAddressees_([
    { email: 'pat.papa@cawgcap.org', role: 'Personnel Officer' },
    { email: 'dana.delta@cawgcap.org', role: 'Deputy Commander' }
  ]);

  check('the first remaining staffer is addressed', selected.addressee.email, 'pat.papa@cawgcap.org');
  check('the rest are copied', selected.cc, ['dana.delta@cawgcap.org']);
}

// ---------------------------------------------------------------------------
section('21. The real account beats the derived one');
{
  // Derivation reproduces the DEFAULT account name, so it is wrong for exactly
  // the people whose account is not the default — a .2 duplicate, a manual
  // creation, a rename. Those would otherwise be mailed into the void.
  const { rcResolveRecipientEmail_ } = load();
  const info = { firstName: 'Ada', lastName: 'Alpha' };

  check('directory wins over derivation',
    rcResolveRecipientEmail_(info, '900001', {}, { '900001': 'ada.alpha.2@cawgcap.org' }),
    'ada.alpha.2@cawgcap.org');
  check('derivation still used when the directory has no entry',
    rcResolveRecipientEmail_(info, '900001', {}, {}), 'ada.alpha@cawgcap.org');
  check('CAPWATCH primary remains the last resort',
    rcResolveRecipientEmail_({}, '900001', { '900001': 'fallback@example.com' }, {}),
    'fallback@example.com');
}

// ---------------------------------------------------------------------------
section('22. The command map is only built where command staff actually live');
{
  // The run fetches the directory once (the account conditions need it) and
  // hands the listing in — this helper never issues its own read.
  const seniors = load();
  check('seniors reads the passed listing',
    seniors.rcBuildCommandDirectoryMap_([{ capid: '900001', email: 'real.address@cawgcap.org' }]),
    { '900001': 'real.address@cawgcap.org' });

  // Cadets: command staff are seniors in the OTHER tenant's directory. The
  // passed listing is this tenant's own CADET accounts — a hit there would be
  // some cadet-domain account, not the mailbox the commander reads.
  const cadets = load({
    CONFIG: Object.assign({}, CONFIG, {
      EMAIL_DOMAIN: '@cawgcadets.org',
      COMMAND_EMAIL_DOMAIN: '@cawgcap.org'
    })
  });
  check('cadets ignores the local listing',
    cadets.rcBuildCommandDirectoryMap_([{ capid: '900001', email: 'some.cadet@cawgcadets.org' }]),
    {});
}

done();
