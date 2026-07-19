/**
 * RecoveryEmailNotify.gs — the compliance predicate, the suppression window, and
 * recipient addressing.
 *
 * These are the three decisions that determine whether the module reports the
 * right people, at the right interval, to the right inbox. Each is checked
 * directly here; RecoveryEmailNotify.run.test.js then drives the whole run.
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
    'rcBuildDigestHtml_', 'rcSelectAddressees_', 'rcBuildCommandDirectoryMap_'
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

// ---------------------------------------------------------------------------
section('1. The wing-intended setup is compliant');
{
  const { rcEvaluateMembers_ } = load();
  // CAP address in PRIMARY, personal address in SECONDARY (so recoveryEmail
  // derived to that personal address).
  const r = rcEvaluateMembers_({
    '1': member({ capid: '1', primary: 'a.member@cawgcap.org', recovery: 'a.member@gmail.com' })
  });

  check('nobody is flagged', Object.keys(r.flagged).length, 0);
  check('counted compliant', r.compliant, 1);
  check('and evaluated', r.evaluated, 1);
}

// ---------------------------------------------------------------------------
section('2. A personal address in PRIMARY is flagged on primary only');
{
  const { rcEvaluateMembers_ } = load();
  // They CAN recover (personal primary feeds recoveryEmail), so the only thing
  // wrong is which slot it sits in.
  const r = rcEvaluateMembers_({
    '2': member({ capid: '2', primary: 'a.member@gmail.com', recovery: 'a.member@gmail.com' })
  });

  check('flagged', Object.keys(r.flagged), ['2']);
  check('for primary only', r.flagged['2'].reasons, 'PRIMARY');
  check('recovery is not flagged', r.flagged['2'].flagRecovery, false);
}

// ---------------------------------------------------------------------------
section('3. No personal address anywhere is flagged on recovery only');
{
  const { rcEvaluateMembers_ } = load();
  // CAP address in PRIMARY (correct), but nothing personal to recover to.
  const r = rcEvaluateMembers_({
    '3': member({ capid: '3', primary: 'a.member@cawgcap.org', recovery: '' })
  });

  check('flagged for recovery only', r.flagged['3'].reasons, 'RECOVERY');
  check('primary is not flagged', r.flagged['3'].flagPrimary, false);
}

// ---------------------------------------------------------------------------
section('4. Both wrong is reported as BOTH, in one row');
{
  const { rcEvaluateMembers_ } = load();
  const r = rcEvaluateMembers_({
    '4': member({ capid: '4', primary: 'someone@cawg.cap.gov', recovery: '' })
  });

  // Primary IS a tenant address here (secondary domain), so only recovery fails.
  check('the secondary domain counts as a CAP address', r.flagged['4'].reasons, 'RECOVERY');

  const r2 = rcEvaluateMembers_({
    '5': member({ capid: '5', primary: '', recovery: '' })
  });
  check('no primary at all + no recovery = BOTH', r2.flagged['5'].reasons, 'BOTH');
  check('reported once, not twice', Object.keys(r2.flagged).length, 1);
}

// ---------------------------------------------------------------------------
section('5. A cadet covered by a parent email is not chased for recovery');
{
  // Cadets are evaluated on the CADETS tenant, so their CAP address is on the
  // cadet domain — load that tenant's config rather than the senior one.
  const { rcEvaluateMembers_ } = load({
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
  const r = rcEvaluateMembers_({
    '6': member({
      capid: '6', type: 'CADET',
      primary: 'a.cadet@cawgcadets.org',
      secondary: undefined,
      recovery: 'parent@example.com'
    })
  });

  check('not flagged at all', Object.keys(r.flagged).length, 0);
  check('counted compliant', r.compliant, 1);
}

// ---------------------------------------------------------------------------
section('6. Manual-sheet members are skipped, not reported');
{
  const { rcEvaluateMembers_ } = load();
  // Merged after addContactInfo, so they carry no recoveryEmail at all. Read
  // naively every one of them looks like it has neither a CAP primary nor a
  // recovery address — an artefact of where they came from, not something a
  // commander can fix.
  const r = rcEvaluateMembers_({
    '7': member({ capid: '7', manual: true }),
    '8': member({ capid: '8', primary: 'ok@cawgcap.org', recovery: 'ok@gmail.com' })
  });

  check('the manual member is not flagged', Object.keys(r.flagged).length, 0);
  check('and is counted as skipped', r.skipped, 1);
  check('the real member is still evaluated', r.evaluated, 1);
}

// ---------------------------------------------------------------------------
section('7. Tenant-domain matching tolerates the configured "@" and case');
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
section('8. The three-month window opens on the day, not a day early');
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
section('9. Command staff are addressed at their CAP account');
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
section('10. On a cadet tenant, command staff are addressed on the SENIOR domain');
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
section('11. The recipient directory collects CC, personnel officers and deputies');
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

  const byOrg = rcBuildRecipientDirectory_();
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
section('12. The digest states the issue, the fix, and the self-service route');
{
  const { rcBuildDigestHtml_ } = load();
  const html = rcBuildDigestHtml_(
    { rank: 'Maj', lastName: 'Alpha', email: 'ada.alpha@cawgcap.org' },
    [
      { capid: '111', name: 'Capt A Member', type: 'SENIOR', flagPrimary: true, flagRecovery: false },
      { capid: '222', name: '2d Lt B Member', type: 'SENIOR', flagPrimary: false, flagRecovery: true }
    ]
  );

  check('addresses the commander', html.indexOf('Maj Alpha,') > -1, true);
  check('names the primary-email issue',
    html.indexOf('CAP email is not their PRIMARY email') > -1, true);
  check('names the missing-recovery issue',
    html.indexOf('No personal (non-CAP) email on file') > -1, true);
  check('lists both members', /<td>111<\/td>/.test(html) && /<td>222<\/td>/.test(html), true);
  check('tells them to contact the member',
    html.indexOf('contact these members') > -1, true);
  check('and that they may fix it themselves',
    html.indexOf('https://www.capnhq.gov/CAP.PersonnelInfo.Web/') > -1, true);
  check('explains the three-month quiet period',
    html.indexOf('not be reported again for 3 months') > -1, true);
  check('uses the configured wing label, not a hard-coded one',
    html.indexOf('CAWG account password') > -1, true);
}

// ---------------------------------------------------------------------------
section('13. CAPWATCH-sourced text is escaped');
{
  const { rcBuildDigestHtml_ } = load();
  const html = rcBuildDigestHtml_(
    { rank: 'Maj', lastName: 'Alpha' },
    [{ capid: '1', name: '<script>alert(1)</script>', type: 'SENIOR', flagPrimary: true, flagRecovery: false }]
  );

  check('no raw script tag survives', html.indexOf('<script>'), -1);
  check('it is escaped instead', html.indexOf('&lt;script&gt;') > -1, true);
}

// ---------------------------------------------------------------------------
section('14. One person holding several duties is mailed once');
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
section('15. With no commander, the remaining staff still get it');
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
section('16. The real account beats the derived one');
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
section('17. The directory is only consulted where command staff actually live');
{
  // Seniors: command domain == this tenant's own domain, so its directory is
  // authoritative for commanders.
  const seniors = load({ getActiveUsers: () => [{ capid: '900001', email: 'real.address@cawgcap.org' }] });
  check('seniors reads its own directory',
    seniors.rcBuildCommandDirectoryMap_(), { '900001': 'real.address@cawgcap.org' });

  // Cadets: command staff are seniors in the OTHER tenant's directory, which
  // this script cannot read. A local hit would be some cadet-domain account, not
  // the mailbox the commander reads — so it must not look.
  const cadets = load({
    CONFIG: Object.assign({}, CONFIG, {
      EMAIL_DOMAIN: '@cawgcadets.org',
      COMMAND_EMAIL_DOMAIN: '@cawgcap.org'
    }),
    getActiveUsers: () => { throw new Error('must not be called on the cadets tenant'); }
  });
  check('cadets does not consult the local directory', cadets.rcBuildCommandDirectoryMap_(), {});

  // A directory read that throws must not abort an otherwise-fine run.
  const broken = load({ getActiveUsers: () => { throw new Error('insufficient permission'); } });
  check('a failed directory read degrades to derivation', broken.rcBuildCommandDirectoryMap_(), {});
}

done();
