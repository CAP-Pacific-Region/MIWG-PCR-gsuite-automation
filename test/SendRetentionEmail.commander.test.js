/**
 * SendRetentionEmail.gs — commander CC addressing, and template rendering.
 *
 * Two things worth pinning here.
 *
 * The commander CC goes to the commander's CAP account, falling back to their
 * CAPWATCH personal address only when no org address can be found. That order is
 * `rcResolveRecipientEmail_()` from RecoveryEmailNotify.gs, reused across the
 * module boundary — so this test loads the REAL resolver out of that module
 * rather than faking it. If someone changes the resolution order there, this
 * fails here, which is the point: the two modules are supposed to agree on how
 * you reach a commander.
 *
 * The templates are member-facing and carry no wing of their own. Rendering them
 * against two different tenants catches a hard-coded wing name or role holder
 * creeping back in — the exact regression this module has had before.
 *
 * Run: npm test
 */
const fs = require('fs');
const path = require('path');
const { loadModule, makeLogger, makeChecker, Session, Utilities } = require('./helpers/apps-script');

const SRC = path.join(__dirname, '..', 'src');
const RETENTION = path.join(SRC, 'recruiting-and-retention', 'SendRetentionEmail.gs');
const NOTIFY = path.join(SRC, 'notifications', 'RecoveryEmailNotify.gs');

const { section, check, done } = makeChecker();

const SENIORS = {
  DOMAIN: 'cawgcap.org',
  EMAIL_DOMAIN: '@cawgcap.org',
  COMMAND_EMAIL_DOMAIN: '@cawgcap.org', // command staff are on THIS tenant
  WING_NAME: 'California Wing',
  ORG_LABEL: 'CAWG'
};

const CADETS = {
  DOMAIN: 'cawgcadets.org',
  EMAIL_DOMAIN: '@cawgcadets.org',
  COMMAND_EMAIL_DOMAIN: '@cawgcap.org', // command staff are seniors, elsewhere
  WING_NAME: 'California Wing',
  ORG_LABEL: 'CAWG'
};

const HIWG = {
  DOMAIN: 'hiwgcap.org',
  EMAIL_DOMAIN: '@hiwgcap.org',
  COMMAND_EMAIL_DOMAIN: '@hiwgcap.org',
  WING_NAME: 'Hawaii Wing',
  ORG_LABEL: 'HIWG'
};

// Commanders.txt row: ORGID=0, CAPID=4, NameLast=8, NameFirst=9, Rank=12.
function commanderRow(orgid, capid, last, first, rank) {
  const row = new Array(13).fill('');
  row[0] = orgid;
  row[4] = capid;
  row[8] = last;
  row[9] = first;
  row[12] = rank || 'Maj';
  return row;
}

// MbrContact.txt row: CAPID=0, Type=1, Priority=2, Contact=3.
const contactRow = (capid, email) => [capid, 'EMAIL', 'PRIMARY', email];

// DutyPosition.txt row: CAPID=0, Duty=1, Asst=4, ORGID=7. The trailing space on
// the duty value is not decoration — the live feed ships duty titles padded, so
// matching has to normalize. See formatDutyTitle_ in UpdateMembers.gs.
function dutyRow(capid, duty, orgid, isAssistant) {
  const row = new Array(8).fill('');
  row[0] = capid;
  row[1] = duty + ' ';
  row[4] = isAssistant ? '1' : '0';
  row[7] = orgid;
  return row;
}

// Member.txt row, only the columns the duty-holder name lookup reads:
// CAPID=0, NameLast=2, NameFirst=3, Rank=14.
function memberRow(capid, last, first, rank) {
  const row = new Array(25).fill('');
  row[0] = capid;
  row[2] = last;
  row[3] = first;
  row[14] = rank || 'Capt';
  return row;
}

const CC_DUTY_TITLES = {
  AGE_MILESTONE: ['Deputy Commander for Cadets'],
  RENEWAL: ['Recruiting Officer']
};

// The real normalizer's behavior: trim, collapse whitespace, apply the CAPR 30-1
// renames. Kept in step with DUTY_TITLE_OVERRIDES in UpdateMembers.gs.
const DUTY_TITLE_OVERRIDES = {
  'RECRUITING & RETENTION OFFICER': 'Recruiting Officer',
  'RECRUITING AND RETENTION OFFICER': 'Recruiting Officer'
};
function formatDutyTitle_(dutyId) {
  const title = String(dutyId || '').trim().replace(/\s+/g, ' ');
  return DUTY_TITLE_OVERRIDES[title.toUpperCase()] || title;
}

/** Reads the real template files, so the slash-prefixed name is exercised too. */
const HtmlService = {
  createHtmlOutputFromFile: name => ({
    getContent: () => fs.readFileSync(path.join(SRC, name + '.html'), 'utf8')
  })
};

/**
 * Loads SendRetentionEmail.gs with the REAL address resolver from
 * RecoveryEmailNotify.gs wired in.
 *
 * @param {Object} opts - { config, commanders, contacts, accounts, directorName }
 * @returns {Object} Exported internals, plus counters for the injected fakes
 */
function load(opts) {
  const o = opts || {};
  const config = o.config || SENIORS;
  const counts = { getActiveUsers: 0, parseFile: {} };

  const notify = loadModule(NOTIFY, {
    Logger: makeLogger().logger,
    Session: Session,
    Utilities: Utilities,
    CONFIG: config,
    parseFile: () => [],
    createEmailMap: () => ({}),
    getActiveUsers: () => []
  }, ['rcBuildCommandDirectoryMap_', 'rcResolveRecipientEmail_', 'rcDeriveCommandEmail_']);

  const files = {
    Commanders: o.commanders || [],
    MbrContact: o.contacts || [],
    Member: o.members || [],
    DutyPosition: o.duties || []
  };

  const { logger, calls } = makeLogger();
  const mod = loadModule(RETENTION, {
    Logger: logger,
    Session: Session,
    Utilities: Utilities,
    HtmlService: HtmlService,
    CONFIG: config,
    RETENTION_CONFIG: { CC_DUTY_TITLES: o.ccDutyTitles || CC_DUTY_TITLES },
    formatDutyTitle_: formatDutyTitle_,
    DIRECTOR_RECRUITING_NAME: o.directorName || '',
    sanitizeEmail: e => {
      const v = String(e || '').trim().toLowerCase();
      return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v) ? v : null;
    },
    parseFile: name => {
      counts.parseFile[name] = (counts.parseFile[name] || 0) + 1;
      return files[name] || [];
    },
    getActiveUsers: () => {
      counts.getActiveUsers++;
      if (o.directoryThrows) throw new Error('Insufficient permission');
      return o.accounts || [];
    },
    rcBuildCommandDirectoryMap_: notify.rcBuildCommandDirectoryMap_,
    rcResolveRecipientEmail_: notify.rcResolveRecipientEmail_,
    rcDeriveCommandEmail_: notify.rcDeriveCommandEmail_
  }, ['getCommanderInfo', 'retentionUnitStaffIndex_', 'retentionCcList_', 'retentionRenderTemplate_', 'retentionSignatureHtml_']);

  return Object.assign({}, mod, { counts, logCalls: calls });
}

// ---------------------------------------------------------------------------
section('1. Seniors tenant: the real Workspace account wins over CAPWATCH');
{
  const m = load({
    config: SENIORS,
    commanders: [commanderRow('070', '600001', 'Alvarez', 'Rosa', 'Lt Col')],
    contacts: [contactRow('600001', 'rosa.personal@example.com')],
    // Their real account is NOT first.last — a rename or a duplicate. This is
    // exactly the case derivation alone cannot see.
    accounts: [{ capid: '600001', email: 'r.alvarez2@cawgcap.org' }]
  });

  const cc = m.getCommanderInfo('070');
  check('uses the directory account', cc.email, 'r.alvarez2@cawgcap.org');
  check('not the derived address', cc.email !== 'rosa.alvarez@cawgcap.org', true);
  check('not the CAPWATCH personal address', cc.email !== 'rosa.personal@example.com', true);
  check('name still comes from Commanders.txt', [cc.rank, cc.firstName, cc.lastName], ['Lt Col', 'Rosa', 'Alvarez']);
  check('capid carried', cc.capid, '600001');
}

// ---------------------------------------------------------------------------
section('2. Not in the directory: derive the org address, still not CAPWATCH');
{
  const m = load({
    config: SENIORS,
    commanders: [commanderRow('070', '600001', 'Alvarez', 'Rosa')],
    contacts: [contactRow('600001', 'rosa.personal@example.com')],
    accounts: [] // no account readable for this CAPID
  });

  check('derived first.last on the command domain',
    m.getCommanderInfo('070').email, 'rosa.alvarez@cawgcap.org');
}

// ---------------------------------------------------------------------------
section('3. Cadets tenant: command staff live on the senior domain');
{
  const m = load({
    config: CADETS,
    commanders: [commanderRow('071', '600002', 'Nakamura', 'Ken')],
    contacts: [contactRow('600002', 'ken.personal@example.com')],
    // A cadet-tenant account listing can never contain the senior commander, and
    // must not be mistaken for one even if a CAPID happens to collide.
    accounts: [{ capid: '600002', email: 'ken.nakamura@cawgcadets.org' }]
  });

  check('derives onto the SENIOR domain, ignoring the cadet directory',
    m.getCommanderInfo('071').email, 'ken.nakamura@cawgcap.org');
}

// ---------------------------------------------------------------------------
section('4. CAPWATCH personal address is the last resort, not the first');
{
  const m = load({
    config: SENIORS,
    // No usable name, so no address can be derived.
    commanders: [commanderRow('072', '600003', '', '')],
    contacts: [contactRow('600003', 'lee.personal@example.com')],
    accounts: []
  });

  check('falls back to CAPWATCH primary',
    m.getCommanderInfo('072').email, 'lee.personal@example.com');
}

// ---------------------------------------------------------------------------
section('5. No route at all yields null, and the send site drops the CC');
{
  const m = load({
    config: SENIORS,
    commanders: [commanderRow('073', '600004', '', '')],
    contacts: [],
    accounts: []
  });

  const cc = m.getCommanderInfo('073');
  check('email is null', cc.email, null);
  check('record still returned', cc.capid, '600004');
}

// ---------------------------------------------------------------------------
section('6. Unknown org warns and returns null');
{
  const m = load({ commanders: [commanderRow('070', '600001', 'Alvarez', 'Rosa')] });
  check('returns null', m.getCommanderInfo('999'), null);
  check('warned', m.logCalls.warn.some(w => /Commander not found/.test(w.msg)), true);
}

// ---------------------------------------------------------------------------
section('7. Built once per run, not once per member');
{
  const m = load({
    commanders: [commanderRow('070', '600001', 'Alvarez', 'Rosa')],
    contacts: [contactRow('600001', 'rosa.personal@example.com')],
    accounts: [{ capid: '600001', email: 'rosa.alvarez@cawgcap.org' }]
  });

  for (let i = 0; i < 25; i++) m.getCommanderInfo('070');

  check('one directory listing', m.counts.getActiveUsers, 1);
  check('Commanders.txt read once', m.counts.parseFile.Commanders, 1);
  check('MbrContact.txt read once', m.counts.parseFile.MbrContact, 1);
}

// ---------------------------------------------------------------------------
section('8. First row per org wins, as before');
{
  const m = load({
    commanders: [
      commanderRow('070', '600001', 'Alvarez', 'Rosa'),
      commanderRow('070', '600009', 'Later', 'Duplicate')
    ],
    accounts: []
  });

  check('first row', m.getCommanderInfo('070').capid, '600001');
}

// ---------------------------------------------------------------------------
section('9. An unreadable directory degrades instead of failing the run');
{
  const m = load({
    commanders: [commanderRow('070', '600001', 'Alvarez', 'Rosa')],
    contacts: [contactRow('600001', 'rosa.personal@example.com')],
    directoryThrows: true
  });

  check('still resolves, by derivation', m.getCommanderInfo('070').email, 'rosa.alvarez@cawgcap.org');
  check('warned about the directory',
    m.logCalls.warn.some(w => /Directory unreadable/.test(w.msg)), true);
}

// ---------------------------------------------------------------------------
section('10. Templates carry no wing of their own');
{
  const member = { rank: 'C/CMSgt', lastName: 'Okonkwo', expiration: '8/31/2026' };
  const templates = ['Turning18Email', 'Turning21Email', 'ExpiringEmail'];

  const ca = load({ config: SENIORS, directorName: 'Maj Dana Reyes' });
  const hi = load({ config: HIWG, directorName: '' });

  templates.forEach(t => {
    const out = ca.retentionRenderTemplate_(t, member);
    check(t + ': every placeholder resolved', (out.match(/{{\w+}}/g) || []), []);
    check(t + ': member name rendered', /Okonkwo/.test(out), true);

    const other = hi.retentionRenderTemplate_(t, member);
    check(t + ': no CAWG literal on a Hawaii tenant', /CAWG|California/.test(other), false);
    check(t + ': renders the tenant wing', /Hawaii Wing/.test(other), true);
    check(t + ': no unfilled form placeholder', /LINK TO FORM/.test(other), false);
  });

  check('expiration only where the template asks for it',
    /8\/31\/2026/.test(ca.retentionRenderTemplate_('ExpiringEmail', member)), true);
}

// ---------------------------------------------------------------------------
section('11. Signature: named role holder, or the office alone');
{
  const named = load({ config: SENIORS, directorName: 'Maj Dana Reyes' }).retentionSignatureHtml_();
  check('name present', /Maj Dana Reyes/.test(named), true);
  check('office present', /Director of Recruiting/.test(named), true);
  check('wing present', /California Wing Civil Air Patrol/.test(named), true);

  const blank = load({ config: HIWG, directorName: '' }).retentionSignatureHtml_();
  check('office stands alone', blank, '<strong>Director of Recruiting</strong><br>Hawaii Wing Civil Air Patrol');

  const padded = load({ config: SENIORS, directorName: '   ' }).retentionSignatureHtml_();
  check('whitespace-only name treated as blank', /Director of Recruiting<\/strong>/.test(padded), true);
}

// ---------------------------------------------------------------------------
section('12. A value containing $-patterns cannot corrupt the output');
{
  const m = load({ config: SENIORS });
  const out = m.retentionRenderTemplate_('ExpiringEmail', {
    rank: 'Capt',
    lastName: "O'$&Brien$`",
    expiration: '8/31/2026'
  });

  check('literal name survives substitution', /O'\$&Brien\$`/.test(out), true);
}

// ---------------------------------------------------------------------------
section('13. Age-milestone CC adds the DCC; renewal CC adds the recruiting officer');
{
  const m = load({
    config: SENIORS,
    commanders: [commanderRow('070', '600001', 'Alvarez', 'Rosa')],
    members: [
      memberRow('600005', 'Bright', 'Dana'),
      memberRow('600006', 'Cole', 'Sam')
    ],
    duties: [
      dutyRow('600005', 'Deputy Commander for Cadets', '070'),
      dutyRow('600006', 'Recruiting Officer', '070')
    ],
    accounts: []
  });

  check('turning 18/21 goes to commander + DCC',
    m.retentionCcList_('070', CC_DUTY_TITLES.AGE_MILESTONE),
    'rosa.alvarez@cawgcap.org,dana.bright@cawgcap.org');

  check('renewal goes to commander + recruiting officer',
    m.retentionCcList_('070', CC_DUTY_TITLES.RENEWAL),
    'rosa.alvarez@cawgcap.org,sam.cole@cawgcap.org');
}

// ---------------------------------------------------------------------------
section('14. "if one is assigned" — an unfilled duty just leaves the commander');
{
  const m = load({
    config: SENIORS,
    commanders: [commanderRow('070', '600001', 'Alvarez', 'Rosa')],
    members: [memberRow('600005', 'Bright', 'Dana')],
    duties: [dutyRow('600005', 'Deputy Commander for Cadets', '070')],
    accounts: []
  });

  check('no recruiting officer at this unit', m.retentionCcList_('070', CC_DUTY_TITLES.RENEWAL),
    'rosa.alvarez@cawgcap.org');
  check('a duty held at ANOTHER unit does not leak in',
    m.retentionCcList_('071', CC_DUTY_TITLES.AGE_MILESTONE), '');
}

// ---------------------------------------------------------------------------
section('15. The extra CC rides on the commander CC, never replaces it');
{
  // No commander resolvable for this org, but a DCC exists. The requirement is
  // that these staff join the commander — so with no commander there is no CC,
  // rather than the mail quietly going to a different person.
  const m = load({
    config: SENIORS,
    commanders: [],
    members: [memberRow('600005', 'Bright', 'Dana')],
    duties: [dutyRow('600005', 'Deputy Commander for Cadets', '070')],
    accounts: []
  });

  check('no commander means no CC at all',
    m.retentionCcList_('070', CC_DUTY_TITLES.AGE_MILESTONE), '');
}

// ---------------------------------------------------------------------------
section('16. Primary beats assistant, whatever order the rows arrive in');
{
  const assistantFirst = load({
    config: SENIORS,
    commanders: [commanderRow('070', '600001', 'Alvarez', 'Rosa')],
    members: [memberRow('600007', 'Asst', 'Andy'), memberRow('600006', 'Cole', 'Sam')],
    duties: [
      dutyRow('600007', 'Recruiting Officer', '070', true),
      dutyRow('600006', 'Recruiting Officer', '070', false)
    ],
    accounts: []
  });
  check('primary wins when listed second',
    assistantFirst.retentionCcList_('070', CC_DUTY_TITLES.RENEWAL),
    'rosa.alvarez@cawgcap.org,sam.cole@cawgcap.org');

  const primaryFirst = load({
    config: SENIORS,
    commanders: [commanderRow('070', '600001', 'Alvarez', 'Rosa')],
    members: [memberRow('600006', 'Cole', 'Sam'), memberRow('600007', 'Asst', 'Andy')],
    duties: [
      dutyRow('600006', 'Recruiting Officer', '070', false),
      dutyRow('600007', 'Recruiting Officer', '070', true)
    ],
    accounts: []
  });
  check('primary still wins when listed first',
    primaryFirst.retentionCcList_('070', CC_DUTY_TITLES.RENEWAL),
    'rosa.alvarez@cawgcap.org,sam.cole@cawgcap.org');

  const assistantOnly = load({
    config: SENIORS,
    commanders: [commanderRow('070', '600001', 'Alvarez', 'Rosa')],
    members: [memberRow('600007', 'Asst', 'Andy')],
    duties: [dutyRow('600007', 'Recruiting Officer', '070', true)],
    accounts: []
  });
  check('an assistant is used when nobody holds it primary',
    assistantOnly.retentionCcList_('070', CC_DUTY_TITLES.RENEWAL),
    'rosa.alvarez@cawgcap.org,andy.asst@cawgcap.org');
}

// ---------------------------------------------------------------------------
section('17. One person wearing two hats appears once');
{
  const m = load({
    config: SENIORS,
    commanders: [commanderRow('070', '600001', 'Alvarez', 'Rosa')],
    members: [memberRow('600001', 'Alvarez', 'Rosa')],
    // The commander is also the unit's recruiting officer — common in a small unit.
    duties: [dutyRow('600001', 'Recruiting Officer', '070')],
    accounts: []
  });

  check('deduplicated', m.retentionCcList_('070', CC_DUTY_TITLES.RENEWAL),
    'rosa.alvarez@cawgcap.org');
}

// ---------------------------------------------------------------------------
section('18. The legacy CAPR 30-1 duty title still matches');
{
  const m = load({
    config: SENIORS,
    commanders: [commanderRow('070', '600001', 'Alvarez', 'Rosa')],
    members: [memberRow('600006', 'Cole', 'Sam')],
    // Rows predating the ICL rename still read "Recruiting & Retention Officer".
    duties: [dutyRow('600006', 'Recruiting & Retention Officer', '070')],
    accounts: []
  });

  check('normalized to Recruiting Officer and matched',
    m.retentionCcList_('070', CC_DUTY_TITLES.RENEWAL),
    'rosa.alvarez@cawgcap.org,sam.cole@cawgcap.org');
}

// ---------------------------------------------------------------------------
section('19. Duty holders resolve through the same address chain as commanders');
{
  const m = load({
    config: SENIORS,
    commanders: [commanderRow('070', '600001', 'Alvarez', 'Rosa')],
    members: [memberRow('600006', 'Cole', 'Sam')],
    duties: [dutyRow('600006', 'Recruiting Officer', '070')],
    contacts: [contactRow('600006', 'sam.personal@example.com')],
    // Real account is not first.last, so only the directory can find it.
    accounts: [{ capid: '600006', email: 's.cole2@cawgcap.org' }]
  });

  check('directory account preferred for a duty holder too',
    m.retentionCcList_('070', CC_DUTY_TITLES.RENEWAL),
    'rosa.alvarez@cawgcap.org,s.cole2@cawgcap.org');
}

// ---------------------------------------------------------------------------
section('20. DutyPosition.txt is only read when some email type wants staff');
{
  const m = load({
    config: SENIORS,
    commanders: [commanderRow('070', '600001', 'Alvarez', 'Rosa')],
    ccDutyTitles: { AGE_MILESTONE: [], RENEWAL: [] },
    accounts: []
  });

  m.retentionCcList_('070', []);
  check('no duty file read', m.counts.parseFile.DutyPosition, undefined);
  check('no Member.txt name pass either', m.counts.parseFile.Member, undefined);
  check('commander still resolved', m.retentionCcList_('070', []), 'rosa.alvarez@cawgcap.org');
}

done();
