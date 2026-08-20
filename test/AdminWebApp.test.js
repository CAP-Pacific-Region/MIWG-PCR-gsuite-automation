/**
 * admin-webapp/ — the domain admin help-desk app.
 *
 * WHAT THIS FILE IS FOR
 *
 * The app is a SECOND implementation of things src/ already does: the welcome
 * email and the policy governing when it may be re-sent, the CAPID account
 * ranking, and the format of the welcome-email audit ledger. It carries copies
 * because the two live in different Apps Script projects, and a project cannot
 * call another's functions without becoming a library with its own deploy-version
 * step.
 *
 * A copy that drifts is worse than no copy at all. Each of these has a specific
 * failure mode if it does:
 *
 *   - the eligibility policy: the app would mail credentials into the mailbox
 *     they unlock, or reset a password on an account someone is actively using —
 *     the two things src/ was carefully built to refuse.
 *   - the welcome template: a member's credentials would look different
 *     depending on which route created their account.
 *   - the account ranking: a help desk would reset the password on a member's
 *     dead twin account while they keep failing to sign into the live one.
 *   - the ledger version: src/'s monthly audit REFUSES TO RUN against a ledger
 *     whose version it does not recognise.
 *
 * So the first section runs both copies over the same inputs and requires
 * identical answers. The rest pins the decisions this app makes on its own —
 * chiefly the two gates that make it safe to hand to a help-desk volunteer: only
 * a super admin may act on an admin, and only a configured group may be changed.
 *
 * All members, addresses and CAPIDs here are synthetic.
 *
 * Run: npm test
 */
const fs = require('fs');
const path = require('path');
const { loadModule, makeLogger, makeChecker, Session, Utilities } = require('./helpers/apps-script');
const { memberFile } = require('./helpers/capwatch-fixtures');

const { section, check, done } = makeChecker();

const ROOT = path.join(__dirname, '..');
const APP = path.join(ROOT, 'admin-webapp');
const SRC_RESEND = path.join(ROOT, 'src', 'accounts-and-groups', 'WelcomeEmailResend.gs');
const SRC_GUARD = path.join(ROOT, 'src', 'accounts-and-groups', 'DuplicateAccountGuard.gs');
const SRC_AUDIT = path.join(ROOT, 'src', 'accounts-and-groups', 'WelcomeEmailAudit.gs');

const TENANT = {
  DOMAIN: 'cawgcap.org',
  EMAIL_DOMAIN: '@cawgcap.org',
  SECONDARY_EMAIL_DOMAIN: '@cawg.cap.gov',
  WING: 'CA',
  ORG_LABEL: 'CAWG',
  SUPPORT_EMAIL: 'it@cawgcap.org',
  CAPWATCH_DATA_FOLDER_ID: 'folder-id',
  TWO_SV_GROUP: 'ca.2sv-setup@cawgcap.org',
  MANAGED_GROUPS: ['ca.newsletter@cawgcap.org'],
  ALLOWED_ROLES: ['_HELP_DESK_ADMIN_ROLE'],
  ADMIN_GROUP: '',
  AUDIT_SPREADSHEET_ID: '',
  AUDIT_SHEET_NAME: 'Admin Web App Log'
};

// ============================================================================
// PARITY WITH src/
// ============================================================================

section('Welcome-email eligibility is the same policy as src/WelcomeEmailResend.gs');

const srcResend = loadModule(SRC_RESEND, {
  Logger: makeLogger().logger, AdminDirectory: {}, CONFIG: {}, console: console
}, ['welcomeResendEligibility_']);

const appActions = loadModule(path.join(APP, 'Actions.gs'), {
  Logger: makeLogger().logger,
  ADMIN_CONFIG: TENANT,
  AdminDirectory: {},
  MailApp: { sendEmail: () => {} },
  admManagedGroups_: () => [TENANT.TWO_SV_GROUP].concat(TENANT.MANAGED_GROUPS),
  console: console
}, ['admWelcomeEligibility_', 'admTenantDomains_', 'admOffTenantRecipients_', 'admGroupFromRequest_']);

const TENANT_DOMAINS = { tenantDomains: ['cawgcap.org', '@cawgcap.org', '@cawg.cap.gov'] };

/**
 * The cases below are the whole decision space of the policy: every refusal
 * reason, both force paths, and the recipient de-duplication. Both copies must
 * agree on all of them, verdict, reason slug, sentence and recipients alike.
 */
const CASES = [
  ['eligible, personal primary', {
    member: { firstName: 'Dana', lastName: 'Okonkwo', email: 'dana@example.com', secondaryEmail: null },
    account: { email: 'dana.okonkwo@cawgcap.org', suspended: false, archived: false, neverSignedIn: true }
  }],
  ['no CAPWATCH record', { member: null, account: { email: 'x@cawgcap.org', neverSignedIn: true } }],
  ['no account', { member: { email: 'dana@example.com' }, account: null }],
  ['archived account', {
    member: { email: 'dana@example.com' },
    account: { email: 'dana@cawgcap.org', archived: true, neverSignedIn: true }
  }],
  ['only address is the account itself', {
    member: { email: 'dana.okonkwo@cawgcap.org', secondaryEmail: null },
    account: { email: 'dana.okonkwo@cawgcap.org', neverSignedIn: true }
  }],
  ['only address is on the secondary domain', {
    member: { email: 'dana.okonkwo@cawg.cap.gov', secondaryEmail: null },
    account: { email: 'dana.okonkwo@cawgcap.org', neverSignedIn: true }
  }],
  ['suspended, not forced', {
    member: { email: 'dana@example.com' },
    account: { email: 'dana@cawgcap.org', suspended: true, neverSignedIn: true }
  }],
  ['suspended, forced', {
    member: { email: 'dana@example.com' },
    account: { email: 'dana@cawgcap.org', suspended: true, neverSignedIn: true },
    opts: { force: true }
  }],
  ['has login history, not forced', {
    member: { email: 'dana@example.com' },
    account: { email: 'dana@cawgcap.org', neverSignedIn: false }
  }],
  ['has login history, forced', {
    member: { email: 'dana@example.com' },
    account: { email: 'dana@cawgcap.org', neverSignedIn: false },
    opts: { force: true }
  }],
  ['both addresses, one duplicated', {
    member: { email: 'dana@example.com', secondaryEmail: 'dana@example.com' },
    account: { email: 'dana@cawgcap.org', neverSignedIn: true }
  }],
  ['both addresses, distinct', {
    member: { email: 'dana@example.com', secondaryEmail: 'dana.o@example.net' },
    account: { email: 'dana@cawgcap.org', neverSignedIn: true }
  }],
  ['force does NOT override a missing recipient', {
    member: { email: 'dana.okonkwo@cawgcap.org' },
    account: { email: 'dana.okonkwo@cawgcap.org', neverSignedIn: false },
    opts: { force: true }
  }]
];

/**
 * The DECISION is compared, not the prose. Two of the refusal sentences are
 * deliberately reworded for this app — src/ tells an operator in the Apps Script
 * editor to "pass {force: true}", which is meaningless to someone looking at a
 * web page with a checkbox on it. Everything that governs behaviour (the verdict,
 * the reason slug, the recipient list) must still match exactly, and the two
 * reworded sentences are pinned by name below so a THIRD divergence cannot slip
 * in unnoticed.
 */
const decision = (v) => ({ ok: v.ok, reason: v.reason, recipients: v.recipients });

CASES.forEach(([name, c]) => {
  const opts = Object.assign({}, TENANT_DOMAINS, c.opts || {});
  check(name,
    decision(appActions.admWelcomeEligibility_(c.member, c.account, opts)),
    decision(srcResend.welcomeResendEligibility_(c.member, c.account, opts)));
});

const REWORDED = ['suspended', 'already-signed-in'];
CASES.forEach(([name, c]) => {
  const opts = Object.assign({}, TENANT_DOMAINS, c.opts || {});
  const app = appActions.admWelcomeEligibility_(c.member, c.account, opts);
  const src = srcResend.welcomeResendEligibility_(c.member, c.account, opts);
  if (REWORDED.indexOf(app.reason) !== -1) return;
  check(name + ' — same sentence', app.detail, src.detail);
});

check('the suspended refusal names the checkbox, not a function argument',
  appActions.admWelcomeEligibility_(
    { email: 'dana@example.com' },
    { email: 'dana@cawgcap.org', suspended: true, neverSignedIn: true },
    TENANT_DOMAINS).detail,
  'The account is suspended — credentials would not work. Resolve the ' +
  'suspension first, or tick "force" if you are about to lift it.');

check('so does the login-history refusal',
  appActions.admWelcomeEligibility_(
    { email: 'dana@example.com' },
    { email: 'dana@cawgcap.org', neverSignedIn: false },
    TENANT_DOMAINS).detail,
  'The member has signed into this account, so they already have a working ' +
  'password. A resend RESETS it and would lock them out. Tick "force" only ' +
  'if they have genuinely lost access.');

check('the app derives the same tenant domain list',
  appActions.admTenantDomains_(),
  ['cawgcap.org', '@cawgcap.org', '@cawg.cap.gov']);

section('The welcome email template is byte-identical to the one src/ sends');

check('WelcomeEmail.html matches src/recruiting-and-retention/WelcomeEmail.html',
  fs.readFileSync(path.join(APP, 'WelcomeEmail.html'), 'utf8'),
  fs.readFileSync(path.join(ROOT, 'src', 'recruiting-and-retention', 'WelcomeEmail.html'), 'utf8'));

section('2SV instructions ride along with credentials, but only when needed');

/**
 * Credentials mail reaches a member at the one moment they can act on this —
 * they are about to sign in. But telling someone who already has 2SV on to go
 * and turn it on reads as a system that does not know what it is talking about,
 * and the next real instruction from IT gets skimmed. So it is conditional, and
 * the condition is the directory's own flag rather than an assumption about
 * whether the account is new.
 */
const mailSent = [];
const appCreds = loadModule(path.join(APP, 'Credentials.gs'), {
  Logger: makeLogger().logger,
  ADMIN_CONFIG: TENANT,
  AdminDirectory: {},
  DriveApp: {},
  HtmlService: {
    createTemplateFromFile: () => ({
      getRawContent: () => fs.readFileSync(path.join(APP, 'WelcomeEmail.html'), 'utf8')
    })
  },
  MailApp: { sendEmail: (options) => mailSent.push(options) },
  admEscape_: (v) => String(v == null ? '' : v).replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])),
  Utilities: Utilities, Session: Session, console: console
}, ['admSendWelcomeEmail_', 'adm2SvInstructionsHtml_', 'adm2SvInsert_']);

const MEMBER = { firstName: 'Dana', lastName: 'Okonkwo', rank: 'Maj', email: 'dana@example.com',
  secondaryEmail: null, capsn: '100001' };

mailSent.length = 0;
appCreds.admSendWelcomeEmail_(MEMBER, 'dana.okonkwo@cawgcap.org', 'Pw!23', ['dana@example.com'], true);
const withBlock = mailSent[0].htmlBody;

mailSent.length = 0;
appCreds.admSendWelcomeEmail_(MEMBER, 'dana.okonkwo@cawgcap.org', 'Pw!23', ['dana@example.com'], false);
const withoutBlock = mailSent[0].htmlBody;

check('an unenrolled account gets the setup steps',
  withBlock.indexOf('turn on 2-Step Verification') !== -1, true);
check('an enrolled account does not',
  withoutBlock.indexOf('turn on 2-Step Verification') !== -1, false);

// The template's own one-line mention is untouched in both — this adds the how,
// it does not replace the whether.
check('the template\'s existing 2SV link survives either way',
  [withBlock, withoutBlock].map(h => h.indexOf('signinoptions/two-step-verification') !== -1),
  [true, true]);

// Placed above the footer bar so it reads as part of the message rather than
// something bolted on after the sign-off.
check('the block lands above the footer, not after the sign-off',
  withBlock.indexOf('turn on 2-Step Verification') < withBlock.indexOf('<div class="footer">'),
  true);
check('the credentials still appear before it',
  withBlock.indexOf('Temporary Password') < withBlock.indexOf('turn on 2-Step Verification'),
  true);

// A missing anchor must degrade to "further down the email", never to a failed
// send: the password is already live by the time this runs.
check('a template with no footer marker still gets the block',
  appCreds.adm2SvInsert_('<html><body><p>hi</p></body></html>')
    .indexOf('turn on 2-Step Verification') !== -1, true);
check('and one with neither marker still gets it',
  appCreds.adm2SvInsert_('bare string')
    .indexOf('turn on 2-Step Verification') !== -1, true);

check('the support address is named when configured',
  appCreds.adm2SvInstructionsHtml_().indexOf('it@cawgcap.org') !== -1, true);

section('The audit ledger format matches, or src/ refuses to run against it');

const srcAudit = loadModule(SRC_AUDIT, {
  Logger: makeLogger().logger, DriveApp: {}, CONFIG: {},
  Utilities: Utilities, Session: Session, console: console
}, ['WELCOME_AUDIT_CONFIG']);

const appWelcome = loadModule(path.join(APP, 'Credentials.gs'), {
  Logger: makeLogger().logger, ADMIN_CONFIG: TENANT, AdminDirectory: {}, DriveApp: {},
  HtmlService: {}, MailApp: {}, Utilities: Utilities, Session: Session, console: console
}, ['ADM_WELCOME_LEDGER_VERSION', 'ADM_WELCOME_LEDGER_FILE_NAME']);

check('ledger version', appWelcome.ADM_WELCOME_LEDGER_VERSION, srcAudit.WELCOME_AUDIT_CONFIG.LEDGER_VERSION);
check('ledger file name', appWelcome.ADM_WELCOME_LEDGER_FILE_NAME, srcAudit.WELCOME_AUDIT_CONFIG.LEDGER_FILE_NAME);

section('The authoritative-account ranking matches DuplicateAccountGuard.gs');

const srcGuard = loadModule(SRC_GUARD, {
  Logger: makeLogger().logger, AdminDirectory: {}, CONFIG: {}, console: console
}, ['chooseAuthoritativeAccount_']);

const appAccounts = loadModule(path.join(APP, 'Accounts.gs'), {
  Logger: makeLogger().logger, ADMIN_CONFIG: TENANT, ADM_CAPID_RE: /^\d{5,7}$/,
  AdminDirectory: {}, admDirectoryUser_: () => null, admBuildMemberRecord_: () => null,
  admSearchMembersByName_: () => [], ADM_SEARCH_LIMIT: 25, console: console
}, ['admChooseAuthoritativeAccount_', 'admManagedGroups_', 'admCanonicalLocalpart_']);

/**
 * The pair that taught src/ this lesson: the account a member actually uses is
 * the older, oddly-named one, and the canonically-named twin has never been
 * signed into. A ranking that preferred the canonical name would send a help
 * desk to reset the wrong account.
 */
const ACCOUNT_SETS = [
  ['login history beats a canonical name', [
    { email: 'dana.okonkwo@cawgcap.org', suspended: false, created: '2026-01-01T00:00:00Z', lastLogin: null, neverSignedIn: true },
    { email: 'dana.okonkwo.2@cawgcap.org', suspended: false, created: '2020-01-01T00:00:00Z', lastLogin: '2026-08-01T00:00:00Z', neverSignedIn: false }
  ]],
  ['most recent login wins when both have history', [
    { email: 'a@cawgcap.org', suspended: false, created: '2020-01-01T00:00:00Z', lastLogin: '2026-05-01T00:00:00Z', neverSignedIn: false },
    { email: 'b@cawgcap.org', suspended: false, created: '2024-01-01T00:00:00Z', lastLogin: '2026-08-01T00:00:00Z', neverSignedIn: false }
  ]],
  ['active beats suspended', [
    { email: 'a@cawgcap.org', suspended: true, created: '2024-01-01T00:00:00Z', lastLogin: null, neverSignedIn: true },
    { email: 'b@cawgcap.org', suspended: false, created: '2020-01-01T00:00:00Z', lastLogin: null, neverSignedIn: true }
  ]],
  ['the configured domain beats a legacy one', [
    { email: 'dana.okonkwo@legacy.example.org', suspended: false, created: '2024-01-01T00:00:00Z', lastLogin: null, neverSignedIn: true },
    { email: 'dana.okonkwo@cawgcap.org', suspended: false, created: '2020-01-01T00:00:00Z', lastLogin: null, neverSignedIn: true }
  ]],
  ['single account', [
    { email: 'solo@cawgcap.org', suspended: false, created: '2024-01-01T00:00:00Z', lastLogin: null, neverSignedIn: true }
  ]]
];

ACCOUNT_SETS.forEach(([name, accounts]) => {
  check(name,
    appAccounts.admChooseAuthoritativeAccount_(accounts, 'dana.okonkwo', '@cawgcap.org'),
    srcGuard.chooseAuthoritativeAccount_(accounts, 'dana.okonkwo', '@cawgcap.org'));
});

check('no accounts resolves to null',
  appAccounts.admChooseAuthoritativeAccount_([], 'dana.okonkwo', '@cawgcap.org'),
  srcGuard.chooseAuthoritativeAccount_([], 'dana.okonkwo', '@cawgcap.org'));

check('a collision suffix is stripped the same way',
  appAccounts.admCanonicalLocalpart_('dana.okonkwo.2@cawgcap.org'), 'dana.okonkwo');

// ============================================================================
// THE APP'S OWN DECISIONS
// ============================================================================

section('Only a configured group may be changed');

check('the 2SV group is managed',
  appActions.admGroupFromRequest_('ca.2sv-setup@cawgcap.org'), 'ca.2sv-setup@cawgcap.org');
check('case and whitespace do not matter',
  appActions.admGroupFromRequest_('  CA.2SV-Setup@cawgcap.org '), 'ca.2sv-setup@cawgcap.org');
check('an extra managed group is accepted',
  appActions.admGroupFromRequest_('ca.newsletter@cawgcap.org'), 'ca.newsletter@cawgcap.org');

const refuses = (group) => {
  try { appActions.admGroupFromRequest_(group); return 'accepted'; }
  catch (e) { return 'refused'; }
};
check('an arbitrary group is refused', refuses('ca.wing-admins@cawgcap.org'), 'refused');
check('a blank group is refused', refuses(''), 'refused');
check('a group that merely CONTAINS a managed address is refused',
  refuses('ca.2sv-setup@cawgcap.org.example.com'), 'refused');

const noGroups = loadModule(path.join(APP, 'Actions.gs'), {
  Logger: makeLogger().logger, ADMIN_CONFIG: TENANT, AdminDirectory: {}, MailApp: {},
  admManagedGroups_: () => [], console: console
}, ['admGroupFromRequest_']);
let unconfigured = 'accepted';
try { noGroups.admGroupFromRequest_('anything@cawgcap.org'); } catch (e) { unconfigured = 'refused'; }
check('with nothing configured, every group is refused', unconfigured, 'refused');

check('the managed list is the 2SV group plus the extras, de-duplicated',
  appAccounts.admManagedGroups_(), ['ca.2sv-setup@cawgcap.org', 'ca.newsletter@cawgcap.org']);

section('Credentials are never mailed to an address on our own domains');

check('a personal address is a valid recipient',
  appActions.admOffTenantRecipients_(
    { email: 'dana@example.com', secondaryEmail: null }, 'dana.okonkwo@cawgcap.org'),
  ['dana@example.com']);
check('the account being reset is excluded',
  appActions.admOffTenantRecipients_(
    { email: 'dana.okonkwo@cawgcap.org', secondaryEmail: null }, 'dana.okonkwo@cawgcap.org'),
  []);
check('another address on the primary domain is excluded',
  appActions.admOffTenantRecipients_(
    { email: 'someone.else@cawgcap.org', secondaryEmail: null }, 'dana.okonkwo@cawgcap.org'),
  []);
check('the secondary domain is excluded',
  appActions.admOffTenantRecipients_(
    { email: 'dana@cawg.cap.gov', secondaryEmail: 'dana@example.net' }, 'dana.okonkwo@cawgcap.org'),
  ['dana@example.net']);
check('no member record means no recipients',
  appActions.admOffTenantRecipients_(null, 'dana.okonkwo@cawgcap.org'), []);

section('A lookalike domain is not one of ours');

const appConfig = loadModule(path.join(APP, 'Config.gs'), {
  PropertiesService: {
    getScriptProperties: () => ({
      getProperty: (key) => ({
        TENANT_EMAIL_DOMAIN: '@cawgcap.org',
        TENANT_SECONDARY_EMAIL_DOMAIN: '@cawg.cap.gov',
        TENANT_DOMAIN: 'cawgcap.org',
        TENANT_WING: 'CA'
      }[key] || null)
    })
  },
  console: console
}, ['admIsTenantAddress_', 'admMissingConfig_', 'admEscape_']);

check('our own domain', appConfig.admIsTenantAddress_('dana@cawgcap.org'), true);
check('our secondary domain', appConfig.admIsTenantAddress_('dana@cawg.cap.gov'), true);
check('a suffix lookalike', appConfig.admIsTenantAddress_('dana@cawgcap.org.example.com'), false);
check('a prefix lookalike', appConfig.admIsTenantAddress_('dana@notcawgcap.org'), false);
check('a personal address', appConfig.admIsTenantAddress_('dana@example.com'), false);
check('not an address at all', appConfig.admIsTenantAddress_('dana'), false);
check('a missing CAPWATCH folder is a blocking config gap',
  appConfig.admMissingConfig_(), ['TENANT_CAPWATCH_DATA_FOLDER_ID']);
check('markup in a member name cannot reach the page as markup',
  appConfig.admEscape_('<script>x</script>'), '&lt;script&gt;x&lt;/script&gt;');

section('A member on the other tenant is flagged, never silently unusable');

/**
 * CAPWATCH is scoped to the WING, not to the tenant: a seniors-tenant extract
 * lists every cadet in the wing. They are findable here and actionable nowhere
 * on this page, and a help desk that is not told so concludes either "this
 * member does not exist" or "their account is missing" — the second of which
 * ends in someone creating an account on the wrong Workspace.
 */
function loadConfigFor(profile, extras) {
  const props = Object.assign({
    TENANT_EMAIL_DOMAIN: '@cawgcap.org',
    TENANT_DOMAIN: 'cawgcap.org',
    TENANT_WING: 'CA',
    TENANT_CAPWATCH_DATA_FOLDER_ID: 'folder-id',
    TENANT_PROFILE: profile
  }, extras || {});
  return loadModule(path.join(APP, 'Config.gs'), {
    PropertiesService: {
      getScriptProperties: () => ({ getProperty: (key) => props[key] || null })
    },
    console: console
  }, ['admTenantProvisionsType_', 'admElsewhereSentence_', 'ADM_PROFILE_MEMBER_TYPES']);
}

/**
 * The managed-group list as the running app would compute it: Config.gs reads
 * the properties, Accounts.gs turns them into the allow-list. Loaded together
 * because the fallback between the two property names lives in the first and
 * only shows up in the second.
 */
function loadGroupsFor(props) {
  const all = Object.assign({ TENANT_PROFILE: 'seniors', TENANT_WING: 'CA' }, props);
  const cfg = loadModule(path.join(APP, 'Config.gs'), {
    PropertiesService: {
      getScriptProperties: () => ({ getProperty: (key) => all[key] || null })
    },
    console: console
  }, ['ADMIN_CONFIG']);

  return loadModule(path.join(APP, 'Accounts.gs'), {
    Logger: makeLogger().logger, ADMIN_CONFIG: cfg.ADMIN_CONFIG, ADM_CAPID_RE: /^\d{5,7}$/,
    AdminDirectory: {}, admDirectoryUser_: () => null, admBuildMemberRecord_: () => null,
    admSearchMembersByName_: () => [], ADM_SEARCH_LIMIT: 25, console: console
  }, ['admManagedGroups_']).admManagedGroups_();
}

const seniorsCfg = loadConfigFor('seniors');
check('a senior is ours on the seniors tenant', seniorsCfg.admTenantProvisionsType_('SENIOR'), true);
check('so is a fifty-year member', seniorsCfg.admTenantProvisionsType_('FIFTY YEAR'), true);
check('so is a cadet sponsor', seniorsCfg.admTenantProvisionsType_('CADET SPONSOR'), true);
check('a CADET is NOT', seniorsCfg.admTenantProvisionsType_('CADET'), false);
check('case does not matter', seniorsCfg.admTenantProvisionsType_('cadet'), false);
check('an unknown type reads as not ours', seniorsCfg.admTenantProvisionsType_('PATRON'), false);
check('a blank type reads as not ours', seniorsCfg.admTenantProvisionsType_(''), false);

const cadetsCfg = loadConfigFor('cadets');
check('on the cadets tenant a cadet IS ours', cadetsCfg.admTenantProvisionsType_('CADET'), true);
check('…and a senior is not', cadetsCfg.admTenantProvisionsType_('SENIOR'), false);

// The region tenant provisions cadets as well as seniors, which is why this is a
// table and not "cadets belong to somebody else".
const regionCfg = loadConfigFor('region');
check('on the region tenant both are ours',
  [regionCfg.admTenantProvisionsType_('CADET'), regionCfg.admTenantProvisionsType_('SENIOR')],
  [true, true]);

/**
 * The table is a copy of MEMBER_TYPES_ACTIVE in src/config.gs, which pads its
 * arrays with empty strings for the sheet UI. Compared with those dropped: if a
 * profile's population ever changes there, this must follow or the app will send
 * a help desk to the wrong site.
 */
const srcConfig = fs.readFileSync(path.join(ROOT, 'src', 'config.gs'), 'utf8');
const srcTypes = (srcConfig.match(/MEMBER_TYPES_ACTIVE:\s*\[[^\]]*\]/g) || [])
  .map(line => (line.match(/'([^']*)'/g) || [])
    .map(s => s.replace(/'/g, '')).filter(Boolean));

check('the profile→member-type table matches src/config.gs',
  [seniorsCfg.ADM_PROFILE_MEMBER_TYPES.seniors,
   seniorsCfg.ADM_PROFILE_MEMBER_TYPES.cadets,
   seniorsCfg.ADM_PROFILE_MEMBER_TYPES.region],
  srcTypes);

check('the sentence names the configured peer admin site',
  loadConfigFor('seniors', { WEBAPP_PEER_ADMIN_URL: 'https://cadets.example.org/admin' })
    .admElsewhereSentence_('CADET'),
  'Cadet accounts are not on this Workspace, so nothing on this page can act on them. ' +
  'Use the other tenant\'s admin site instead: https://cadets.example.org/admin');

check('the earlier property name still works where it was set',
  loadConfigFor('seniors', { WEBAPP_CADET_TOOLS_URL: 'https://cadets.example.org/admin' })
    .admElsewhereSentence_('CADET')
    .indexOf('https://cadets.example.org/admin') > 0, true);

check('with no URL it names the peer domain from XT_PEER_DOMAIN',
  loadConfigFor('seniors', { XT_PEER_DOMAIN: 'cawgcadets.org' })
    .admElsewhereSentence_('CADET'),
  'Cadet accounts are not on this Workspace, so nothing on this page can act on them. ' +
  'Their accounts live on cawgcadets.org, which has its own admin page.');

check('with neither configured it is still not a dead end',
  loadConfigFor('seniors').admElsewhereSentence_('CADET'),
  'Cadet accounts are not on this Workspace, so nothing on this page can act on them. ' +
  'Their accounts live on the other tenant\'s Workspace, which has its own admin page.');

/**
 * The direction that matters for the cadet deployment. On the CADETS tenant the
 * members this app cannot help are SENIORS, and the notice must point at the
 * seniors site — telling a cadet-tenant admin "their accounts live on
 * cawgcadets.org" would point them at the page they are already standing on.
 */
check('on the cadets tenant a senior is the one who is elsewhere',
  loadConfigFor('cadets', { XT_PEER_DOMAIN: 'cawgcap.org' }).admElsewhereSentence_('SENIOR'),
  'Senior accounts are not on this Workspace, so nothing on this page can act on them. ' +
  'Their accounts live on cawgcap.org, which has its own admin page.');

check('…and a cadet there is not elsewhere at all',
  loadConfigFor('cadets').admTenantProvisionsType_('CADET'), true);

check('an unrecognised type is named neutrally',
  loadConfigFor('seniors').admElsewhereSentence_('PATRON')
    .indexOf('This member\'s accounts are not on this Workspace'), 0);

section('The 2SV group address is accepted under either property name');

check('WEBAPP_2SV_SETUP_GROUP is read',
  loadGroupsFor({ WEBAPP_2SV_SETUP_GROUP: 'a@x.org' }), ['a@x.org']);
// src/TwoSvSetupGroup.gs prunes the same group under the TENANT_ name; accepting
// it here means copying a tenant's canonical values across does the right thing.
check('TENANT_2SV_SETUP_GROUP is accepted as a fallback',
  loadGroupsFor({ TENANT_2SV_SETUP_GROUP: 'b@x.org' }), ['b@x.org']);
check('the WEBAPP_ name wins when both are set',
  loadGroupsFor({ WEBAPP_2SV_SETUP_GROUP: 'a@x.org', TENANT_2SV_SETUP_GROUP: 'b@x.org' }),
  ['a@x.org']);

section('Who may use the app, and whom they may act on');

/**
 * A directory holding one super admin, one help-desk admin, one admin with an
 * unrelated custom role, and one ordinary member.
 */
const DIRECTORY = {
  'boss@cawgcap.org': { id: '1', primaryEmail: 'boss@cawgcap.org', isAdmin: true },
  'helpdesk@cawgcap.org': { id: '2', primaryEmail: 'helpdesk@cawgcap.org', isDelegatedAdmin: true },
  'reports@cawgcap.org': { id: '3', primaryEmail: 'reports@cawgcap.org', isDelegatedAdmin: true },
  'member@cawgcap.org': { id: '4', primaryEmail: 'member@cawgcap.org' }
};
const ROLES = { '2': ['_HELP_DESK_ADMIN_ROLE'], '3': ['Reporting Viewer'] };

function makeDirectory(opts) {
  const options = opts || {};
  return {
    Users: {
      get: (email) => {
        const u = DIRECTORY[String(email).toLowerCase()];
        if (!u) throw new Error('Resource Not Found: userKey');
        return u;
      }
    },
    RoleAssignments: {
      list: (customer, params) => {
        if (options.rolesThrow) throw new Error('Backend Error');
        const names = ROLES[params.userKey] || [];
        return { items: names.map((n, i) => ({ roleId: params.userKey + ':' + i })) };
      }
    },
    Roles: {
      get: (customer, roleId) => {
        const [userKey, index] = String(roleId).split(':');
        return { roleName: (ROLES[userKey] || [])[Number(index)] };
      }
    },
    Members: {
      hasMember: () => ({ isMember: false })
    }
  };
}

function loadAuth(directory) {
  return loadModule(path.join(APP, 'Auth.gs'), {
    Logger: makeLogger().logger,
    ADMIN_CONFIG: TENANT,
    AdminDirectory: directory,
    Session: { getActiveUser: () => ({ getEmail: () => '' }) },
    admIsTenantAddress_: appConfig.admIsTenantAddress_,
    admMissingConfig_: () => [],
    admSupportSentence_: () => 'Contact IT.',
    console: console
  }, ['admActorPrivileges_', 'admAssertMayActOn_']);
}

const auth = loadAuth(makeDirectory());
const via = (email) => auth.admActorPrivileges_(email).via;

check('a super admin is allowed', via('boss@cawgcap.org'), 'super-admin');
check('a help desk admin is allowed', via('helpdesk@cawgcap.org'), 'role:_HELP_DESK_ADMIN_ROLE');
check('an admin holding an unrelated role is not', via('reports@cawgcap.org'), 'none');
check('an ordinary member is not', via('member@cawgcap.org'), 'none');
check('an unknown account is not', via('ghost@cawgcap.org'), 'none');
check('an off-domain caller is not', via('someone@example.com'), 'none');
check('no identity at all is not', via(''), 'none');

const authBroken = loadAuth(makeDirectory({ rolesThrow: true }));
check('a role lookup that fails denies rather than passes',
  authBroken.admActorPrivileges_('helpdesk@cawgcap.org').allowed, false);
check('…but a super admin does not depend on it',
  authBroken.admActorPrivileges_('boss@cawgcap.org').allowed, true);

const superActor = { email: 'boss@cawgcap.org', superAdmin: true, via: 'super-admin' };
const helpdeskActor = { email: 'helpdesk@cawgcap.org', superAdmin: false, via: 'role:_HELP_DESK_ADMIN_ROLE' };
const mayAct = (actor, account) => {
  try { auth.admAssertMayActOn_(actor, account); return 'allowed'; }
  catch (e) { return 'refused'; }
};

check('a help desk admin may act on a member',
  mayAct(helpdeskActor, DIRECTORY['member@cawgcap.org']), 'allowed');
check('a help desk admin may NOT act on a super admin',
  mayAct(helpdeskActor, DIRECTORY['boss@cawgcap.org']), 'refused');
check('a help desk admin may NOT act on another delegated admin',
  mayAct(helpdeskActor, DIRECTORY['reports@cawgcap.org']), 'refused');
check('a help desk admin may not act on themselves either, being an admin',
  mayAct(helpdeskActor, DIRECTORY['helpdesk@cawgcap.org']), 'refused');
check('a super admin may act on an admin',
  mayAct(superActor, DIRECTORY['helpdesk@cawgcap.org']), 'allowed');
check('an unreadable account is refused, not assumed safe',
  mayAct(superActor, null), 'refused');

section('CAPWATCH contact addresses follow addContactInfo()');

const MBR_CONTACT = [
  'CAPID,Type,Priority,Contact,UsrID,DateMod,DoNotContact',
  '"100001","EMAIL","PRIMARY","dana@example.com","u","1/1/2026","False"',
  '"100001","EMAIL","SECONDARY","dana.o@example.net","u","1/1/2026","False"',
  '"100002","EMAIL","PRIMARY","quiet@example.com","u","1/1/2026","True"',
  '"100002","EMAIL","SECONDARY","quiet.two@example.net","u","1/1/2026","False"',
  '"100003","CELL PHONE","PRIMARY","555-0100","u","1/1/2026","False"',
  ''
].join('\n');

const CAPWATCH_FILES = {
  'MbrContact.txt': MBR_CONTACT,
  'Member.txt': memberFile([
    { capid: '100001', orgid: '100', type: 'SENIOR' },
    { capid: '100004', orgid: '100', type: 'CADET' }
  ])
};

const appMember = loadModule(path.join(APP, 'MemberRecord.gs'), {
  Logger: makeLogger().logger,
  ADMIN_CONFIG: TENANT,
  ADM_CAPID_RE: /^\d{5,7}$/,
  Utilities: Utilities,
  admTenantProvisionsType_: seniorsCfg.admTenantProvisionsType_,
  DriveApp: {
    getFolderById: () => ({
      getFilesByName: (name) => {
        const content = CAPWATCH_FILES[name];
        let consumed = false;
        return {
          hasNext: () => content !== undefined && !consumed,
          next: () => { consumed = true; return { getBlob: () => ({ getDataAsString: () => content }) }; }
        };
      }
    })
  },
  console: console
}, ['admContactEmails_', 'admMemberDisplayName_', 'admSearchMembersByName_']);

check('PRIMARY becomes the contact address, SECONDARY the second',
  appMember.admContactEmails_('100001'),
  { email: 'dana@example.com', secondaryEmail: 'dana.o@example.net',
    primaryEmailValue: 'dana@example.com', primaryEmailDNC: false, secondaryEmailDNC: false });

// The rule src/ applies: a DoNotContact PRIMARY is not used as the contact
// address, but is still reported — many members list a personal address there.
check('a DoNotContact PRIMARY is reported but not used',
  appMember.admContactEmails_('100002'),
  { email: null, secondaryEmail: 'quiet.two@example.net',
    primaryEmailValue: 'quiet@example.com', primaryEmailDNC: true, secondaryEmailDNC: false });

check('a member with no email rows has no addresses',
  appMember.admContactEmails_('100003'),
  { email: null, secondaryEmail: null, primaryEmailValue: null,
    primaryEmailDNC: false, secondaryEmailDNC: false });

// A cadet is RETURNED by the search and flagged, never filtered out: an admin
// who searched a name and got nothing would conclude the member does not exist.
check('a name search returns cadets, flagged as off-tenant',
  appMember.admSearchMembersByName_('last1000', 10).map(c => [c.capid, c.type, c.offTenant]),
  [['100001', 'SENIOR', false], ['100004', 'CADET', true]]);

check('a display name carries grade and suffix',
  appMember.admMemberDisplayName_({ rank: 'Maj', firstName: 'Dana', lastName: 'Okonkwo', suffix: 'Jr' }),
  'Maj Dana Okonkwo Jr');

section('Every client-reachable function is behind requireAdmin_()');

/**
 * The one structural check in this file. AdminApi.gs is the attack surface —
 * anything callable from google.script.run is callable by any signed-in domain
 * user with a console open, button or no button — so a new api* function that
 * forgets its gate is the single most likely way this app grows a hole.
 */
// Line endings normalized: this repo is checked out with core.autocrlf on
// Windows, so the file on disk may use CRLF. Matching a literal newline-brace
// made this check silently test an empty string and report every entry point
// as ungated -- failing for a reason unrelated to what it guards.
const apiSource = fs.readFileSync(path.join(APP, 'AdminApi.gs'), 'utf8').replace(/\r\n/g, '\n');
const apiFunctions = (apiSource.match(/^function (api\w+)/gm) || [])
  .map(line => line.replace('function ', ''));

check('the entry points are the four actions plus state and lookup',
  apiFunctions.sort(),
  ['apiGetState', 'apiLookup', 'apiPreviewWelcomeResend', 'apiResendWelcome',
    'apiResetPassword', 'apiSetGroupMembership'].sort());

apiFunctions.forEach(name => {
  const body = apiSource.slice(apiSource.indexOf('function ' + name));
  const upToNext = body.slice(0, body.indexOf('\n}\n') + 1);
  check(name + ' calls requireAdmin_()', /requireAdmin_\(\)/.test(upToNext), true);
});

done();
