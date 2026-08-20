/**
 * secondary-alias-webapp/Notify.gs — the immediate Send-As + welcome-email process
 * the Add page runs, ported from the nightly module.
 *
 * WHY THIS TEST EXISTS: an alias the web app creates reads as "already present" by
 * the time the nightly run fires, so the nightly notifier skips it — the web app
 * MUST do Send-As and the email itself. This is a hand copy of
 * src/accounts-and-groups/SecondaryDomainAliases.gs (1.4.0), so the two drift
 * silently unless something pins them. This does both: the same behavioral checks
 * as the src test, plus a direct parity check that the rendered emails match.
 *
 * Run: npm test
 */
const fs = require('fs');
const path = require('path');
const { loadModule, makeLogger } = require('./helpers/apps-script');
const { section, check, done } = require('./helpers/apps-script').makeChecker();

const WEBAPP = path.join(__dirname, '..', 'secondary-alias-webapp', 'Notify.js');
const SRC = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'SecondaryDomainAliases.gs');
// Normalize line endings: we assert CONTENT parity between the two hand-copies,
// not their working-tree EOL — which drifts to CRLF on Windows under autocrlf and
// is irrelevant to what Apps Script renders.
const readLF = p => fs.readFileSync(p, 'utf8').replace(/\r\n/g, '\n');
const WEB_TEMPLATE = readLF(
  path.join(__dirname, '..', 'secondary-alias-webapp', 'SecondaryAliasWelcomeEmail.html'));
const SRC_TEMPLATE = readLF(
  path.join(__dirname, '..', 'src', 'accounts-and-groups', 'SecondaryAliasWelcomeEmail.html'));

/** Loads the web app's Notify.gs with fakes; returns fns + captured effects. */
function build(opts) {
  opts = opts || {};
  const props = Object.assign(
    { SA_IMPERSONATION_EMAIL: 'sa@x.iam', SA_PRIVATE_KEY: '-----KEY-----' }, opts.props);
  const sent = [];
  const fetches = [];
  const globals = {
    Logger: makeLogger().logger,
    PropertiesService: {
      getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null) })
    },
    WEBAPP_CONFIG: Object.assign({
      AUTOMATION_SENDER_EMAIL: 'it-automation@cawgcap.org',
      ITSUPPORT_EMAIL: 'itsupport@cawgcap.org',
      ORG_LABEL: 'CAWG',
      WING_NAME: 'California Wing',
      SENDER_NAME: 'CAP Information Technology'
    }, opts.config),
    GmailApp: { sendEmail: (to, subject, body, options) => sent.push({ to, subject, body, options }) },
    HtmlService: { createHtmlOutputFromFile: () => ({ getContent: () => WEB_TEMPLATE }) },
    UrlFetchApp: {
      fetch: (url, options) => {
        fetches.push({ url, options });
        // The impersonation token exchange and the sendAs POST are two different
        // calls to the same fake; keep them independent as they are in reality.
        if (String(url).indexOf('oauth2.googleapis.com/token') !== -1) {
          const tcode = opts.tokenCode == null ? 200 : opts.tokenCode;
          return { getResponseCode: () => tcode, getContentText: () => JSON.stringify({ access_token: 'ya29.tok' }) };
        }
        const code = opts.fetchCode == null ? 200 : opts.fetchCode;
        return { getResponseCode: () => code, getContentText: () => opts.fetchText || '{}' };
      }
    },
    Utilities: {
      base64EncodeWebSafe: s => Buffer.from(String(s)).toString('base64'),
      computeRsaSha256Signature: () => [1, 2, 3]
    }
  };
  const fns = loadModule(WEBAPP, globals, [
    'webappConfigureSendAs_', 'webappSendAsReady_', 'webappMaybeSendWelcome_',
    'webappRenderWelcome_', 'webappSendAsSectionHtml_', 'webappHtmlToPlainText_'
  ]);
  return { fns, sent, fetches, props };
}

// ---------------------------------------------------------------------------
section('1. the ported template is byte-identical to src (no silent drift)');
{
  check('SecondaryAliasWelcomeEmail.html matches src', WEB_TEMPLATE === SRC_TEMPLATE, true);
}

// ---------------------------------------------------------------------------
section('2. webappConfigureSendAs_ — best-effort, never throws');
{
  const ok = build({ fetchCode: 200 });
  check('a 2xx -> configured', ok.fns.webappConfigureSendAs_('j@cawgcap.org', 'j@cawg.cap.gov', 'Maj J', false), 'configured');
  check('posts to the sendAs endpoint', /settings\/sendAs$/.test(ok.fetches[ok.fetches.length - 1].url), true);
  const body = JSON.parse(ok.fetches[ok.fetches.length - 1].options.payload);
  check('for the alias, treated as an alias', [body.sendAsEmail, body.treatAsAlias], ['j@cawg.cap.gov', true]);

  check('a 409 -> exists', build({ fetchCode: 409 }).fns
    .webappConfigureSendAs_('a@b.org', 'a@c.gov', 'X', false), 'exists');

  const admin = build();
  check('an admin -> skipped-admin', admin.fns.webappConfigureSendAs_('boss@b', 'boss@c', 'B', true), 'skipped-admin');
  check('an admin calls nothing', admin.fetches.length, 0);

  const noCreds = build({ props: { SA_IMPERSONATION_EMAIL: '', SA_PRIVATE_KEY: '' } });
  check('no credentials -> unavailable', noCreds.fns.webappConfigureSendAs_('a@b', 'a@c', 'X', false), 'unavailable');
  check('no credentials calls nothing', noCreds.fetches.length, 0);

  check('a 500 -> failed', build({ fetchCode: 500 }).fns.webappConfigureSendAs_('a@b', 'a@c', 'X', false), 'failed');

  // A token-exchange failure (the OAuth call itself 4xxs) must not throw.
  const tokenFail = build({ tokenCode: 401, fetchText: 'bad jwt' });
  let threw = false, res;
  try { res = tokenFail.fns.webappConfigureSendAs_('a@b', 'a@c', 'X', false); } catch (e) { threw = true; }
  check('a token failure does not throw', threw, false);
  check('a token failure -> failed', res, 'failed');
}

// ---------------------------------------------------------------------------
section('3. the email adapts, carries CAPR §6.9 + the compose-switch note, sends safe');
{
  const { fns } = build();
  const ready = fns.webappRenderWelcome_('j@cawgcap.org', 'j@cawg.cap.gov', 'Maj J', 'configured');
  const manual = fns.webappRenderWelcome_('j@cawgcap.org', 'j@cawg.cap.gov', 'Maj J', 'skipped-admin');

  check('READY says already set up', /already set up/i.test(ready), true);
  check('READY omits the manual add step', ready.indexOf('Add another email address') === -1, true);
  check('MANUAL walks through Send mail as', manual.indexOf('Send mail as') !== -1, true);

  [['READY', ready], ['MANUAL', manual]].forEach(([label, html]) => {
    check(label + ': cites CAPR 120-1 §6.9', html.indexOf('CAPR 120-1') !== -1 && /6\.9/.test(html), true);
    check(label + ': names the forbidden uses', /commercial/i.test(html) && /fundrais/i.test(html), true);
    check(label + ': explains switching in Compose', /From/.test(html) && /compose/i.test(html), true);
    check(label + ': no leftover placeholder', /{{\w+}}/.test(html), false);
  });

  const happy = build();
  const okSend = happy.fns.webappMaybeSendWelcome_('j@cawgcap.org', 'j@cawg.cap.gov', 'Maj J', 'configured');
  check('sends and reports true', [okSend, happy.sent.length], [true, 1]);
  check('to the member, from the automation sender',
    [happy.sent[0].to, happy.sent[0].options.from], ['j@cawgcap.org', 'it-automation@cawgcap.org']);
  check('plain-text alternative is tag-free', /<[a-z]/i.test(happy.sent[0].body), false);

  const noSender = build({ config: { AUTOMATION_SENDER_EMAIL: '' } });
  check('no sender -> false, nothing sent',
    [noSender.fns.webappMaybeSendWelcome_('a@b', 'a@c', 'X', 'configured'), noSender.sent.length], [false, 0]);
}

// ---------------------------------------------------------------------------
section('4. PARITY — web app and src render the same email for the same inputs');
{
  // Load the src copy with the src template and compare its output to the web app's,
  // so a change to one that is not mirrored to the other fails here.
  const srcGlobals = {
    Logger: makeLogger().logger,
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    CONFIG: { ORG_LABEL: 'CAWG', WING_NAME: 'California Wing' },
    HtmlService: { createHtmlOutputFromFile: () => ({ getContent: () => SRC_TEMPLATE }) }
  };
  const src = loadModule(SRC, srcGlobals, ['renderSecondaryAliasWelcome_']);
  const web = build().fns;

  [['configured', 'ready'], ['skipped-admin', 'manual'], ['failed', 'manual']].forEach(([outcome]) => {
    const a = src.renderSecondaryAliasWelcome_('j@cawgcap.org', 'j@cawg.cap.gov', 'Maj J', outcome);
    const b = web.webappRenderWelcome_('j@cawgcap.org', 'j@cawg.cap.gov', 'Maj J', outcome);
    check('identical rendered email for outcome "' + outcome + '"', a === b, true);
  });
}

// ---------------------------------------------------------------------------
section('5. webappHtmlToPlainText_ — no residual markup, no double-unescape (CodeQL)');
{
  const t = build().fns.webappHtmlToPlainText_;
  check('strips whole tags', /<[a-z/]/i.test(t('<p>Hi <strong>there</strong></p>')), false);
  check('removes script blocks entirely, content included', t('<script>alert(1)</script>Hello'), 'Hello');
  check('removes style blocks entirely', t('<style>.a{}</style>Hello'), 'Hello');
  // The incomplete-sanitization worry: an unclosed "<script" must not survive.
  check('an unclosed <script leaves no angle bracket', /[<>]/.test(t('safe <script foo')), false);
  // The double-unescape worry: &amp; decodes to a literal & in ONE pass and does
  // not re-form an entity for a second pass.
  check('&amp; decodes once', t('a &amp; b'), 'a & b');
  check('&amp;bull; does not become a bullet', t('x &amp;bull; y'), 'x &bull; y');
  check('real entities still decode', t('<p>see &rarr; here &bull; done</p>'), 'see -> here • done');
}

// ---------------------------------------------------------------------------
section('6. Config wing derivation — TENANT_WING alone yields CAWG / California Wing');
{
  // Config.gs declares its own Logger, so do NOT inject one (a param + `const
  // Logger` would be a redeclaration). It only needs PropertiesService at load.
  function cfg(props) {
    return loadModule(
      path.join(__dirname, '..', 'secondary-alias-webapp', 'Config.js'),
      { PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null) }) } },
      ['webappWingName_', 'webappWingAbbreviation_']);
  }
  const c = cfg({});

  check('CA -> California Wing', c.webappWingName_('CA', ''), 'California Wing');
  check('CA -> CAWG', c.webappWingAbbreviation_('CA', ''), 'CAWG');
  check('lower-case ca still resolves', c.webappWingName_('ca', ''), 'California Wing');
  check('HI -> Hawaii Wing', c.webappWingName_('HI', ''), 'Hawaii Wing');
  check('an explicit name override wins', c.webappWingName_('CA', 'Some Other Wing'), 'Some Other Wing');
  check('an explicit abbreviation override wins', c.webappWingAbbreviation_('CA', 'cawg'), 'CAWG');
  check('an unknown code falls back to its abbreviation', c.webappWingName_('ZZ', ''), 'ZZWG');
  check('no wing set at all -> CAP, not an empty masthead', c.webappWingAbbreviation_('', ''), 'CAP');
}

done();
