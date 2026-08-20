/**
 * SecondaryDomainAliases.gs — the member-facing side of alias creation (1.4.0):
 * turning on Gmail "Send mail as" for the member, and the email that tells them.
 *
 * What matters here, and why each is pinned:
 *   1. Send-As setup is BEST-EFFORT and never throws — the alias is already
 *      created, so an admin account, missing credentials, or an API error must
 *      degrade to "send the manual steps," never fail the run.
 *   2. The email reflects reality: when we set sending up, it says so and skips the
 *      manual steps; when we could not, it walks the member through them. Either
 *      way it carries the CAPR 120-1 §6.9 official-use reminder.
 *   3. Every {{placeholder}} is substituted — a stray one ships raw to a member.
 *
 * Run: npm test
 */
const fs = require('fs');
const path = require('path');
const { loadModule, makeLogger } = require('./helpers/apps-script');
const { section, check, done } = require('./helpers/apps-script').makeChecker();

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'SecondaryDomainAliases.gs');
const TEMPLATE = fs.readFileSync(
  path.join(__dirname, '..', 'src', 'accounts-and-groups', 'SecondaryAliasWelcomeEmail.html'), 'utf8');

/** Builds the module with fakes, plus handles to what got sent / fetched. */
function build(opts) {
  opts = opts || {};
  // Impersonation credentials present by default, so the send-as path is exercised
  // rather than short-circuiting on 'unavailable'.
  const props = Object.assign(
    { SA_IMPERSONATION_EMAIL: 'sa@x.iam', SA_PRIVATE_KEY: '-----KEY-----' }, opts.props);
  const sent = [];
  const fetches = [];

  const globals = {
    Logger: makeLogger().logger,
    PropertiesService: {
      getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null) })
    },
    CONFIG: Object.assign({
      AUTOMATION_SENDER_EMAIL: 'it-automation@cawgcap.org',
      ITSUPPORT_EMAIL: 'itsupport@cawgcap.org',
      ORG_LABEL: 'CAWG',
      WING_NAME: 'California Wing'
    }, opts.config),
    SENDER_NAME: 'CAP Information Technology',
    executeWithRetry: fn => fn(),
    GmailApp: { sendEmail: (to, subject, body, options) => sent.push({ to, subject, body, options }) },
    HtmlService: { createHtmlOutputFromFile: () => ({ getContent: () => TEMPLATE }) },
    // Impersonation + REST, mirroring UpdateMembers.gs. Configurable per test.
    getImpersonatedToken_: () => {
      if (opts.tokenThrows) throw new Error('Token exchange failed (403)');
      return 'ya29.test-token';
    },
    UrlFetchApp: {
      fetch: (url, options) => {
        fetches.push({ url, options });
        const code = opts.fetchCode == null ? 200 : opts.fetchCode;
        return { getResponseCode: () => code, getContentText: () => opts.fetchText || '{}' };
      }
    },
    AdminDirectory: {}, SpreadsheetApp: {}, LockService: {}
  };

  const fns = loadModule(MODULE, globals, [
    'maybeSendSecondaryAliasWelcome_', 'renderSecondaryAliasWelcome_',
    'secondaryAliasNotifyEnabled_', 'htmlToPlainText_',
    'configureSecondaryAliasSendAs_', 'secondaryAliasSendAsReady_',
    'secondaryAliasSendAsSectionHtml_'
  ]);
  return { fns, sent, fetches, props };
}

// ---------------------------------------------------------------------------
section('1. secondaryAliasNotifyEnabled_ — opt-out, default on');
{
  check('unset -> enabled', build().fns.secondaryAliasNotifyEnabled_(), true);
  ['false', 'FALSE', '0', 'no', 'off', ' Off '].forEach(v => {
    check('"' + v + '" -> disabled',
      build({ props: { SECONDARY_ALIAS_NOTIFY: v } }).fns.secondaryAliasNotifyEnabled_(), false);
  });
  ['true', '1', 'yes', 'anything-else'].forEach(v => {
    check('"' + v + '" -> enabled',
      build({ props: { SECONDARY_ALIAS_NOTIFY: v } }).fns.secondaryAliasNotifyEnabled_(), true);
  });
}

// ---------------------------------------------------------------------------
section('2. configureSecondaryAliasSendAs_ — best-effort, never throws');
{
  // The happy path: POST the send-as identity and report it configured.
  const ok = build({ fetchCode: 200 });
  const r1 = ok.fns.configureSecondaryAliasSendAs_('jane.doe@cawgcap.org', 'jane.doe@cawg.cap.gov', 'Maj Jane Doe', false);
  check('a 2xx -> configured', r1, 'configured');
  check('posted exactly one request', ok.fetches.length, 1);
  check('to the sendAs endpoint', /settings\/sendAs$/.test(ok.fetches[0].url), true);
  check('as a POST', ok.fetches[0].options.method, 'post');
  const body = JSON.parse(ok.fetches[0].options.payload);
  check('for the alias address', body.sendAsEmail, 'jane.doe@cawg.cap.gov');
  check('treated as an alias', body.treatAsAlias, true);
  check('carries a bearer token', /^Bearer /.test(ok.fetches[0].options.headers.Authorization), true);

  // An existing identity is still "ready".
  check('a 409 -> exists', build({ fetchCode: 409 }).fns
    .configureSecondaryAliasSendAs_('a@b.org', 'a@c.gov', 'X', false), 'exists');

  // An admin cannot be impersonated for settings — skip WITHOUT calling anything.
  const adminBuild = build();
  check('an admin -> skipped-admin', adminBuild.fns
    .configureSecondaryAliasSendAs_('boss@cawgcap.org', 'boss@cawg.cap.gov', 'Boss', true), 'skipped-admin');
  check('an admin triggers no impersonation call', adminBuild.fetches.length, 0);

  // No credentials configured -> unavailable, nothing called.
  const noCreds = build({ props: { SA_IMPERSONATION_EMAIL: '', SA_PRIVATE_KEY: '' } });
  check('no credentials -> unavailable', noCreds.fns
    .configureSecondaryAliasSendAs_('a@b.org', 'a@c.gov', 'X', false), 'unavailable');
  check('no credentials -> no call', noCreds.fetches.length, 0);

  // API error and token error both degrade to 'failed', never throw.
  check('a 500 -> failed', build({ fetchCode: 500, fetchText: 'boom' }).fns
    .configureSecondaryAliasSendAs_('a@b.org', 'a@c.gov', 'X', false), 'failed');

  let threw = false, result;
  try {
    result = build({ tokenThrows: true }).fns
      .configureSecondaryAliasSendAs_('a@b.org', 'a@c.gov', 'X', false);
  } catch (e) { threw = true; }
  check('a token failure does NOT throw', threw, false);
  check('a token failure -> failed', result, 'failed');

  check('ready helper: configured/exists are ready, others are not', [
    build().fns.secondaryAliasSendAsReady_('configured'),
    build().fns.secondaryAliasSendAsReady_('exists'),
    build().fns.secondaryAliasSendAsReady_('skipped-admin'),
    build().fns.secondaryAliasSendAsReady_('failed'),
    build().fns.secondaryAliasSendAsReady_(undefined)
  ], [true, true, false, false, false]);
}

// ---------------------------------------------------------------------------
section('3. the email body adapts to whether we set sending up');
{
  const { fns } = build();

  const ready = fns.renderSecondaryAliasWelcome_(
    'jane.doe@cawgcap.org', 'jane.doe@cawg.cap.gov', 'Maj Jane Doe', 'configured');
  check('READY variant: says it is already set up', /already set up/i.test(ready), true);
  check('READY variant: points at the From field', ready.indexOf('From') !== -1, true);
  check('READY variant: does NOT hand out the manual "Add another email address" step',
    ready.indexOf('Add another email address') === -1, true);

  const manual = fns.renderSecondaryAliasWelcome_(
    'jane.doe@cawgcap.org', 'jane.doe@cawg.cap.gov', 'Maj Jane Doe', 'skipped-admin');
  check('MANUAL variant: walks through Send mail as', manual.indexOf('Send mail as') !== -1, true);
  check('MANUAL variant: names the Accounts tab', manual.indexOf('Accounts') !== -1, true);
  check('MANUAL variant: mentions treat as an alias', /treat as an alias/i.test(manual), true);

  // Both variants must carry the regulatory reminder and the identity, and neither
  // may leak a raw placeholder.
  [['READY', ready], ['MANUAL', manual]].forEach(([label, html]) => {
    check(label + ': greets the member', html.indexOf('Maj Jane Doe') !== -1, true);
    check(label + ': shows the alias', html.indexOf('jane.doe@cawg.cap.gov') !== -1, true);
    check(label + ': cites CAPR 120-1 §6.9', html.indexOf('CAPR 120-1') !== -1 && /6\.9/.test(html), true);
    check(label + ': names the forbidden uses', /commercial/i.test(html) && /fundrais/i.test(html), true);
    check(label + ': sets the ~day timing expectation', /up to a day/i.test(html), true);
    // The compose-window switching note the members most often miss: names the
    // From menu and BOTH addresses so it is clear you choose per message.
    check(label + ': explains switching in the compose window',
      /From/.test(html) && /compose/i.test(html), true);
    check(label + ': the switching note names both addresses',
      html.indexOf('jane.doe@cawgcap.org') !== -1 && html.indexOf('jane.doe@cawg.cap.gov') !== -1, true);
    check(label + ': no leftover placeholder', /{{\w+}}/.test(html), false);
  });

  check('a missing name falls back rather than greeting nobody',
    fns.renderSecondaryAliasWelcome_('a@b.org', 'a@c.gov', '', 'configured').indexOf('Member') !== -1, true);
}

// ---------------------------------------------------------------------------
section('4. maybeSendSecondaryAliasWelcome_ — sends correctly, fails safe');
{
  const happy = build();
  const ok = happy.fns.maybeSendSecondaryAliasWelcome_(
    'jane.doe@cawgcap.org', 'jane.doe@cawg.cap.gov', 'Maj Jane Doe', 'configured');
  check('reports it sent', ok, true);
  check('sent exactly one message', happy.sent.length, 1);
  check('to the member', happy.sent[0].to, 'jane.doe@cawgcap.org');
  check('from the automation sender, not the trigger account',
    happy.sent[0].options.from, 'it-automation@cawgcap.org');
  check('reply-to is IT support', happy.sent[0].options.replyTo, 'itsupport@cawgcap.org');
  check('subject names the address', happy.sent[0].subject.indexOf('jane.doe@cawg.cap.gov') !== -1, true);
  check('the HTML body reflects the ready state', /already set up/i.test(happy.sent[0].options.htmlBody), true);
  check('a non-empty plain-text alternative', happy.sent[0].body.length > 50, true);
  check('plain text is stripped of tags', /<[a-z]/i.test(happy.sent[0].body), false);

  const off = build({ props: { SECONDARY_ALIAS_NOTIFY: 'false' } });
  check('disabled -> returns false', off.fns.maybeSendSecondaryAliasWelcome_('a@b.org', 'a@c.gov', 'X', 'configured'), false);
  check('disabled -> sends nothing', off.sent.length, 0);

  const noSender = build({ config: { AUTOMATION_SENDER_EMAIL: '' } });
  check('no sender configured -> returns false',
    noSender.fns.maybeSendSecondaryAliasWelcome_('a@b.org', 'a@c.gov', 'X', 'configured'), false);
  check('no sender configured -> sends nothing', noSender.sent.length, 0);

  // A send that throws is swallowed — the alias is already created.
  const throwing = (function () {
    const props = { SA_IMPERSONATION_EMAIL: 's', SA_PRIVATE_KEY: 'k' };
    const g = {
      Logger: makeLogger().logger,
      PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null) }) },
      CONFIG: { AUTOMATION_SENDER_EMAIL: 's@x.org', ITSUPPORT_EMAIL: '', ORG_LABEL: 'CAWG', WING_NAME: 'California Wing' },
      SENDER_NAME: 'IT', executeWithRetry: fn => fn(),
      GmailApp: { sendEmail: () => { throw new Error('Service invoked too many times'); } },
      HtmlService: { createHtmlOutputFromFile: () => ({ getContent: () => TEMPLATE }) },
      getImpersonatedToken_: () => 't', UrlFetchApp: { fetch: () => ({ getResponseCode: () => 200, getContentText: () => '{}' }) },
      AdminDirectory: {}, SpreadsheetApp: {}, LockService: {}
    };
    return loadModule(MODULE, g, ['maybeSendSecondaryAliasWelcome_']);
  })();
  let threw = false, result;
  try { result = throwing.maybeSendSecondaryAliasWelcome_('a@b.org', 'a@c.gov', 'X', 'configured'); }
  catch (e) { threw = true; }
  check('a send failure does NOT throw (alias must not look failed)', threw, false);
  check('a send failure reports false', result, false);
}

done();
