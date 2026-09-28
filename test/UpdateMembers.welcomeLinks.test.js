/**
 * UpdateMembers.gs — the welcome email's per-tenant help links.
 *
 * Both links used to be hardcoded (California's help-desk site, the Pacific Region
 * ticket portal), so every other wing's new members were sent to someone else's
 * help desk. Blank now omits the link. All URLs and addresses are synthetic.
 *
 * Run: npm test
 */
const fs = require('fs');
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const { section, check, done } = makeChecker();
const TEMPLATE = fs.readFileSync(
  path.join(__dirname, '..', 'src', 'recruiting-and-retention', 'WelcomeEmail.html'), 'utf8');
const { applyWelcomeEmailLinks_ } = loadModule(
  path.join(__dirname, '..', 'src', 'accounts-and-groups', 'UpdateMembers.gs'),
  { Logger: makeLogger().logger }, ['applyWelcomeEmailLinks_']);

const HELP = 'https://help.example.org/workspace';
const SUP = 'https://support.example.org';
const render = (h, s, m) => applyWelcomeEmailLinks_(TEMPLATE, h, s, m);

section('1. The template carries no wing-specific literals');
check('no California help-desk site', /cawgintranet|cawgcap\.org/i.test(TEMPLATE), false);
check('no Pacific Region ticket portal', /pcrcap\.org/i.test(TEMPLATE), false);

section('2. Both URLs set');
{
  const h = render(HELP, SUP, 'it@example.org');
  check('quick-reference link is the configured one', h.includes('href="' + HELP + '"'), true);
  check('support link is the configured one', h.includes('href="' + SUP + '"'), true);
  check('no placeholder left behind', /{{(HELP_GUIDE_URL|SUPPORT_BLOCK)}}|HELP_GUIDE_START/.test(h), false);
}

section('3. No help guide: the whole section goes');
{
  const h = render('', SUP, 'it@example.org');
  check('no quick-reference text', /Quick\s+Reference/.test(h), false);
  check('support still present', h.includes(SUP), true);
}

section('4. No support URL: falls back to mailing IT');
{
  const h = render(HELP, '', 'it@example.org');
  check('mailto the IT address', h.includes('href="mailto:it@example.org"'), true);
  check('no ticket wording', /support ticket/.test(h), false);
}

section('5. Nothing configured: no dead links, no stray markup');
{
  const h = render('', '', '');
  check('no support section at all', /IT Support:/.test(h), false);
  check('no quick-reference section', /Quick\s+Reference/.test(h), false);
  check('no unfilled placeholder', /{{(HELP_GUIDE_URL|SUPPORT_BLOCK)}}/.test(h), false);
}

section('6. A bad value is dropped, never emitted');
{
  const h = render('javascript:alert(1)', 'not a url', 'it@example.org');
  check('javascript: link is not rendered', /javascript:/.test(h), false);
  check('a non-URL support value falls back to mail', h.includes('mailto:it@example.org'), true);
  const q = render('https://x.example.org/a"b', SUP, '');
  check('a quote cannot break out of the attribute', q.includes('a&quot;b') && !q.includes('a"b'), true);
}

done();
