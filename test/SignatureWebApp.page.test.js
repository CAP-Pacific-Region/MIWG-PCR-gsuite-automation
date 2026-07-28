/**
 * signature-webapp/Index.html — the page's own logic.
 *
 * WHY THIS FILE EXISTS
 *
 * The server side was tested to the point of pedantry: only an explicit `true`
 * includes the phone row, every other value suppresses it, and the render is
 * rebuilt from CAPWATCH on apply. All of that was correct, and a member still had
 * their phone number silently dropped from a signature they had ticked the box
 * for — because the BROWSER sent false.
 *
 *   apply() called setBusy(true), which disables the checkbox to show a request is
 *   in flight. includePhone() then read `!box.disabled` as part of deciding the
 *   answer. So by the time the value was computed, the box was always disabled and
 *   the answer was always false. The server obediently did as it was told and
 *   reported "the phone row was left out" — about a box the member had ticked.
 *
 * The lesson: a value the user chose must be READ BEFORE the UI starts mutating
 * itself for the request, and "is this control currently disabled" is not an
 * opinion about what the user wants. The unit under test was never the signature
 * generator; it was the wire between the checkbox and the call.
 *
 * So this loads the real <script> out of Index.html over a small fake DOM and
 * asserts what actually reaches google.script.run. All members here are synthetic.
 *
 * Run: npm test
 */
const fs = require('fs');
const path = require('path');
const { makeChecker } = require('./helpers/apps-script');

const { section, check, done } = makeChecker();

const PAGE = path.join(__dirname, '..', 'signature-webapp', 'Index.html');

/** Every element id the page reaches for, with just enough behavior to run. */
const IDS = ['flash', 'loading', 'app', 'mName', 'mCapid', 'phoneChoice',
  'includePhone', 'phoneValue', 'noPhone', 'noPhoneText', 'preview', 'addrIntro',
  'addrs', 'applyBtn', 'applySpin', 'dutyChoice', 'dutyList', 'dutyNote', 'noDuty',
  'renameNote'];

/**
 * The starting `class` of an element, read out of the markup rather than assumed.
 *
 * It matters: `phoneChoice` ships as "card hidden" and the page only ever assigns
 * a class when it decides to SHOW something. A stub that started every element
 * blank would report a hidden panel as visible and vice versa — the test would be
 * describing itself.
 */
function initialClass(html, id) {
  const tag = html.match(new RegExp('<[^>]*\\bid="' + id + '"[^>]*>'));
  if (!tag) return '';
  const cls = tag[0].match(/\bclass="([^"]*)"/);
  return cls ? cls[1] : '';
}

function makeElement(html, id) {
  return {
    id: id,
    className: initialClass(html, id),
    textContent: '',
    innerHTML: '',
    srcdoc: '',
    disabled: false,
    checked: /<[^>]*\bid="includePhone"[^>]*\bchecked\b/.test(html) && id === 'includePhone',
    handlers: {},
    addEventListener: function (event, fn) { this.handlers[event] = fn; }
  };
}

/**
 * Loads the page script over a fresh fake DOM.
 *
 * @param {Object} state - what apiGetState() returns to the page
 * @returns {Object} { el, calls } — calls records every server call's argument
 */
function loadPage(state) {
  const html = fs.readFileSync(PAGE, 'utf8');
  const source = html.slice(html.lastIndexOf('<script>') + '<script>'.length,
    html.lastIndexOf('</script>'));

  const elements = {};
  IDS.forEach(id => { elements[id] = makeElement(html, id); });

  // The duty checkboxes do not exist in the markup — the page writes them into
  // #dutyList and then looks each one up by id. So the stub has to mint them on
  // assignment, the way a browser would, or the page would be testing nothing.
  let listHtml = '';
  Object.defineProperty(elements.dutyList, 'innerHTML', {
    get: () => listHtml,
    set: value => {
      listHtml = value;
      const box = /<input type="checkbox" id="(duty\d+)"( checked)?>/g;
      let match;
      while ((match = box.exec(value)) !== null) {
        const created = makeElement(html, match[1]);
        created.checked = !!match[2];
        elements[match[1]] = created;
      }
    }
  });

  const calls = { apiPreview: [], apiApply: [], dutyKeys: [] };

  // google.script.run's chainable shape, resolving synchronously.
  function makeRunner() {
    let onSuccess = function () {};
    let onFailure = function () {};
    const runner = {
      withSuccessHandler: function (fn) { onSuccess = fn; return runner; },
      withFailureHandler: function (fn) { onFailure = fn; return runner; },
      apiGetState: function () { onSuccess(state); },
      apiPreview: function (arg, dutyKeys) {
        calls.apiPreview.push(arg);
        calls.dutyKeys.push(dutyKeys);
        onSuccess({ html: '<html>preview ' + String(arg) + '</html>', includedPhone: arg === true });
      },
      apiApply: function (arg, dutyKeys) {
        calls.apiApply.push(arg);
        calls.dutyKeys.push(dutyKeys);
        // Mirrors the server: an apply that would write to nothing THROWS rather
        // than returning an empty success (see apiApply in SignatureApi.gs). A stub
        // that resolved anyway would let the page pass tests it fails in life.
        if (!state.addresses.length) {
          onFailure(new Error('None of the send-as addresses on your mailbox belong to ' +
            state.orgLabel + ', so there was nothing to update.'));
          return;
        }
        onSuccess({
          updated: state.addresses, skipped: [], failed: [],
          includedPhone: arg === true && !!state.phone
        });
      }
    };
    return runner;
  }

  const document = { getElementById: id => elements[id] || null };
  const google = { script: { get run() { return makeRunner(); } } };

  new Function('document', 'google', source)(document, google);
  return { el: elements, calls: calls };
}

const DUTIES = [
  { key: 'Director of IT|WING|CALIFORNIA WING HQ|P', line: 'California Wing Director of IT',
    level: 'WING', assistant: false, selected: true },
  { key: 'Safety Officer|WING|CALIFORNIA WING HQ|A',
    line: 'California Wing Assistant Safety Officer',
    level: 'WING', assistant: true, selected: false },
  { key: 'Commander|UNIT|EXAMPLE SR SQDN 80|P', line: 'Example Senior Squadron 80 Commander',
    level: 'UNIT', assistant: false, selected: true }
];

const SENIOR_STATE = {
  actor: 'rowan.ashford@example.org',
  orgLabel: 'CAWG',
  capid: '600001',
  name: 'Maj. Rowan Ashford',
  dutyLines: ['California Wing Director of IT', 'Example Senior Squadron 80 Commander'],
  duties: DUTIES,
  maxDuties: 2,
  phone: '555.010.1234',
  isCadet: false,
  addresses: ['rowan.ashford@example.org', 'rowan.ashford@example.cap.gov'],
  addressError: '',
  html: '<html>initial</html>'
};

/** The duty checkbox stubs the page minted, in render order. */
function dutyBoxes(page) {
  return DUTIES.map((d, i) => page.el['duty' + i]).filter(Boolean);
}

// ---------------------------------------------------------------------------
section('1. A ticked box reaches the server as true');
{
  const page = loadPage(SENIOR_STATE);

  check('the page loaded and showed the record', page.el.mName.textContent, 'Maj. Rowan Ashford');
  check('the duty choices are offered', page.el.dutyChoice.className, 'card');
  check('the phone choice is offered', page.el.phoneChoice.className, 'card');
  check('...ticked, with the number beside it', page.el.phoneValue.textContent, '555.010.1234');

  // THE REGRESSION. Clicking apply disables the checkbox on its way out; the value
  // must already have been read.
  page.el.applyBtn.handlers.click();
  check('apply sends true for a ticked box', page.calls.apiApply, [true]);
  check('...even though applying disabled the checkbox on the way',
    page.el.includePhone.disabled, false);   // setBusy(false) ran in the handler

  // And the message the member reads is the honest one.
  check('so the member is not told their phone was left out',
    page.el.flash.textContent.indexOf('left out'), -1);
  check('the flash reports the addresses written',
    page.el.flash.textContent.indexOf('rowan.ashford@example.cap.gov') !== -1, true);
}

// ---------------------------------------------------------------------------
section('2. The toggle survives a round trip');
{
  const page = loadPage(SENIOR_STATE);

  page.el.includePhone.checked = false;
  page.el.includePhone.handlers.change();
  check('unticking previews without the phone', page.calls.apiPreview, [false]);

  // The same disabled-checkbox trap lived here: re-ticking used to preview WITHOUT
  // the phone, so the control looked broken as well as acting wrongly.
  page.el.includePhone.checked = true;
  page.el.includePhone.handlers.change();
  check('re-ticking brings it back', page.calls.apiPreview, [false, true]);
  check('and the preview shown is the one just fetched',
    page.el.preview.srcdoc, '<html>preview true</html>');

  page.el.applyBtn.handlers.click();
  check('apply then agrees with the box', page.calls.apiApply, [true]);
}

// ---------------------------------------------------------------------------
section('3. A member with no phone cannot send true');
{
  // A cadet, or a senior with no publishable cell: the checkbox is still in the
  // markup and still ticked, but there is no number for it to mean anything about.
  const page = loadPage(Object.assign({}, SENIOR_STATE, {
    phone: '', isCadet: true, name: 'Cadet Chief Master Sgt. Imani Brightwater'
  }));

  check('no toggle is offered', page.el.phoneChoice.className, 'card hidden');
  check('...an explanation is', page.el.noPhone.className, 'card');
  check('and it says why', page.el.noPhoneText.textContent.indexOf('Cadet phone numbers') !== -1, true);

  page.el.applyBtn.handlers.click();
  check('the ticked-by-default box cannot publish a number nobody has',
    page.calls.apiApply, [false]);
  check('and the member is not told a phone row was dropped',
    page.el.flash.textContent.indexOf('left out'), -1);
}

// ---------------------------------------------------------------------------
section('4. Nothing to write to');
{
  const page = loadPage(Object.assign({}, SENIOR_STATE, { addresses: [] }));

  check('apply is disabled when no org address was found', page.el.applyBtn.disabled, true);
  check('...and the page says so',
    page.el.addrIntro.textContent.indexOf('nothing to write to') !== -1, true);

  // A preview toggle must not re-enable it: setBusy(false) consults STATE.canApply.
  page.el.includePhone.handlers.change();
  check('a preview round trip leaves it disabled', page.el.applyBtn.disabled, true);
}

// ---------------------------------------------------------------------------
section('5. Duty selection: the cap is visible, and what is ticked is what is sent');
{
  const page = loadPage(SENIOR_STATE);
  const boxes = dutyBoxes(page);

  check('every duty the member holds is offered, assistants included', boxes.length, 3);
  check('the default pick starts ticked',
    boxes.map(b => b.checked), [true, false, true]);
  // The server renders each option's line, and an assistant duty's line says so
  // itself ("… Assistant Supply Officer"). The page must show that line verbatim —
  // a separate badge would repeat it, and rewriting it here would be a second
  // implementation of a title the generator owns.
  check('an assistant duty reads as one, from the line the server sent',
    page.el.dutyList.innerHTML.indexOf('California Wing Assistant Safety Officer') !== -1, true);
  check('...and the page adds no badge of its own',
    page.el.dutyList.innerHTML.indexOf('class="tag"'), -1);

  // At the cap, the unticked ones grey out rather than letting a member tick a
  // third and be refused by the server.
  check('at the cap, the unticked box is disabled', boxes[1].disabled, true);
  check('...and the ticked ones stay live', [boxes[0].disabled, boxes[2].disabled], [false, false]);
  check('the note says how many may be picked',
    page.el.dutyNote.textContent.indexOf('Pick up to 2') !== -1, true);
  check('...and that order is not the member\'s to set',
    page.el.dutyNote.textContent.indexOf('highest organizational level first') !== -1, true);

  // Untick one: the third becomes available again.
  boxes[2].checked = false;
  boxes[2].handlers.change();
  check('unticking frees the cap', boxes[1].disabled, false);
  check('...and previews the remaining choice',
    page.calls.dutyKeys[page.calls.dutyKeys.length - 1], [DUTIES[0].key]);

  // Swap the squadron command for the wing assistant.
  boxes[1].checked = true;
  boxes[1].handlers.change();
  check('ticking another sends both', page.calls.dutyKeys[page.calls.dutyKeys.length - 1],
    [DUTIES[0].key, DUTIES[1].key]);

  page.el.applyBtn.handlers.click();
  check('and apply sends exactly what is ticked',
    page.calls.dutyKeys[page.calls.dutyKeys.length - 1], [DUTIES[0].key, DUTIES[1].key]);
  check('...alongside the phone choice', page.calls.apiApply, [true]);
}

// ---------------------------------------------------------------------------
section('6. Duty selection: the edges');
{
  // Touching nothing must still send the default pick, so the preview a member
  // read and the signature they get cannot disagree.
  const untouched = loadPage(SENIOR_STATE);
  untouched.el.applyBtn.handlers.click();
  check('an untouched page applies the default pick explicitly',
    untouched.calls.dutyKeys[0], [DUTIES[0].key, DUTIES[2].key]);

  // Every box off is a real choice: no duty line at all.
  const cleared = loadPage(SENIOR_STATE);
  const boxes = dutyBoxes(cleared);
  boxes.forEach(b => { if (b.checked) { b.checked = false; b.handlers.change(); } });
  cleared.el.applyBtn.handlers.click();
  check('unticking everything sends an empty selection, not a missing one',
    cleared.calls.dutyKeys[cleared.calls.dutyKeys.length - 1], []);
  check('...and the member is told the duty line was left out',
    cleared.el.flash.textContent.indexOf('No duty assignment was shown') !== -1, true);

  // A member with nothing on their record gets an explanation, not an empty card.
  const none = loadPage(Object.assign({}, SENIOR_STATE, { duties: [], dutyLines: [] }));
  check('no duties on file: no chooser', none.el.dutyChoice.className, 'card hidden');
  check('...an explanation instead', none.el.noDuty.className, 'card');
  none.el.applyBtn.handlers.click();
  check('...and an empty selection is sent', none.calls.dutyKeys[0], []);
  check('...without claiming a duty line was dropped',
    none.el.flash.textContent.indexOf('No duty assignment'), -1);

  // One duty: the cap is irrelevant, so the note should not talk about picking two.
  const single = loadPage(Object.assign({}, SENIOR_STATE, { duties: [DUTIES[0]] }));
  check('a single duty gets its own wording',
    single.el.dutyNote.textContent.indexOf('Untick it to leave the duty line out') !== -1, true);
}

// ---------------------------------------------------------------------------
section('7. The one thing the API cannot do is explained, not attempted');
{
  // Gmail's multiple-signature feature gives each signature a NAME. The SendAs
  // resource has no field for it — signature is the HTML and nothing else — so the
  // only place it can be changed is Gmail itself. Saying so beats a member hunting
  // for a control that could not exist.
  const page = loadPage(SENIOR_STATE);

  check('the advice is not shown before anything was written',
    page.el.renameNote.className, 'note hidden');

  page.el.applyBtn.handlers.click();
  check('...and appears once a signature exists to rename', page.el.renameNote.className, 'note');

  const html = fs.readFileSync(PAGE, 'utf8');
  check('it names the Gmail path a member can actually follow',
    /Settings .*General .*Signature/.test(html.replace(/&rarr;/g, '')), true);
  check('...and says the name is not something this page withheld',
    html.indexOf('Google&rsquo;s API has no field for it') !== -1, true);
  check('...nor something recipients ever see',
    html.indexOf('never appears on mail you send') !== -1, true);

  // A failed apply must not tell a member to go rename something that isn't there.
  const failing = loadPage(Object.assign({}, SENIOR_STATE, { addresses: [] }));
  failing.el.applyBtn.handlers.click();
  check('an apply that wrote nothing leaves the advice hidden',
    failing.el.renameNote.className.indexOf('hidden') !== -1, true);
  check('...and shows the server\'s reason instead',
    failing.el.flash.textContent.indexOf('nothing to update') !== -1, true);
}

// ---------------------------------------------------------------------------
section('8. The authority for the format is named, and reachable');
{
  // Not decoration: a member whose signature looks wrong needs somewhere that is
  // not "because the wing's script says so", and a member this app turns away
  // needs a way to get a correct signature by hand.
  const html = fs.readFileSync(PAGE, 'utf8');

  check('the footer links to the generator, by template variable rather than a copy',
    html.indexOf('href="<?= generatorUrl ?>"') !== -1, true);
  check('...opening it does not navigate away from an unsaved choice',
    /<a href="<\?= generatorUrl \?>"[^>]*target="_blank"/.test(html), true);
  check('...with rel=noopener', /<a href="<\?= generatorUrl \?>"[^>]*rel="noopener"/.test(html), true);
  check('and it is outside #app, so it is there before the record loads',
    html.indexOf('class="foot"') > html.indexOf('<div id="app"'), true);

  // The URL itself lives in Config.gs, once.
  const config = fs.readFileSync(path.join(__dirname, '..', 'signature-webapp', 'Config.gs'), 'utf8');
  check('the URL is defined once, server-side',
    /const SIG_GENERATOR_URL = 'https:\/\/cap-brand-tools\.netlify\.app\/signature-generator\/index\.html'/.test(config),
    true);
  check('and the page hard-codes no copy of it', html.indexOf('cap-brand-tools'), -1);
}

done();
