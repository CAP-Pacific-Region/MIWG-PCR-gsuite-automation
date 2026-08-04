/***********************************************
 * File: SignatureApi.gs
 * Description: Web app entry point and the server functions the browser calls.
 * Shows a member the signature CAPWATCH says they should have, and — only when
 * they say so — writes it to their own CAP addresses.
 * Author: Maj Isaac Wilson IV, California Wing
 * Version: 1.1.0
 * Date: 2026-07-28
 * Changes: 1.1.0 — apiGetState() offers every duty the member holds with the
 *            default pick marked; apiPreview/apiApply take a duty selection, which
 *            sigValidatedDutyKeys_() checks against the member's own CAPWATCH
 *            record and refuses rather than trims.
 *          1.0.1 — the pages that refuse now hand the member CAP's own signature
 *            generator, so a refusal is not a dead end.
 *          1.0.0 — initial version.
 ***********************************************/

/**
 * WHAT A MEMBER MAY CHANGE, AND WHY IT IS EXACTLY ONE THING
 *
 * The signature is a uniform item. Its content comes from CAPWATCH — grade, name,
 * duty assignments, unit — and its layout comes from the CAP brand style guide;
 * neither is a member's to edit, and a signature that says something CAPWATCH does
 * not is worse than no signature at all. The remedy for a wrong grade or a missing
 * duty is a correction in eServices, which then flows here on the next extract.
 *
 * The single exception is the phone number. It is the only genuinely personal
 * detail in the block, publishing it is a choice, and the style guide treats the
 * line as optional. So the client may send ONE value — a boolean — and it can only
 * ever suppress a row that is already the member's own.
 *
 * Everything else follows from that. The browser never sends HTML, never names an
 * account, and never supplies a name, grade or duty. The server rebuilds the whole
 * signature from CAPWATCH on every call, including on apply — the preview is a
 * rendering of the record, not a document the client can hand back edited.
 */

/**
 * Serves the UI. A caller the app cannot help gets a plain refusal rather than a
 * shell that fails on first click.
 */
function doGet() {
  // Named in the log, not on the page: a member cannot act on a property name, and
  // the person who can is reading the execution log anyway.
  const missing = sigMissingConfig_();
  if (missing.length) {
    Logger.error('Serving the not-configured page: required Script Properties are not set', {
      missing: missing
    });
    return HtmlService.createHtmlOutput(
      '<p style="font:14px/1.5 system-ui,sans-serif;padding:2rem">' +
      'This page is not set up yet.<br>' + sigEscape_(sigSupportSentence_()) +
      sigGeneratorFallbackHtml_() + '</p>'
    ).setTitle('Not set up yet');
  }

  const actor = resolveActor_();
  if (!mayUseSignatureApp_(actor)) {
    Logger.warn('Refused to serve the signature page', { user: actor || '(no identity)' });
    return HtmlService.createHtmlOutput(
      '<p style="font:14px/1.5 system-ui,sans-serif;padding:2rem">' +
      'This page is for ' + sigEscape_(SIG_CONFIG.ORG_LABEL) + ' members signed in to their ' +
      sigEscape_(SIG_CONFIG.ORG_LABEL) + ' account.<br>' +
      'Sign in to that account and open this link again.' +
      sigGeneratorFallbackHtml_() + '</p>'
    ).setTitle('Not available');
  }

  const template = HtmlService.createTemplateFromFile('Index');
  template.actor = actor;
  template.orgLabel = SIG_CONFIG.ORG_LABEL;
  template.generatorUrl = SIG_GENERATOR_URL;
  return template.evaluate()
    .setTitle(SIG_CONFIG.ORG_LABEL + ' Email Signature')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================================
// SERVER FUNCTIONS REACHABLE FROM THE BROWSER
// Every one of these must call requireMember_() first, and must derive the
// account it acts on from that return value alone. google.script.run can invoke
// ANY global function in the project, so an entry point that forgets the gate is
// reachable by any domain user who opens the page source.
// ============================================================================

/**
 * Everything the page needs on load: who the member is according to CAPWATCH,
 * whether a phone number is even in play, where a signature would be written, and
 * the preview itself.
 *
 * @returns {Object}
 */
function apiGetState() {
  const actor = requireMember_();
  const resolved = sigResolveMemberForActor_(actor);
  const member = resolved.member;

  // Best-effort: a member should still be able to SEE their signature when the
  // delegation is misconfigured. Applying it will then fail loudly, which is the
  // right moment for that error to matter.
  let addresses = [];
  let addressError = '';
  try {
    addresses = sigSendAsSnapshot_(actor).orgOwned;
  } catch (err) {
    addressError = err.message;
    Logger.warn('Could not list the caller\'s send-as addresses', {
      user: actor, errorMessage: err.message
    });
  }

  const duty = sigDutyBlock_(member);

  return {
    actor: actor,
    orgLabel: SIG_CONFIG.ORG_LABEL,
    capid: resolved.capid,
    name: sigName_(member),
    dutyLines: duty ? duty.split('<br />') : [],
    duties: sigDutyOptions_(member),
    maxDuties: SIG_MAX_DUTIES,
    /** '' for every cadet, and for any senior with no publishable cell on file. */
    phone: sigFormatPhone_(member.phone),
    isCadet: String(member.type || '').trim().toUpperCase() === 'CADET',
    addresses: addresses,
    addressError: addressError,
    html: sigGenerateEmailSignature_(member)
  };
}

/**
 * Every duty the member holds, as the page offers them: the line each would print,
 * in the order the signature would print them, with the default pick marked.
 *
 * Assistants are included here even though the default pick excludes them — that
 * is the point of offering a choice. They are flagged so the page can label them
 * and so a member is never surprised by which one they ticked.
 *
 * Ordered by the same rules the block uses (echelon, then principal before
 * assistant, then title), so the list a member reads and the signature they get
 * are in the same order. sigDutyBlock_() is the authority for both — this asks it
 * one duty at a time rather than reimplementing the line-writing.
 *
 * @param {Object} member - the resolved CAPWATCH record
 * @returns {Array<Object>}
 */
function sigDutyOptions_(member) {
  // Both arrays: the chooser offers every duty the member holds, and a region or
  // national billet is exactly the kind they most want to show.
  const options = sigAllDuties_(member).map(function (dp) {
    // One duty in isolation renders as its own line, which is what the page shows
    // beside the checkbox — no second implementation of how a line is written.
    return {
      key: sigDutyKey_(dp),
      line: sigDutyBlock_({
        orgName: member.orgName,
        dutyPositions: [dp],
        selectedDutyKeys: [sigDutyKey_(dp)]
      }),
      level: String(dp.level || '').trim().toUpperCase(),
      assistant: !!dp.assistant,
      selected: false
    };
  }).sort(function (a, b) {
    return sigDutyLevelRank_(a.level) - sigDutyLevelRank_(b.level) ||
      (a.assistant ? 1 : 0) - (b.assistant ? 1 : 0) ||
      a.line.localeCompare(b.line);
  });

  // Mark what the member gets if they touch nothing — today's automatic pick,
  // asked of the generator rather than worked out a second time here.
  //
  // Matched line by line and greedily, because a principal and an assistant of the
  // same title at the same unit render the SAME line: marking every match would
  // tick two boxes for one line, and matching before the sort could tick the
  // assistant instead of the principal the default pick actually chose.
  const defaultBlock = sigDutyBlock_(member);
  (defaultBlock ? defaultBlock.split('<br />') : []).forEach(function (line) {
    for (let i = 0; i < options.length; i++) {
      if (!options[i].selected && options[i].line === line) {
        options[i].selected = true;
        return;
      }
    }
  });

  return options;
}

/**
 * Re-renders the preview with the member's current choices.
 *
 * @param {boolean} includePhone
 * @param {Array<string>} [dutyKeys] - duties to show; omit for the default pick
 * @returns {{html: string, includedPhone: boolean, dutyKeys: Array<string>}}
 */
function apiPreview(includePhone, dutyKeys) {
  const actor = requireMember_();
  return sigRender_(actor, includePhone, dutyKeys);
}

/**
 * Writes the signature to the member's own org-owned Send-As identities.
 *
 * Renders from CAPWATCH again rather than trusting anything the preview produced,
 * so what is written is what the server would have shown, not what the client
 * happens to be holding.
 *
 * @param {boolean} includePhone
 * @param {Array<string>} [dutyKeys]
 * @returns {{updated: Array<string>, skipped: Array<string>, failed: Array<Object>, includedPhone: boolean}}
 */
function apiApply(includePhone, dutyKeys) {
  const actor = requireMember_();
  const rendered = sigRender_(actor, includePhone, dutyKeys);

  const result = sigApplyToSendAsIdentities_(actor, rendered.html);

  // Both of these are reported as errors rather than returned as a result the page
  // would render as a success with an empty list of addresses.
  if (!result.updated.length && !result.failed.length) {
    // Nothing to write to at all: every identity on the mailbox is personal.
    throw new Error('None of the send-as addresses on your mailbox belong to ' +
      SIG_CONFIG.ORG_LABEL + ', so there was nothing to update. ' + sigSupportSentence_());
  }
  if (!result.updated.length) {
    throw new Error('Your signature could not be written to ' +
      result.failed.map(function (f) { return f.address; }).join(', ') + '. ' +
      sigSupportSentence_());
  }

  result.includedPhone = rendered.includedPhone;
  return result;
}

/**
 * Renders the caller's signature. The single point where the client's choices are
 * applied, so all the coercion and validation lives here and nowhere else.
 *
 * Anything that is not an explicit `true` suppresses the phone. That is the
 * conservative direction on purpose: a malformed value must not be the reason a
 * personal number gets published, and the page always sends a real boolean.
 *
 * @param {string} actor - from requireMember_()
 * @param {boolean} includePhone
 * @param {Array<string>} [dutyKeys] - see sigValidatedDutyKeys_
 * @returns {{html: string, includedPhone: boolean, dutyKeys: Array<string>}}
 */
function sigRender_(actor, includePhone, dutyKeys) {
  const resolved = sigResolveMemberForActor_(actor);
  const withPhone = includePhone === true;
  const chosen = sigValidatedDutyKeys_(resolved.member, dutyKeys);

  // A shallow copy: these two fields are the whole of what a client can influence,
  // and the record in the cache must not be mutated by a preview toggle.
  const member = Object.assign({}, resolved.member, {
    phone: withPhone ? resolved.member.phone : ''
  });
  if (chosen) member.selectedDutyKeys = chosen;

  return {
    html: sigGenerateEmailSignature_(member),
    includedPhone: !!(withPhone && resolved.member.phone),
    dutyKeys: chosen || []
  };
}

/**
 * The duties a client asked for, checked against the ones the member actually
 * holds — or null for "no selection, use the default pick".
 *
 * WHY EVERY CASE HERE THROWS RATHER THAN TRIMMING
 *
 * A duty key is the one client input that could put WORDS in a signature, so it is
 * the one place a silent correction would be a lie. Quietly dropping an unknown key
 * would publish a signature the member did not ask for and did not see; quietly
 * truncating three duties to two would publish a different one from the preview.
 * The page cannot produce any of these — it enforces the cap and sends keys it was
 * given — so anything that gets here is a bug or a hand-made request, and both
 * deserve to be told.
 *
 * The cap is enforced here and again inside the generator, which is not redundant:
 * this one produces a legible error, that one is the guarantee.
 *
 * @param {Object} member - the resolved CAPWATCH record
 * @param {*} dutyKeys - whatever the client sent
 * @returns {Array<string>|null}
 */
function sigValidatedDutyKeys_(member, dutyKeys) {
  if (dutyKeys === undefined || dutyKeys === null) return null;
  if (!Array.isArray(dutyKeys)) {
    throw new Error('The duty selection could not be read. Reload the page and try again.');
  }
  // An empty array is a real choice — "show no duty at all" — and not the same as
  // sending nothing, which means "pick for me".
  if (!dutyKeys.length) return [];

  if (dutyKeys.length > SIG_MAX_DUTIES) {
    throw new Error('A signature shows at most ' + SIG_MAX_DUTIES +
      ' duty assignments. Pick no more than that.');
  }

  const held = {};
  // Both arrays, so a member may choose the region billet they can see offered —
  // and still nothing outside their own record is selectable.
  sigAllDuties_(member).forEach(function (dp) { held[sigDutyKey_(dp)] = true; });

  const chosen = [];
  dutyKeys.forEach(function (key) {
    const k = String(key == null ? '' : key);
    if (!held[k]) {
      Logger.warn('Rejected a duty that is not on the member\'s CAPWATCH record', {
        capsn: member.capsn, key: k
      });
      throw new Error('One of the duty assignments you picked is not on your CAPWATCH ' +
        'record. Reload the page to see your current assignments.');
    }
    if (chosen.indexOf(k) === -1) chosen.push(k);
  });
  return chosen;
}

/**
 * The "go here instead" line for the pages that refuse.
 *
 * A member turned away by one of them has, without this, been told only what did
 * not work. CAP's generator will give them a correct signature by hand in the
 * meantime, whatever is wrong with their account or this tenant's configuration.
 */
function sigGeneratorFallbackHtml_() {
  return '<br><br>You can build a signature by hand at ' +
    '<a href="' + sigEscape_(SIG_GENERATOR_URL) + '" target="_blank" rel="noopener">' +
    'CAP&rsquo;s own signature generator</a>.';
}

/** HTML-escapes a string for the refusal page, which is assembled by hand. */
function sigEscape_(value) {
  return String(value == null ? '' : value).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}
