/**
 * Cadet → senior transition: personal Contacts.
 *
 * Runs on the CADETS tenant (TRANSITION_CONFIG.ROLE === 'source'). See
 * CadetTransition.gs for the lifecycle.
 *
 * The member's personal contacts — their own address book, distinct from the
 * Domain Shared Contacts the cross-tenant module syncs — die with the account
 * when it is deleted, exactly like Drive. So they are copied first, and
 * closeCompletedTransitions() refuses until ContactsMigrated is set.
 *
 * Simpler than Drive: contacts are a flat list, no folder tree, and People API's
 * batchCreateContacts writes up to 200 in a single call, so a typical address
 * book of a few dozen finishes in one request.
 *
 * SCOPES (SA domain-wide delegation grants):
 *   cadets local SA:  https://www.googleapis.com/auth/contacts   (read the source)
 *   seniors peer SA:  https://www.googleapis.com/auth/contacts   (create in dest)
 *
 * Resumability, like Drive, lives in the DESTINATION: each created contact
 * carries a userDefined field naming the source contact it came from, so a run
 * begins by asking the destination what is already there. A lost cursor cannot
 * duplicate — a re-list simply skips what is already marked.
 */

/** Read+write personal contacts. */
const CONTACTS_SCOPE_ = 'https://www.googleapis.com/auth/contacts';

/** userDefined key stamped on every created contact, naming its source. */
const CONTACTS_MARK_KEY_ = 'xtSrcContact';

/** Contacts per batchCreateContacts call. API max is 200; 100 keeps requests small. */
const CONTACTS_BATCH_ = 100;

/** Source fields worth carrying across. Read-only sub-metadata is stripped. */
const CONTACTS_PERSON_FIELDS_ = [
  'names', 'nicknames', 'emailAddresses', 'phoneNumbers', 'organizations',
  'addresses', 'biographies', 'urls', 'birthdays', 'events', 'relations',
  'occupations', 'imClients', 'userDefined'
];

/** Continuation trigger + scope property. Distinct from Gmail and Drive. */
const CONTACTS_CONTINUATION_FN_ = 'continueCadetTransitionContactsMigration';
const CONTACTS_CONTINUATION_SCOPE_PROP_ = 'CONTACTS_CONTINUATION_SCOPE';

// ============================================================================
// ENTRY POINTS
// ============================================================================

/**
 * Copies one member's personal contacts to their senior account.
 *
 * @param {string} capid
 * @returns {{copied: number, complete: boolean}}
 */
function migrateSingleTransitionContacts(capid) {
  if (TRANSITION_CONFIG.ROLE !== 'source') throw new Error('Not the source tenant');
  return withTransitionLock_(() => migrateSingleTransitionContacts_(capid),
    { copied: 0, complete: false });
}

function migrateSingleTransitionContacts_(capid) {
  const rows = readTransitions_();
  const row = rows[String(capid)];
  if (!row) throw new Error('No transition row for CAPID ' + capid);
  if (!row.SeniorEmail) throw new Error('No destination account for CAPID ' + capid);

  return migrateOneContacts_(row, new Date());
}

/**
 * Copies contacts for every migrated member who has not had it done yet.
 *
 * @returns {{processed: number}}
 */
function migrateAllTransitionContacts() {
  if (TRANSITION_CONFIG.ROLE !== 'source') {
    Logger.info('Contacts migration skipped — not the source tenant');
    return { processed: 0 };
  }
  return withTransitionLock_(migrateAllTransitionContacts_, { processed: 0 });
}

function migrateAllTransitionContacts_() {
  const started = new Date();
  const rows = readTransitions_();
  let processed = 0;

  for (const capid in rows) {
    const row = rows[capid];
    if (row.MigrationStatus !== TRANSITION_CONFIG.STATUS.COMPLETE) continue;
    if (String(row.ContactsMigrated || '') !== '') continue;
    if (!row.SeniorEmail) continue;

    const result = migrateOneContacts_(row, started);
    processed++;
    if (!result.complete) break;   // migrateOneContacts_ scheduled its own continuation
  }

  Logger.info('Contacts migration pass completed', {
    duration: new Date() - started + 'ms', processed: processed
  });
  return { processed: processed };
}

/** Continuation trigger target. Drains the queue across executions. */
function continueCadetTransitionContactsMigration(e) {
  if (e && e.triggerUid) {
    ScriptApp.getProjectTriggers()
      .filter(t => t.getUniqueId() === e.triggerUid)
      .forEach(t => ScriptApp.deleteTrigger(t));
  }

  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(CONTACTS_CONTINUATION_SCOPE_PROP_);
  props.deleteProperty(CONTACTS_CONTINUATION_SCOPE_PROP_);

  let scope = null;
  try { scope = raw ? JSON.parse(raw) : null; } catch (err) {}

  if (scope && scope.capid) {
    const result = migrateSingleTransitionContacts(scope.capid);
    if (result && result.complete) scheduleContactsContinuation_({});   // sweep for more
  } else {
    migrateAllTransitionContacts();
  }
}

// ============================================================================
// PER-MEMBER COPY
// ============================================================================

/**
 * Copies one member's contacts, resuming from the destination's record.
 *
 * @param {Object} row
 * @param {Date} started
 * @returns {{copied: number, complete: boolean}}
 */
function migrateOneContacts_(row, started) {
  const cfg = getCrossTenantConfig_();
  const srcToken = getImpersonatedToken_(row.CadetEmail, CONTACTS_SCOPE_);
  const destToken = xtPeerToken_(CONTACTS_SCOPE_, cfg, row.SeniorEmail);

  Logger.info('Contacts migration starting', {
    capid: row.CAPID, from: row.CadetEmail, to: row.SeniorEmail
  });

  const done = loadContactsProgress_(destToken);   // source ids already in the dest
  let pageToken = row.ContactsCursor || '';
  let copied = Number(row.ContactsMigrated) || 0;
  let sawAny = false;

  do {
    if (new Date() - started > MIGRATE_SOFT_TIME_LIMIT_MS_) {
      setTransitionField_(row._rowNumber, 'ContactsCursor', pageToken);
      setTransitionField_(row._rowNumber, 'ContactsMigrated', copied);
      scheduleContactsContinuation_({ capid: String(row.CAPID) });
      Logger.info('Contacts migration paused — will resume', { capid: row.CAPID, copied: copied });
      return { copied: copied, complete: false };
    }

    const page = listSourceContacts_(srcToken, pageToken);
    sawAny = sawAny || (page.connections.length > 0);

    // Only contacts not already in the destination, and only those with at least
    // a name or an email — People rejects an empty contact.
    const todo = page.connections.filter(p =>
      !done.has(p.resourceName) && personHasContent_(p)
    );

    for (const group of chunk_(todo, CONTACTS_BATCH_)) {
      const created = batchCreateContacts_(destToken, group);
      created.forEach(srcId => done.add(srcId));
      copied += created.length;
    }

    pageToken = page.nextPageToken || '';
    setTransitionField_(row._rowNumber, 'ContactsCursor', pageToken);
    setTransitionField_(row._rowNumber, 'ContactsMigrated', copied);

  } while (pageToken);

  setTransitionField_(row._rowNumber, 'ContactsCursor', '');
  Logger.info('Contacts migration complete', { capid: row.CAPID, copied: copied });
  return { copied: copied, complete: true };
}

/**
 * Source contact ids already present in the destination, from the userDefined
 * marker. This is the resume mechanism — the destination is authoritative.
 *
 * @param {string} destToken
 * @returns {Set<string>}
 */
function loadContactsProgress_(destToken) {
  const done = new Set();
  let pageToken = '';

  do {
    const url = 'https://people.googleapis.com/v1/people/me/connections' +
      '?personFields=userDefined&pageSize=1000' +
      (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

    const body = contactsFetch_(url, { method: 'get', token: destToken }, 'load contacts progress');

    (body.connections || []).forEach(p => {
      (p.userDefined || []).forEach(u => {
        if (u.key === CONTACTS_MARK_KEY_ && u.value) done.add(u.value);
      });
    });

    pageToken = body.nextPageToken || '';
  } while (pageToken);

  Logger.info('Contacts progress loaded', { alreadyInDest: done.size });
  return done;
}

/**
 * One page of the source's connections.
 *
 * @param {string} srcToken
 * @param {string} pageToken
 * @returns {{connections: Array<Object>, nextPageToken: string}}
 */
function listSourceContacts_(srcToken, pageToken) {
  const url = 'https://people.googleapis.com/v1/people/me/connections' +
    '?personFields=' + encodeURIComponent(CONTACTS_PERSON_FIELDS_.join(',')) +
    '&pageSize=200' +
    (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

  const body = contactsFetch_(url, { method: 'get', token: srcToken }, 'list source contacts');
  return { connections: body.connections || [], nextPageToken: body.nextPageToken || '' };
}

/**
 * Creates a batch of contacts in the destination, each marked with its source id.
 *
 * @param {string} destToken
 * @param {Array<Object>} sourcePeople
 * @returns {Array<string>} source resourceNames successfully created
 */
function batchCreateContacts_(destToken, sourcePeople) {
  const body = contactsFetch_(
    'https://people.googleapis.com/v1/people:batchCreateContacts',
    {
      method: 'post', token: destToken,
      payload: JSON.stringify({
        contacts: sourcePeople.map(p => ({ contactPerson: cleanPersonForCreate_(p) })),
        readMask: 'names'
      })
    },
    'batch create contacts'
  );

  // The response order matches the request order, so map results back to the
  // source ids we sent — that is what gets recorded as done.
  const results = body.createdPeople || [];
  const out = [];
  results.forEach((_, i) => { if (sourcePeople[i]) out.push(sourcePeople[i].resourceName); });
  return out;
}

/**
 * Strips a source person down to writable fields and stamps the source marker.
 *
 * People API read responses carry per-field `metadata` and top-level
 * resourceName/etag that a create call rejects. This copies only the value
 * fields and drops metadata, then adds the userDefined marker so the copy can be
 * recognised on a later run.
 *
 * @param {Object} person
 * @returns {Object} contactPerson for batchCreateContacts
 */
function cleanPersonForCreate_(person) {
  const out = {};

  CONTACTS_PERSON_FIELDS_.forEach(field => {
    if (field === 'userDefined') return;   // handled below
    if (!Array.isArray(person[field]) || !person[field].length) return;
    out[field] = person[field].map(item => {
      const clean = {};
      Object.keys(item).forEach(k => { if (k !== 'metadata') clean[k] = item[k]; });
      return clean;
    });
  });

  // Carry the source's own userDefined entries, then append our marker.
  const userDefined = (person.userDefined || [])
    .map(u => ({ key: u.key, value: u.value }))
    .filter(u => u.key !== CONTACTS_MARK_KEY_);
  userDefined.push({ key: CONTACTS_MARK_KEY_, value: person.resourceName });
  out.userDefined = userDefined;

  return out;
}

/**
 * True if a person has enough to be worth creating (People rejects the empty).
 *
 * @param {Object} person
 * @returns {boolean}
 */
function personHasContent_(person) {
  return !!((person.names && person.names.length) ||
            (person.emailAddresses && person.emailAddresses.length) ||
            (person.phoneNumbers && person.phoneNumbers.length) ||
            (person.organizations && person.organizations.length));
}

// ============================================================================
// HTTP
// ============================================================================

/**
 * One People API call with retry. Mirrors gmailFetch_/driveFetch_ — the shared
 * executeWithRetry() sleeps before every call and misreads raw UrlFetchApp
 * failures as non-transient.
 *
 * @param {string} url
 * @param {{method: string, token: string, payload: string=}} opts
 * @param {string} what
 * @returns {Object}
 */
function contactsFetch_(url, opts, what) {
  const params = {
    method: opts.method,
    headers: { Authorization: 'Bearer ' + opts.token },
    muteHttpExceptions: true
  };
  if (opts.payload) {
    params.contentType = 'application/json';
    params.payload = opts.payload;
  }

  let lastError = '';
  for (let attempt = 1; attempt <= MIGRATE_MAX_ATTEMPTS_; attempt++) {
    const resp = UrlFetchApp.fetch(url, params);
    const code = resp.getResponseCode();
    const text = resp.getContentText();

    if (code >= 200 && code < 300) return text ? JSON.parse(text) : {};

    lastError = `People ${what} failed (${code}): ${text}`;
    const transient = code === 429 || (code >= 500 && code < 600);
    if (!transient || attempt === MIGRATE_MAX_ATTEMPTS_) throw new Error(lastError);
    Utilities.sleep(MIGRATE_BACKOFF_BASE_MS_ * Math.pow(2, attempt - 1));
  }
  throw new Error(lastError);
}

// ============================================================================
// CONTINUATION
// ============================================================================

/** Schedules a one-shot Contacts continuation, guarded against stacking. */
function scheduleContactsContinuation_(scope) {
  const already = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === CONTACTS_CONTINUATION_FN_);
  if (already) return;

  PropertiesService.getScriptProperties().setProperty(
    CONTACTS_CONTINUATION_SCOPE_PROP_, JSON.stringify(scope || {})
  );
  ScriptApp.newTrigger(CONTACTS_CONTINUATION_FN_).timeBased().after(60 * 1000).create();
  Logger.info('Contacts continuation scheduled', { scope: scope || {} });
}

/** Removes pending Contacts continuations. */
function clearContactsContinuations() {
  let removed = 0;
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === CONTACTS_CONTINUATION_FN_)
    .forEach(t => { ScriptApp.deleteTrigger(t); removed++; });
  return { removed: removed };
}

// ============================================================================
// PREVIEW
// ============================================================================

/**
 * Proves both credentials and counts contacts, without copying. Read-only.
 *
 * @param {string} capid
 */
function previewSingleTransitionContacts(capid) {
  const rows = readTransitions_();
  const row = rows[String(capid)];
  if (!row) { console.log('No row for ' + capid); return; }
  if (!row.SeniorEmail) { console.log(row.Name + ': no destination account'); return; }

  const cfg = getCrossTenantConfig_();
  try {
    const srcToken = getImpersonatedToken_(row.CadetEmail, CONTACTS_SCOPE_);
    let count = 0, pageToken = '';
    do {
      const page = listSourceContacts_(srcToken, pageToken);
      count += page.connections.filter(personHasContent_).length;
      pageToken = page.nextPageToken;
    } while (pageToken);

    const destToken = xtPeerToken_(CONTACTS_SCOPE_, cfg, row.SeniorEmail);
    const already = loadContactsProgress_(destToken);

    console.log(`${row.Name}: ${count} source contacts, ${already.size} already in destination`);
    console.log('Both credentials work. Nothing was copied.');
  } catch (e) {
    console.log('FAILED: ' + (e && e.message ? e.message : String(e)));
  }
}
