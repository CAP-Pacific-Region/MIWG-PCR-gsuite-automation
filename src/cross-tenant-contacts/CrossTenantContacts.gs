/*************************************************************
 * CrossTenantContacts.gs
 * Cross-tenant Domain Shared Contacts sync (sheet-free).
 *
 * Purpose
 *   Make the PEER Workspace tenant's members appear in THIS tenant's Global
 *   Address List, so (e.g.) seniors can email cadets by name and vice-versa.
 *   Replaces the old two-project export->sheet->import->sync pipeline: the
 *   consuming tenant now reads everything from live data it already has.
 *
 * Data sources (NO spreadsheets)
 *   - Roster + attributes for EVERY peer member ............ CAPWATCH (getMembers)
 *     One wing CAPWATCH pull contains both tenants' members (one CAP wing, two
 *     Workspace tenants), so this includes cadet-lite members with no account.
 *   - Authoritative peer Workspace email (+ existence) ..... peer Directory API
 *     via a read-only peer-tenant service account (DWD), same pattern as
 *     getImpersonatedToken_() in UpdateMembers.gs.
 *   - Fallback personal email for no-account members ....... CAPWATCH MbrContact
 *     (already attached to member.email by createMemberObject/addContactInfo,
 *     with the doNotContact flag respected).
 *
 * Email resolution waterfall (per member) — see xtResolvePeerEmail_():
 *   1. peer Workspace primaryEmail by CAPID   (authoritative; fixes collisions/renames)
 *   2. CAPWATCH personal primary email        (cadet-lite / no-account members)
 *   3. do.not.contact+<CAPID>@<selfDomain>    (presence-only sentinel; opt-in)
 *
 * Identical source on every project. Behavior is selected by TENANT_PROFILE
 * (PROFILE_.CROSS_TENANT, version-controlled in config.gs); identity/secrets
 * come from Script Properties (read in getCrossTenantConfig_). Pacific and any
 * tenant with RUN_INBOUND=false no-op.
 *
 * Requires (beyond the shared manifest): the legacy Domain Shared Contacts
 * scope  https://www.google.com/m8/feeds  (added to appsscript.json).
 *
 * Version: 0.1.2 (draft)
 * 0.1.2: dedupe desired set by email — Domain Shared Contacts key on address, so
 *   siblings sharing a parent email churned; collapse to one entry per email.
 *   Also adds update/delete diagnostic sampling in the reconcile.
 * 0.1.1: scope roster to peer types only — getMembers() merges the Manual
 *   Members sheet unconditionally, which leaked non-peer accounts into the set.
 *************************************************************/

// --- tuning (safe defaults; identical on every tenant) ----------------------
const XT_MAX_WRITES_PER_RUN = 3000;               // hard cap on create+update+delete per run
const XT_SOFT_TIME_LIMIT_MS = 25 * 60 * 1000;     // pause before the 30-min trigger ceiling
const XT_WRITE_THROTTLE_EVERY = 50;               // sleep every N writes
const XT_WRITE_THROTTLE_MS = 200;
const XT_MANAGED_HASH_FIELD = 'xtSyncHash';       // gContact:userDefinedField key (stateless diff)

/*************************************************************
 * PUBLIC ENTRY POINTS
 *************************************************************/

/**
 * Main entry: reconcile the peer tenant's members into this tenant's Domain
 * Shared Contacts. Gated by PROFILE_.CROSS_TENANT.RUN_INBOUND.
 */
function syncCrossTenantContacts() {
  const cfg = getCrossTenantConfig_();
  if (!cfg.runInbound) {
    Logger.info('syncCrossTenantContacts skipped (RUN_INBOUND=false for this profile)');
    return;
  }

  Logger.info('Cross-tenant contact sync starting', {
    wing: cfg.wing, selfDomain: cfg.selfDomain, peerDomain: cfg.peerDomain,
    peerTypes: cfg.peerTypes, emitPlaceholders: cfg.emitPlaceholders
  });

  const desired = xtBuildDesiredContacts_(cfg);
  if (desired.list.length === 0) {
    Logger.warn('Cross-tenant sync aborted: 0 desired contacts resolved (guard against mass delete)');
    return;
  }

  const existing = xtListManagedContacts_(cfg, cfg.wing);
  const summary = xtReconcile_(desired, existing, cfg);

  Logger.info('Cross-tenant contact sync complete', Object.assign({
    desired: desired.list.length,
    emailSources: desired.stats
  }, summary));
  return summary;
}

/**
 * Parent/guardian group contacts. Publishes the peer tenant's parent
 * distribution groups (e.g. ca346.parents@<peerDomain>) into this tenant's GAL,
 * so members can email a squadron's parents by name. Source is the peer Groups
 * directory (not CAPWATCH); managed under the cfg.wing + '_PARENTS' marker so it
 * never overlaps the member sync. Gated by PROFILE_.CROSS_TENANT.RUN_PARENTS.
 */
function syncCrossTenantParentContacts() {
  const cfg = getCrossTenantConfig_();
  if (!cfg.runParents) {
    Logger.info('syncCrossTenantParentContacts skipped (RUN_PARENTS=false for this profile)');
    return;
  }

  const marker = cfg.wing + '_PARENTS';
  Logger.info('Cross-tenant parent-group sync starting', {
    marker, selfDomain: cfg.selfDomain, peerDomain: cfg.peerDomain
  });

  const desired = xtBuildDesiredParentContacts_(cfg, marker);
  if (desired.list.length === 0) {
    Logger.warn('Parent-group sync aborted: 0 peer parent groups found (guard against mass delete)');
    return;
  }

  const existing = xtListManagedContacts_(cfg, marker);
  const summary = xtReconcile_(desired, existing, cfg);

  Logger.info('Cross-tenant parent-group sync complete', Object.assign({
    desired: desired.list.length
  }, summary));
  return summary;
}

/**
 * ONE-TIME per project: fill in THIS project's cross-tenant values below and
 * Run once. Writes the XT_PEER_* Script Properties, which survive every `clasp
 * push`. Blank fields are skipped (safe to re-run). Canonical non-secret values
 * are version-controlled in config-tenants/<tenant>.json.
 *
 * The peer service-account PRIVATE KEY is a secret and is intentionally NOT set
 * here — add XT_PEER_SA_KEY by hand in Project Settings > Script Properties, the
 * same way SA_PRIVATE_KEY is handled (see config-tenants/README.md). Store the
 * PEM with literal \n or real newlines; xtPeerToken_ normalizes either.
 */
function setupCrossTenantConfig() {
  const values = {
    XT_PEER_DOMAIN: '',     // peer tenant Workspace domain (seniors proj: cawgcadets.org; cadets proj: cawgcap.org)
    XT_PEER_SA_EMAIL: '',   // peer read-only service account client_email
    XT_PEER_SA_SUBJECT: ''  // a peer-tenant super admin for the SA to impersonate
    // XT_PEER_SA_KEY       // SECRET — set by hand in Script Properties, not here
  };

  const props = PropertiesService.getScriptProperties();
  const applied = [];
  Object.keys(values).forEach(function (k) {
    const v = String(values[k] == null ? '' : values[k]).trim();
    if (v !== '') { props.setProperty(k, v); applied.push(k); }
  });

  console.log('✅ Applied cross-tenant properties: ' + JSON.stringify(applied));
  console.log('ℹ️  Blank fields skipped. Set XT_PEER_SA_KEY by hand (secret).');
  console.log('➡️  Now run validateCrossTenantConfig() to confirm.');
  return applied;
}

/**
 * Validates that this project has the Script Properties the enabled features
 * need. Run after setting properties, before the first sync. Mirrors
 * validateTenantConfig() in config.gs.
 */
function validateCrossTenantConfig() {
  const xt = (PROFILE_ && PROFILE_.CROSS_TENANT) || {};
  if (!xt.RUN_INBOUND && !xt.RUN_PARENTS) {
    console.log('ℹ️  Cross-tenant sync is OFF for this profile (nothing to validate).');
    return { ok: true, missing: [] };
  }
  const p = PropertiesService.getScriptProperties();
  const required = ['XT_PEER_DOMAIN', 'XT_PEER_SA_EMAIL', 'XT_PEER_SA_KEY', 'XT_PEER_SA_SUBJECT'];
  const missing = required.filter(k => {
    const v = p.getProperty(k);
    return v === null || String(v).trim() === '';
  });
  if (missing.length) {
    console.error('❌ Missing cross-tenant Script Properties: ' + missing.join(', '));
  } else {
    console.log('✅ Cross-tenant config OK — peerDomain=' + p.getProperty('XT_PEER_DOMAIN') +
                ', peerSA=' + p.getProperty('XT_PEER_SA_EMAIL'));
  }
  return { ok: missing.length === 0, missing: missing };
}

/*************************************************************
 * CONFIG RESOLUTION (profile behavior + Script-Property identity)
 *************************************************************/

function getCrossTenantConfig_() {
  const xt = (PROFILE_ && PROFILE_.CROSS_TENANT) || {};
  const p = PropertiesService.getScriptProperties();
  const prop = (k) => {
    const v = p.getProperty(k);
    return (v === null || String(v).trim() === '') ? '' : String(v).trim();
  };

  const cfg = {
    // behavior (version-controlled, identical source)
    runInbound: !!xt.RUN_INBOUND,
    runParents: !!xt.RUN_PARENTS,
    peerTypes: Array.isArray(xt.PEER_TYPES) ? xt.PEER_TYPES : [],
    peerLabel: xt.PEER_LABEL || '',              // notes tag, e.g. 'CADET' / 'SENIOR'
    emitPlaceholders: xt.EMIT_PLACEHOLDERS !== false,

    // identity/secrets (per-project Script Properties; never in shared source)
    wing: CONFIG.WING,
    selfDomain: CONFIG.DOMAIN,
    peerDomain: prop('XT_PEER_DOMAIN'),
    peerSa: {
      email: prop('XT_PEER_SA_EMAIL'),
      key: prop('XT_PEER_SA_KEY'),
      subject: prop('XT_PEER_SA_SUBJECT')       // a peer-tenant admin to impersonate
    }
  };

  if (cfg.runInbound) {
    const need = { XT_PEER_DOMAIN: cfg.peerDomain, XT_PEER_SA_EMAIL: cfg.peerSa.email,
                   XT_PEER_SA_KEY: cfg.peerSa.key, XT_PEER_SA_SUBJECT: cfg.peerSa.subject };
    const missing = Object.keys(need).filter(k => !need[k]);
    if (missing.length) {
      throw new Error('Cross-tenant sync enabled but missing Script Properties: ' + missing.join(', '));
    }
    if (!cfg.peerTypes.length) {
      throw new Error('PROFILE_.CROSS_TENANT.PEER_TYPES is empty for this profile');
    }
    if (!cfg.wing || !cfg.selfDomain) {
      throw new Error('CONFIG.WING / CONFIG.DOMAIN unset (run validateTenantConfig())');
    }
  }
  return cfg;
}

/*************************************************************
 * DESIRED STATE — CAPWATCH roster + email waterfall
 *************************************************************/

/**
 * Builds the desired managed-contact set for the peer tenant.
 * @returns {{byCapid: Object, list: Array, stats: Object}}
 */
function xtBuildDesiredContacts_(cfg) {
  // 1) Peer roster from CAPWATCH. getMembers() attaches member.email from
  //    MbrContact (doNotContact-respected) and duty/org attributes. The self
  //    profile's CADET_LITE filter keys off cadet grades only, so fetching the
  //    peer's types never over-filters here (seniors fetching CADET: self
  //    CADET_LITE=false -> all cadets incl. lite; cadets fetching SENIOR: senior
  //    grades never match the cadet exclusion list).
  const roster = getMembers(cfg.peerTypes, true);

  // getMembers() merges the tenant's "Manual Members" sheet UNCONDITIONALLY,
  // after its type filter — so non-peer-type manual accounts (e.g. seniors) can
  // leak into the roster. Keep only members whose type is actually a peer type,
  // or they get published as bogus peer contacts and churn every run.
  const peerTypeSet = {};
  (cfg.peerTypes || []).forEach(t => { peerTypeSet[String(t).trim().toUpperCase()] = true; });

  // 2) Authoritative peer Workspace email (+ existence) by CAPID.
  const peerWsByCapid = xtPeerWorkspaceEmailByCapid_(cfg);

  const byCapid = {};
  const stats = { workspace: 0, capwatch: 0, placeholder: 0, dropped: 0, filteredType: 0, dedupedEmail: 0 };

  Object.keys(roster).forEach(capid => {
    const m = roster[capid];
    if (!peerTypeSet[String(m.type || '').trim().toUpperCase()]) { stats.filteredType++; return; }
    const resolved = xtResolvePeerEmail_(m, cfg, peerWsByCapid);

    if (!resolved.email) { stats.dropped++; return; }         // no email & placeholders off
    stats[resolved.source]++;

    const contact = {
      capid: String(m.capsn),
      givenName: String(m.firstName || '').trim(),
      familyName: String(m.lastName || '').trim(),
      grade: String(m.rank || '').trim(),
      email: resolved.email,
      emailSource: resolved.source,
      dutyTitle: xtDutyTitle_(m),
      department: toTitleCase(m.orgName || m.charter || ''),
      notes: xtNotes_(m, cfg)
    };
    contact.displayName = xtDisplayName_(contact);
    contact.syncHash = xtHash_(contact);

    // CAPID is the stable identity; last write wins on the rare duplicate.
    byCapid[contact.capid] = contact;
  });

  // Domain Shared Contacts effectively key on the email address: Google keeps a
  // single entry when two contacts share an address, so a duplicate email churns
  // forever (the second CAPID never finds its own record). Collapse to one entry
  // per email — siblings sharing a parent's CAPWATCH email keep the lowest CAPID.
  // Workspace and do.not.contact+CAPID addresses are unique, so only shared
  // personal emails are affected.
  const byEmail = {};
  const collisions = [];
  Object.keys(byCapid).sort((a, b) => Number(a) - Number(b)).forEach(capid => {
    const c = byCapid[capid];
    const key = String(c.email).toLowerCase();
    if (byEmail[key]) {
      stats.dedupedEmail++;
      if (collisions.length < 25) collisions.push({ email: key, kept: byEmail[key].capid, dropped: c.capid });
      delete byCapid[capid];          // keep desired.byCapid consistent with the deduped set
      return;
    }
    byEmail[key] = c;
  });
  const list = Object.keys(byEmail).map(k => byEmail[k]);

  Logger.info('Cross-tenant desired set built', { total: list.length, sources: stats });
  if (collisions.length) Logger.info('Cross-tenant email collisions deduped', { sample: collisions });
  return { byCapid, list, stats };
}

/**
 * Email waterfall for one peer member. Pure; no I/O.
 * @returns {{email: string, source: 'workspace'|'capwatch'|'placeholder'}}
 */
function xtResolvePeerEmail_(member, cfg, peerWsByCapid) {
  const capid = String(member.capsn || '').trim();

  // 1) Peer has a Workspace account -> authoritative address.
  const ws = peerWsByCapid[capid];
  if (ws) return { email: ws, source: 'workspace' };

  // 2) No account (cadet-lite < C/SSgt, etc.) -> CAPWATCH personal email.
  //    member.email is already the sanitized MbrContact primary, or null if the
  //    member set doNotContact.
  const personal = sanitizeEmail(member.email);
  if (personal) return { email: personal, source: 'capwatch' };

  // 3) Present in the directory for name/grade/unit, but not emailable.
  if (cfg.emitPlaceholders && capid) {
    return { email: `do.not.contact+${capid}@${cfg.selfDomain}`, source: 'placeholder' };
  }
  return { email: '', source: 'placeholder' };
}

/**
 * Builds the desired parent-group contact set from the peer tenant's
 * distribution groups whose address ends in '.parents@<peerDomain>' (the
 * squadron-groups parent-list naming convention). Keyed by email (groups have
 * no CAPID), tagged with the given marker.
 * @returns {{byCapid: Object, list: Array, stats: Object}}
 */
function xtBuildDesiredParentContacts_(cfg, marker) {
  const groups = xtPeerParentGroups_(cfg);
  const byEmail = {};

  groups.forEach(g => {
    const email = sanitizeEmail(g.email);
    if (!email) return;
    const displayName = String(g.name || email).trim();
    const description = String(g.description || '').trim();
    const contact = {
      capid: '',
      givenName: '',
      familyName: '',
      grade: '',
      email: email,
      orgMarker: marker,
      dutyTitle: 'Parents & Guardians Group',
      department: description,
      notes: description,
      displayName: displayName
    };
    contact.syncHash = xtHash_(contact);
    byEmail[email] = contact;                       // dedupe by address
  });

  const list = Object.keys(byEmail).map(k => byEmail[k]);
  Logger.info('Cross-tenant parent-group desired set built', { total: list.length });
  return { byCapid: {}, list, stats: { groups: list.length } };
}

/*************************************************************
 * PEER DIRECTORY (read-only, via peer-tenant service account / DWD)
 *************************************************************/

/** CAPID -> primaryEmail for the PEER tenant's Workspace accounts. */
function xtPeerWorkspaceEmailByCapid_(cfg) {
  const token = xtPeerToken_('https://www.googleapis.com/auth/admin.directory.user.readonly', cfg);
  const map = {};
  let pageToken = '';

  do {
    const url = 'https://admin.googleapis.com/admin/directory/v1/users' +
      '?customer=my_customer&maxResults=500&projection=full' +
      '&fields=nextPageToken,users(primaryEmail,externalIds,suspended)' +
      (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    if (code < 200 || code >= 300) {
      throw new Error(`Peer directory list failed (${code}): ${resp.getContentText()}`);
    }

    const body = JSON.parse(resp.getContentText() || '{}');
    (body.users || []).forEach(u => {
      if (u.suspended) return;                         // suspended peers get no live email
      const idField = (u.externalIds || []).find(id => id && id.type === 'organization');
      const capid = String((idField && idField.value) || '').trim();
      const email = String(u.primaryEmail || '').trim().toLowerCase();
      if (capid && email) map[capid] = email;
    });

    pageToken = body.nextPageToken || '';
  } while (pageToken);

  Logger.info('Peer Workspace email map built', { count: Object.keys(map).length, peerDomain: cfg.peerDomain });
  return map;
}

/**
 * Lists the peer tenant's parent distribution groups — email localpart ends in
 * '.parents' — via the peer Directory Groups API. Requires the peer SA's DWD
 * grant to include admin.directory.group.readonly.
 * @returns {Array<{email:string, name:string, description:string}>}
 */
function xtPeerParentGroups_(cfg) {
  const token = xtPeerToken_('https://www.googleapis.com/auth/admin.directory.group.readonly', cfg);
  const suffix = '.parents@' + String(cfg.peerDomain).toLowerCase();
  const groups = [];
  let pageToken = '';

  do {
    const url = 'https://admin.googleapis.com/admin/directory/v1/groups' +
      '?customer=my_customer&maxResults=200' +
      '&fields=nextPageToken,groups(email,name,description)' +
      (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    if (code < 200 || code >= 300) {
      throw new Error(`Peer group list failed (${code}): ${resp.getContentText()}`);
    }

    const body = JSON.parse(resp.getContentText() || '{}');
    (body.groups || []).forEach(g => {
      const email = String(g.email || '').trim().toLowerCase();
      if (email.endsWith(suffix)) {
        groups.push({ email: email, name: String(g.name || ''), description: String(g.description || '') });
      }
    });

    pageToken = body.nextPageToken || '';
  } while (pageToken);

  Logger.info('Peer parent groups discovered', { count: groups.length, peerDomain: cfg.peerDomain });
  return groups;
}

/**
 * Mints an OAuth token for the PEER tenant by JWT-signing with the peer-tenant
 * service account and impersonating a peer admin. Mirrors getImpersonatedToken_
 * in UpdateMembers.gs, but reads the XT_PEER_SA_* Script Properties so a project
 * can read its peer's directory without touching its own SA_* credentials.
 */
function xtPeerToken_(scope, cfg) {
  const sa = (cfg && cfg.peerSa) || {};
  if (!sa.email || !sa.key || !sa.subject) {
    throw new Error('Missing peer SA credentials (XT_PEER_SA_EMAIL / XT_PEER_SA_KEY / XT_PEER_SA_SUBJECT).');
  }
  const key = String(sa.key).replace(/\\n/g, '\n');   // stored with literal \n -> real newlines
  const now = Math.floor(Date.now() / 1000);

  const header = { alg: 'RS256', typ: 'JWT' };
  const claimSet = {
    iss: sa.email, sub: sa.subject, aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600, iat: now, scope: scope
  };
  const toSign =
    Utilities.base64EncodeWebSafe(JSON.stringify(header)) + '.' +
    Utilities.base64EncodeWebSafe(JSON.stringify(claimSet));
  const signature = Utilities.computeRsaSha256Signature(toSign, key);
  const jwt = `${toSign}.${Utilities.base64EncodeWebSafe(signature)}`;

  const resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: jwt },
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`Peer token exchange failed (${code}): ${resp.getContentText()}`);
  }
  return JSON.parse(resp.getContentText()).access_token;
}

/*************************************************************
 * EXISTING STATE — this tenant's managed Domain Shared Contacts
 *************************************************************/

/**
 * Lists managed shared contacts in THIS tenant carrying the given orgName marker
 * (member sync uses cfg.wing; parents use cfg.wing + '_PARENTS'). Stateless: the
 * per-contact hash is read back from the xtSyncHash userDefinedField, so no
 * sync-state sheet is needed.
 */
function xtListManagedContacts_(cfg, marker) {
  const token = ScriptApp.getOAuthToken();       // this tenant's own contacts scope
  let url = `https://www.google.com/m8/feeds/contacts/${encodeURIComponent(cfg.selfDomain)}/full` +
            `?alt=json&max-results=20000`;
  const managed = [];

  while (url) {
    const resp = xtFetch_(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token, 'GData-Version': '3.0' }
    }, 'list');
    if (resp.getResponseCode() !== 200) {
      throw new Error('Managed shared-contact list failed: ' + resp.getContentText());
    }

    const feed = (JSON.parse(resp.getContentText()) || {}).feed || {};
    (feed.entry || []).forEach(e => {
      const org = (e['gd$organization'] || [])[0] || {};
      const orgName = (org['gd$orgName'] && org['gd$orgName'].$t) || '';
      if (orgName !== marker) return;              // only contacts we manage under this marker

      const udf = (e['gContact$userDefinedField'] || [])
        .find(f => f && f.key === XT_MANAGED_HASH_FIELD);
      const editLink = (e.link || []).find(l => l.rel === 'edit');

      managed.push({
        resourceId: (e.id && e.id.$t) ? String(e.id.$t).replace('http://', 'https://') : '',
        editHref: editLink ? String(editLink.href).replace('http://', 'https://') : '',
        etag: e['gd$etag'] || '*',
        capid: String((e['gd$externalId'] && e['gd$externalId'][0] && e['gd$externalId'][0].value) || '').trim(),
        email: String((e['gd$email'] && e['gd$email'][0] && e['gd$email'][0].address) || '').trim().toLowerCase(),
        syncHash: udf ? String(udf.value || '') : ''
      });
    });

    const next = (feed.link || []).find(l => String(l.rel || '').toLowerCase() === 'next');
    url = next && next.href ? String(next.href) : '';
  }

  Logger.info('Managed shared contacts listed', { count: managed.length });
  return managed;
}

/*************************************************************
 * RECONCILE (create / update / delete)
 *************************************************************/

function xtReconcile_(desired, existing, cfg) {
  const token = ScriptApp.getOAuthToken();
  const feedBase = `https://www.google.com/m8/feeds/contacts/${encodeURIComponent(cfg.selfDomain)}/full`;

  // Index existing managed contacts by CAPID (stable), then by email (legacy).
  const existingByCapid = {};
  const existingByEmail = {};
  existing.forEach(c => {
    if (c.capid) existingByCapid[c.capid] = c;
    else if (c.email) existingByEmail[c.email] = c;
  });

  const ops = [];
  const matched = {};
  let creates = 0, updates = 0, unchanged = 0, deletes = 0;
  const updateSample = [];   // diagnostics: which contacts churn and why
  const deleteSample = [];

  desired.list.forEach(want => {
    const byCapid = existingByCapid[want.capid];
    const cur = byCapid || existingByEmail[want.email];
    if (!cur) {
      creates++;
      ops.push(() => xtCreateContact_(feedBase, token, want, cfg));
      return;
    }
    matched[cur.resourceId] = true;

    if (cur.syncHash && cur.syncHash === want.syncHash) { unchanged++; return; }
    updates++;
    if (updateSample.length < 25) {
      updateSample.push({
        capid: want.capid, email: want.email, source: want.emailSource,
        matchedBy: byCapid ? 'capid' : 'email',
        storedHash: cur.syncHash ? (cur.syncHash === want.syncHash ? 'same' : 'differs') : 'EMPTY'
      });
    }
    ops.push(() => xtUpdateContact_(cur, token, want, cfg));
  });

  // Delete managed contacts no longer desired.
  existing.forEach(cur => {
    if (matched[cur.resourceId]) return;
    if (cur.capid && desired.byCapid[cur.capid]) return;
    deletes++;
    if (deleteSample.length < 25) deleteSample.push({ capid: cur.capid, email: cur.email });
    ops.push(() => xtDeleteContact_(cur, token));
  });

  Logger.info('Cross-tenant reconcile plan', {
    creates, updates, deletes, unchanged, totalWrites: ops.length
  });
  if (updateSample.length) Logger.info('Cross-tenant update sample', { sample: updateSample });
  if (deleteSample.length) Logger.info('Cross-tenant delete sample', { sample: deleteSample });

  return xtRunWriteOps_(ops);
}

function xtRunWriteOps_(ops) {
  const start = new Date();
  const cap = Math.min(ops.length, XT_MAX_WRITES_PER_RUN);
  let ok = 0, failed = 0, processed = 0;

  for (let i = 0; i < cap; i++) {
    if (new Date() - start >= XT_SOFT_TIME_LIMIT_MS) {
      Logger.warn('Cross-tenant write loop paused before time limit', { processed, total: ops.length });
      break;
    }
    const r = ops[i]();
    processed++;
    if (r && r.ok) ok++;
    else { failed++; Logger.warn('Cross-tenant write failed', { action: r && r.action, email: r && r.email, code: r && r.code }); }

    if (XT_WRITE_THROTTLE_MS > 0 && (i + 1) % XT_WRITE_THROTTLE_EVERY === 0) Utilities.sleep(XT_WRITE_THROTTLE_MS);
  }

  return {
    totalWrites: ops.length, processed, ok, failed,
    remaining: Math.max(ops.length - processed, 0),
    completedAll: processed >= ops.length,
    durationMs: new Date() - start
  };
}

/*************************************************************
 * CONTACT WRITES (M8 GData feed, XML)
 *************************************************************/

function xtCreateContact_(feedBase, token, contact, cfg) {
  const resp = xtFetch_(feedBase, {
    method: 'post',
    contentType: 'application/atom+xml',
    payload: xtBuildContactXml_(contact, cfg),
    headers: { Authorization: 'Bearer ' + token, 'GData-Version': '3.0' }
  }, 'create');
  const code = resp.getResponseCode();
  const ok = code >= 200 && code < 300;
  return { ok, action: 'create', email: contact.email, code, body: ok ? '' : resp.getContentText() };
}

function xtUpdateContact_(current, token, contact, cfg) {
  const target = current.editHref || current.resourceId;
  const resp = xtFetch_(target, {
    method: 'put',
    contentType: 'application/atom+xml',
    payload: xtBuildContactXml_(contact, cfg),
    headers: { Authorization: 'Bearer ' + token, 'GData-Version': '3.0', 'If-Match': current.etag || '*' }
  }, 'update');
  let code = resp.getResponseCode();
  // If the row vanished between list and write, fall back to create.
  if (code === 404) return xtCreateContact_(
    `https://www.google.com/m8/feeds/contacts/${encodeURIComponent(cfg.selfDomain)}/full`, token, contact, cfg);
  const ok = code >= 200 && code < 300;
  return { ok, action: 'update', email: contact.email, code, body: ok ? '' : resp.getContentText() };
}

function xtDeleteContact_(current, token) {
  const target = current.editHref || current.resourceId;
  const resp = xtFetch_(target, {
    method: 'delete',
    headers: { Authorization: 'Bearer ' + token, 'GData-Version': '3.0', 'If-Match': current.etag || '*' }
  }, 'delete');
  const code = resp.getResponseCode();
  const ok = (code >= 200 && code < 300) || code === 404;
  return { ok, action: 'delete', email: current.email, code };
}

/** Builds an Atom entry. orgName (c.orgMarker, default cfg.wing) marks the
 *  contact as managed; the content hash rides along in a userDefinedField for
 *  stateless diffing. */
function xtBuildContactXml_(c, cfg) {
  const full = xtEsc_(c.displayName);
  const marker = c.orgMarker || cfg.wing;
  const externalId = c.capid
    ? `<gd:externalId label="CAPID" rel="http://schemas.google.com/g/2005#organization" value="${xtEsc_(c.capid)}"/>`
    : '';

  return `
    <entry xmlns="http://www.w3.org/2005/Atom"
           xmlns:gd="http://schemas.google.com/g/2005"
           xmlns:gContact="http://schemas.google.com/contact/2008">
      <category scheme="http://schemas.google.com/g/2005#kind"
                term="http://schemas.google.com/contact/2008#contact"/>
      <title>${full}</title>
      <gd:name>
        <gd:fullName>${full}</gd:fullName>
        <gd:givenName>${xtEsc_(c.givenName)}</gd:givenName>
        <gd:familyName>${xtEsc_(c.familyName)}</gd:familyName>
      </gd:name>
      <gd:email rel="http://schemas.google.com/g/2005#work" primary="true"
                address="${xtEsc_(c.email)}"/>
      <gd:organization rel="http://schemas.google.com/g/2005#work" primary="true">
        <gd:orgName>${xtEsc_(marker)}</gd:orgName>
        <gd:orgTitle>${xtEsc_(c.dutyTitle || '')}</gd:orgTitle>
        <gd:orgDepartment>${xtEsc_(c.department || '')}</gd:orgDepartment>
      </gd:organization>
      ${externalId}
      <gContact:userDefinedField key="${XT_MANAGED_HASH_FIELD}" value="${xtEsc_(c.syncHash)}"/>
      <content type="text">${xtEsc_(c.notes || '')}</content>
    </entry>
  `;
}

/*************************************************************
 * SMALL HELPERS (xt-prefixed to avoid the shared global namespace)
 *************************************************************/

/** "Last, First Grade" — matches the display convention used elsewhere. */
function xtDisplayName_(c) {
  return [c.familyName, ', ', c.givenName, c.grade ? ' ' + c.grade : ''].join('').trim();
}

/** Duty title identical to what provisioning writes to organizations.title. */
function xtDutyTitle_(m) {
  const duties = Array.isArray(m.dutyPositions) ? m.dutyPositions : [];
  if (!duties.length) return 'No Duty Assignment';
  return duties.map(d => d.assistant ? d.id + ' (A)' : d.id).join(', ');
}

/** Notes: "<PEER_LABEL>, <charter>" e.g. "CADET, PCR-CA-057". */
function xtNotes_(m, cfg) {
  return [cfg.peerLabel || m.type || '', m.charter || ''].filter(Boolean).join(', ');
}

/** Stable content hash -> stored in the contact, compared next run. */
function xtHash_(c) {
  const raw = JSON.stringify([
    String(c.email || '').toLowerCase(),
    c.displayName, c.givenName, c.familyName, c.grade,
    c.dutyTitle, c.department, c.notes, c.capid
  ]);
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return digest.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

/** M8 fetch with backoff on 429/5xx. */
function xtFetch_(url, options, label) {
  let delay = 250;
  for (let attempt = 1; attempt <= 6; attempt++) {
    const resp = UrlFetchApp.fetch(url, Object.assign({ muteHttpExceptions: true }, options || {}));
    const code = resp.getResponseCode();
    if (code < 429) return resp;                     // success or non-retryable client error
    if (code === 429 || (code >= 500 && code < 600)) {
      Logger.warn('xtFetch retry', { label, attempt, code, waitMs: delay });
      Utilities.sleep(delay);
      delay = Math.min(delay * 2, 8000);
      continue;
    }
    return resp;
  }
  return UrlFetchApp.fetch(url, Object.assign({ muteHttpExceptions: true }, options || {}));
}

function xtEsc_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
