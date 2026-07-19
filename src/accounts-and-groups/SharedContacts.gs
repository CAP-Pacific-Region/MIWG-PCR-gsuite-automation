/**
 * SharedContacts.gs
 * Domain Shared Contacts sync for "External Contacts" tab.
 *
 * Only public entry point:
 *   syncExternalContactsToDomainSharedContacts_()
 *
 * Version: 1.0.0
 * Date: 2026-07-09
 * Changes: Folded into the shared src/ from the Pacific Region project; gated
 *   behind PROFILE_.RUN_SHARED_CONTACTS (on for pacific, off for wing — the
 *   wing manages shared contacts in a separate project). See PCR_CHANGELOG.md.
 */

/**
 * Syncs contacts from the Automation Spreadsheet tab "External Contacts" into
 * the tenant's Domain Shared Contacts.
 *
 * Sheet tab: External Contacts
 * Columns:
 *   CAPID, LastName, FirstName, MiddleName, Suffix, OrgID, Rank, DutyID,
 *   Assistant, Type, Status, Email, Primary Phone
 *
 * Notes:
 * - Uses the Domain Shared Contacts (legacy M8 feed): https://www.google.com/m8/feeds/contacts/DOMAIN/PROJECTION
 * - Matches existing shared contacts by email address (primary email on the entry).
 * - Upserts (create if missing, update if exists).
 */
function syncExternalContactsToDomainSharedContacts_() {
  if (!PROFILE_.RUN_SHARED_CONTACTS) {
    Logger.info('syncExternalContactsToDomainSharedContacts skipped (RUN_SHARED_CONTACTS=false for this tenant profile)');
    return;
  }
  Logger.info('Starting syncExternalContactsToDomainSharedContacts');

  const contacts = listExternalContactsFromSheet_();
  Logger.info('External contacts loaded', { count: contacts.length });

  let created = 0;
  let updated = 0;
  let skipped = 0;
  let failed = 0;

  contacts.forEach((c, idx) => {
    try {
      if (!c.email) {
        skipped++;
        return;
      }

      const result = upsertDomainSharedContactByEmail_(c);
      if (result === 'created') created++;
      else if (result === 'updated') updated++;
      else skipped++;

      // light rate limit protection
      if (idx > 0 && idx % 25 === 0) Utilities.sleep(250);

    } catch (e) {
      failed++;
      Logger.error('External contact upsert failed', {
        capid: c.capid,
        email: c.email,
        errorMessage: e && e.message ? e.message : String(e)
      });
    }
  });

  Logger.info('syncExternalContactsToDomainSharedContacts completed', {
    total: contacts.length,
    created: created,
    updated: updated,
    skipped: skipped,
    failed: failed
  });
}

/**
 * Manual runner / public entrypoint.
 * Apps Script UI sometimes hides functions that look "private" (ending in underscore).
 */
function runExternalContactsToDomainSharedContacts() {
  return syncExternalContactsToDomainSharedContacts_();
}

/** Reads and normalizes rows from the "External Contacts" sheet */
function listExternalContactsFromSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('External Contacts');
  if (!sheet) {
    Logger.error('External Contacts sheet not found', {
      spreadsheetId: CONFIG.AUTOMATION_SPREADSHEET_ID,
      expectedTab: 'External Contacts'
    });
    return [];
  }

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  const header = values[0].map(h => String(h || '').trim());
  const idx = (name) => header.indexOf(name);

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    const capid = String(row[idx('CAPID')] || '').trim();
    const lastName = String(row[idx('LastName')] || '').trim();
    const firstName = String(row[idx('FirstName')] || '').trim();
    const middleName = String(row[idx('MiddleName')] || '').trim();
    const suffix = String(row[idx('Suffix')] || '').trim();
    const orgid = String(row[idx('OrgID')] || '').trim();
    const rank = String(row[idx('Rank')] || '').trim();
    const dutyId = String(row[idx('DutyID')] || '').trim();
    const isAssistant = String(row[idx('Assistant')] || '').trim() === '1';
    const type = String(row[idx('Type')] || '').trim();
    const status = String(row[idx('Status')] || '').trim();
    const email = sanitizeEmail(String(row[idx('Email')] || '').trim());
    const phoneRaw = String(row[idx('Primary Phone')] || '').trim();

    // Normalize phone to E.164 +1XXXXXXXXXX if possible
    let phone = '';
    if (phoneRaw) {
      const digits = phoneRaw.replace(/\D/g, '');
      if (digits.length >= 10) phone = `+1${digits.slice(-10)}`;
    }

    // Skip empty rows
    if (!email && !lastName && !firstName) continue;

    const displayName = [
      lastName + (suffix ? ' ' + suffix : ''),
      ', ',
      firstName,
      middleName ? ' ' + middleName.charAt(0) : '',
      rank ? ' ' + rank : ''
    ].join('').trim();

    // Build a Directory-like "title" for quick searching
    const dutyTitle = dutyId
      ? (isAssistant ? `${dutyId} (A)` : `${dutyId} (P)`)
      : 'External Contact';

    out.push({
      capid: capid,
      firstName: firstName,
      lastName: lastName,
      middleName: middleName,
      suffix: suffix,
      orgid: orgid,
      rank: rank,
      dutyId: dutyId,
      assistant: isAssistant,
      type: type,
      status: status,
      email: email,
      phone: phone || '',
      displayName: displayName,
      dutyTitle: dutyTitle
    });
  }

  return out;
}

/**
 * Upserts a Domain Shared Contact by matching the email address.
 * Returns: 'created' | 'updated' | 'skipped'
 */
function upsertDomainSharedContactByEmail_(contact) {
  const token = ScriptApp.getOAuthToken();
  const domain = CONFIG.DOMAIN;
  const feedBase = `https://www.google.com/m8/feeds/contacts/${encodeURIComponent(domain)}/full`;

  // 1) Find existing entry by email
  const findUrl = `${feedBase}?q=${encodeURIComponent(contact.email)}&max-results=1&alt=json`;
  const findResp = UrlFetchApp.fetch(findUrl, {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token,
      'GData-Version': '3.0'
    },
    muteHttpExceptions: true
  });

  const findCode = findResp.getResponseCode();
  const findBody = findResp.getContentText();
  if (findCode < 200 || findCode >= 300) {
    throw new Error(`Shared contact search failed (${findCode}): ${findBody}`);
  }

  const found = JSON.parse(findBody);
  const entries = (found.feed && found.feed.entry) ? found.feed.entry : [];

  if (!entries.length) {
    // 2) Create
    const createXml = buildSharedContactEntryXml_(contact);
    const createResp = UrlFetchApp.fetch(feedBase, {
      method: 'post',
      contentType: 'application/atom+xml; charset=UTF-8',
      payload: createXml,
      headers: {
        Authorization: 'Bearer ' + token,
        'GData-Version': '3.0',
        Accept: 'application/atom+xml'
      },
      muteHttpExceptions: true
    });

    const code = createResp.getResponseCode();
    const body = createResp.getContentText();
    if (code < 200 || code >= 300) {
      throw new Error(`Shared contact create failed (${code}): ${body}`);
    }

    return 'created';
  }

  // 3) Update (PUT to edit link)
  const entry = entries[0];
  const links = entry.link || [];
  const editLink = links.find(l => l.rel === 'edit');
  const editHref = editLink ? editLink.href : '';

  if (!editHref) {
    // Can't update without edit URL; skip
    Logger.warn('Shared contact found but missing edit link (skipping update)', {
      email: contact.email
    });
    return 'skipped';
  }

  // Need the current ETag for update. In JSON feed it's in gd$etag.
  const etag = entry['gd$etag'] || '*';

  const updateXml = buildSharedContactEntryXml_(contact, entry);
  const updateResp = UrlFetchApp.fetch(editHref, {
    method: 'put',
    contentType: 'application/atom+xml; charset=UTF-8',
    payload: updateXml,
    headers: {
      Authorization: 'Bearer ' + token,
      'GData-Version': '3.0',
      'If-Match': etag,
      Accept: 'application/atom+xml'
    },
    muteHttpExceptions: true
  });

  const ucode = updateResp.getResponseCode();
  const ubody = updateResp.getContentText();
  if (ucode < 200 || ucode >= 300) {
    throw new Error(`Shared contact update failed (${ucode}): ${ubody}`);
  }

  return 'updated';
}

/**
 * Builds an Atom XML entry for a Domain Shared Contact.
 * If `existingJsonEntry` is provided, preserves the contact id/updated fields.
 */
function buildSharedContactEntryXml_(c, existingJsonEntry) {
  const atomNs = XmlService.getNamespace('http://www.w3.org/2005/Atom');
  const gdNs = XmlService.getNamespace('gd', 'http://schemas.google.com/g/2005');
  const gContactNs = XmlService.getNamespace('gContact', 'http://schemas.google.com/contact/2008');

  const doc = XmlService.createDocument();
  // Creating elements with prefixed Namespace objects causes XmlService to emit xmlns declarations automatically.
  const entry = XmlService.createElement('entry', atomNs);

  // If updating, keep the id (required by the API)
  if (existingJsonEntry && existingJsonEntry.id && existingJsonEntry.id.$t) {
    entry.addContent(
      XmlService.createElement('id', atomNs).setText(existingJsonEntry.id.$t)
    );
  }

  // Title is what shows up in directory search
  const titleText = c.displayName || [c.firstName, c.lastName].filter(Boolean).join(' ').trim();
  entry.addContent(XmlService.createElement('title', atomNs).setText(titleText));

  // Name block (first/last)
  const nameEl = XmlService.createElement('name', gContactNs);
  nameEl.addContent(XmlService.createElement('givenName', gContactNs).setText(c.firstName || ''));
  nameEl.addContent(XmlService.createElement('familyName', gContactNs).setText(c.lastName || ''));
  entry.addContent(nameEl);

  // Primary email
  if (c.email) {
    const emailEl = XmlService.createElement('email', gdNs)
      .setAttribute('rel', 'http://schemas.google.com/g/2005#work')
      .setAttribute('primary', 'true')
      .setAttribute('address', c.email);
    entry.addContent(emailEl);
  }

  // Phone
  if (c.phone) {
    const phoneEl = XmlService.createElement('phoneNumber', gdNs)
      .setAttribute('rel', 'http://schemas.google.com/g/2005#work')
      .setText(c.phone);
    entry.addContent(phoneEl);
  }

  // Organization fields for quick searching
  const orgEl = XmlService.createElement('organization', gdNs)
    .setAttribute('rel', 'http://schemas.google.com/g/2005#work')
    .setAttribute('primary', 'true');
  orgEl.addContent(XmlService.createElement('orgName', gdNs).setText(CONFIG.WING + ' External'));
  orgEl.addContent(XmlService.createElement('orgTitle', gdNs).setText(c.dutyTitle || 'External Contact'));
  entry.addContent(orgEl);

  // Notes: CAPID / OrgID / status so admins can trace provenance
  const notes = [
    c.capid ? `CAPID: ${c.capid}` : '',
    c.orgid ? `OrgID: ${c.orgid}` : '',
    c.rank ? `Rank: ${c.rank}` : '',
    c.type ? `Type: ${c.type}` : '',
    c.status ? `Status: ${c.status}` : ''
  ].filter(Boolean).join(' | ');

  if (notes) {
    entry.addContent(
      XmlService.createElement('content', atomNs)
        .setAttribute('type', 'text')
        .setText(notes)
    );
  }

  doc.setRootElement(entry);
  return XmlService.getPrettyFormat().format(doc);
}

/**
 * Deletes one external contact by address. Requires an explicit argument: this
 * previously hard-coded a real person's address, so picking it from the editor's
 * function dropdown and pressing Run would delete that individual's contact
 * without any prompt. An explicit argument makes that impossible, and keeps a
 * member address out of version control.
 *
 * @param {string} email - Address of the external contact to delete
 * @returns {string} Result from deleteExternalContactByEmail()
 */
function testDeleteExternalContact(email) {
  const target = String(email === undefined || email === null ? '' : email).trim();
  if (!target) {
    throw new Error('testDeleteExternalContact(email) requires an explicit address — ' +
      'refusing to guess which contact to delete. Call it as ' +
      'testDeleteExternalContact(\'someone@example.org\').');
  }

  const result = deleteExternalContactByEmail(target);
  Logger.info("Delete result: " + result);
  return result;
}

/**
 * Deletes a Domain Shared Contact by email address.
 *
 * @param {string} email - The email address of the contact to delete.
 * @returns {string} 'deleted' | 'not_found' | 'failed'
 */
function deleteExternalContactByEmail(email) {
  if (!email) {
    Logger.warn('deleteExternalContactByEmail called with empty email');
    return 'failed';
  }

  const token = ScriptApp.getOAuthToken();
  const domain = CONFIG.DOMAIN;
  const feedBase = `https://www.google.com/m8/feeds/contacts/${encodeURIComponent(domain)}/full`;

  try {
    // 1) Find existing entry by email
    const findUrl = `${feedBase}?q=${encodeURIComponent(email)}&max-results=1&alt=json`;
    const findResp = UrlFetchApp.fetch(findUrl, {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + token,
        'GData-Version': '3.0'
      },
      muteHttpExceptions: true
    });

    const findCode = findResp.getResponseCode();
    if (findCode < 200 || findCode >= 300) {
      Logger.error('Shared contact search failed during delete', {
        email: email,
        code: findCode,
        response: findResp.getContentText()
      });
      return 'failed';
    }

    const found = JSON.parse(findResp.getContentText());
    const entries = (found.feed && found.feed.entry) ? found.feed.entry : [];

    if (!entries.length) {
      Logger.info('Contact not found for deletion', { email: email });
      return 'not_found';
    }

    // 2) Delete (DELETE to edit link)
    const entry = entries[0];
    const links = entry.link || [];
    const editLink = links.find(l => l.rel === 'edit');
    const editHref = editLink ? editLink.href : '';

    if (!editHref) {
      Logger.warn('Contact found but missing edit link', { email: email });
      return 'failed';
    }

    const etag = entry['gd$etag'] || '*';

    const deleteResp = UrlFetchApp.fetch(editHref, {
      method: 'delete',
      headers: {
        Authorization: 'Bearer ' + token,
        'GData-Version': '3.0',
        'If-Match': etag
      },
      muteHttpExceptions: true
    });

    const delCode = deleteResp.getResponseCode();
    if (delCode < 200 || delCode >= 300) {
      Logger.error('Shared contact delete failed', {
        email: email,
        code: delCode,
        response: deleteResp.getContentText()
      });
      return 'failed';
    }

    Logger.info('Shared contact deleted', { email: email });
    return 'deleted';

  } catch (e) {
    Logger.error('Exception during deleteExternalContactByEmail', {
      email: email,
      error: e.message
    });
    return 'failed';
  }
}
