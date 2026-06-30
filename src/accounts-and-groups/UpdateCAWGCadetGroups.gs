/**
 * -------------------------------------------------------------------------
 * File: UpdateCAWGCadetGroups.gs
 * Version: 1.1.5
 * Date: 2026-03-30 21:29 PDT
 * Author: Lt Col Noel Luneau, Pacific Region
 * Description:
 * Writes CAWG split-tenant cadet rows into the automation spreadsheet:
 * - "Groups" tab rows for exact destination groups (manualOnly + Add EXT)
 * - "User Additions" rows for nested cadet-tenant source groups
 *
 * Example nested mappings:
 *   ca007.cadets@cawgcadets.org   -> ca007.cadets
 *   ca007.parents@cawgcadets.org  -> ca007.parents
 * -------------------------------------------------------------------------
 */

/**
 * Rebuilds CAWG cadet nested-group rows in the "User Additions" tab.
 *
 * The script is non-destructive:
 * - Rows whose Email matches the generated cadet-group emails are replaced.
 * - All other User Additions rows are preserved.
 *
 * After this runs, use the normal CAWG group sync flow to apply the rows.
 */
function updateCAWGCadetGroups() {
  const automationId = getCAWGCadetGroupsAutomationSpreadsheetId_();
  const automationSs = SpreadsheetApp.openById(automationId);
  const additionsSheet = getCAWGCadetGroupsUserAdditionsSheet_(automationSs);
  const groupsSheet = getCAWGCadetGroupsDefinitionsSheet_(automationSs);

  if (!additionsSheet) {
    throw new Error('User Additions tab not found in automation spreadsheet');
  }
  if (!groupsSheet) {
    throw new Error('Groups tab not found in automation spreadsheet');
  }

  const cadetTenantDomain = getCAWGCadetsTenantDomain_();
  const desired = buildCAWGCadetManagedRows_(cadetTenantDomain);
  const groupsResult = upsertCAWGCadetGroupDefinitionRows_(groupsSheet, desired.groupDefinitions);
  const additionsResult = upsertCAWGCadetGroupRows_(additionsSheet, desired.userAdditionsRows, cadetTenantDomain);

  Logger.info('CAWG cadet split-tenant rows updated', {
    userAdditionsRowsWritten: desired.userAdditionsRows.length,
    groupsRowsWritten: desired.groupDefinitions.length,
    preservedUserAdditionsRows: additionsResult.preservedCount,
    preservedGroupsRows: groupsResult.preservedCount,
    cadetTenantDomain: cadetTenantDomain,
    userAdditionsSheet: additionsSheet.getName(),
    groupsSheet: groupsSheet.getName()
  });
}

/**
 * Preview helper for execution logs.
 *
 * @returns {Array<Object>}
 */
function previewCAWGCadetGroups() {
  const cadetTenantDomain = getCAWGCadetsTenantDomain_();
  const desired = buildCAWGCadetManagedRows_(cadetTenantDomain);

  Logger.info('CAWG cadet split-tenant preview', {
    userAdditionsRows: desired.userAdditionsRows.length,
    groupDefinitions: desired.groupDefinitions.length,
    userAdditionsSample: desired.userAdditionsRows.slice(0, 10),
    groupsSample: desired.groupDefinitions.slice(0, 10)
  });

  return desired;
}

function buildCAWGCadetManagedRows_(cadetTenantDomain) {
  if (typeof getSquadrons !== 'function') {
    throw new Error('getSquadrons() is required to build CAWG cadet groups');
  }
  if (typeof parseFile !== 'function') {
    throw new Error('parseFile() is required to build CAWG cadet groups');
  }

  const squadrons = getSquadrons();
  const wingCode = getCAWGCadetGroupsWingCode_().toLowerCase();
  const activeCadetOrgIds = getCAWGActiveCadetOrgIds_();
  const targets = buildCAWGCadetTargets_(squadrons, activeCadetOrgIds, wingCode);
  const userAdditionsRows = [];
  const cadetGroupIds = [];
  const parentGroupIds = [];

  targets.forEach(target => {
    const cadetsTargetGroupId = `${target.prefix}.cadets`;
    const parentsTargetGroupId = `${target.prefix}.parents`;
    const parentGroupPrefix = getCAWGParentGroupPrefix_(target.org, squadrons);
    const cadetsGroups = [cadetsTargetGroupId, `${target.prefix}.all`];
    const parentsGroups = [parentsTargetGroupId];

    if (parentGroupPrefix) {
      cadetsGroups.push(`${parentGroupPrefix}.cadets`);
      cadetsGroups.push(`${parentGroupPrefix}.all`);
      parentsGroups.push(`${parentGroupPrefix}.parents`);
    }

    cadetGroupIds.push(cadetsTargetGroupId);
    parentGroupIds.push(parentsTargetGroupId);

    userAdditionsRows.push(buildCAWGCadetNestedGroupEntry_(
      buildCAWGCadetStandardGroupName_(target.org, squadrons, 'Cadets'),
      buildCAWGCadetSourceGroupEmail_(target.prefix, target.scope, 'cadets', cadetTenantDomain),
      cadetsGroups.join(',')
    ));

    userAdditionsRows.push(buildCAWGCadetNestedGroupEntry_(
      buildCAWGCadetStandardGroupName_(target.org, squadrons, 'Parents & Guardians'),
      buildCAWGCadetSourceGroupEmail_(target.prefix, target.scope, 'parents', cadetTenantDomain),
      parentsGroups.join(',')
    ));
  });

  userAdditionsRows.sort((a, b) => {
    const c1 = String(a.groups || '').localeCompare(String(b.groups || ''));
    if (c1 !== 0) return c1;
    return String(a.email || '').localeCompare(String(b.email || ''));
  });

  return {
    userAdditionsRows: userAdditionsRows,
    groupDefinitions: [
      buildCAWGCadetGroupDefinition_('cadets', 'Cadets', cadetGroupIds),
      buildCAWGCadetGroupDefinition_('parents', 'Parents & Guardians', parentGroupIds)
    ]
  };
}

function upsertCAWGCadetGroupRows_(sheet, desiredRows, cadetTenantDomain) {
  const defaultHeader = ['Name', 'Email', 'Role', 'Groups'];
  const values = sheet.getDataRange().getValues();

  let headerRow = values && values.length ? values[0] : defaultHeader.slice();
  if (!headerRow.length || headerRow.every(cell => !String(cell || '').trim())) {
    headerRow = defaultHeader.slice();
  }

  const header = headerRow.map(h => String(h || '').trim());
  const idxName = header.findIndex(h => h.toLowerCase() === 'name');
  const idxEmail = header.findIndex(h => h.toLowerCase() === 'email');
  const idxRole = header.findIndex(h => h.toLowerCase() === 'role');
  const idxGroups = header.findIndex(h => h.toLowerCase() === 'groups');

  if (idxEmail < 0) {
    throw new Error('User Additions tab is missing an Email column');
  }
  if (idxGroups < 0) {
    throw new Error('User Additions tab is missing a Groups column');
  }

  const desiredEmails = new Set(
    (Array.isArray(desiredRows) ? desiredRows : [])
      .map(row => String((row && row.email) || '').trim().toLowerCase())
      .filter(Boolean)
  );

  const preservedRows = [headerRow];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const email = String(row[idxEmail] || '').trim().toLowerCase();
    if (!email) {
      preservedRows.push(row);
      continue;
    }

    if (desiredEmails.has(email)) continue;
    if (isManagedCAWGCadetGroupEmail_(email, cadetTenantDomain)) continue;

    preservedRows.push(row);
  }

  const width = Math.max(headerRow.length, defaultHeader.length);
  const out = preservedRows.slice();

  (Array.isArray(desiredRows) ? desiredRows : []).forEach(entry => {
    const email = String((entry && entry.email) || '').trim();
    if (!email) return;

    const row = new Array(width).fill('');
    if (idxName >= 0) row[idxName] = String(entry.name || '').trim();
    row[idxEmail] = email;
    if (idxRole >= 0) row[idxRole] = String(entry.role || 'MEMBER').trim();
    row[idxGroups] = String(entry.groups || '').trim();
    out.push(row);
  });

  const sortKey_ = (row) => {
    const name = idxName >= 0 ? String(row[idxName] || '').trim().toLowerCase() : '';
    const email = String(row[idxEmail] || '').trim().toLowerCase();
    return { name, email };
  };

  const body = out.slice(1);
  body.sort((a, b) => {
    const ka = sortKey_(a);
    const kb = sortKey_(b);
    const c1 = ka.name.localeCompare(kb.name);
    if (c1 !== 0) return c1;
    return ka.email.localeCompare(kb.email);
  });

  out.length = 0;
  out.push(headerRow);
  body.forEach(row => out.push(row));

  sheet.clear({ contentsOnly: true });
  sheet.getRange(1, 1, out.length, width).setValues(out);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, width);

  return {
    preservedCount: preservedRows.length - 1,
    rowCount: out.length - 1
  };
}

function getCAWGActiveCadetOrgIds_() {
  const memberRows = parseFile('Member') || [];
  const out = new Set();

  for (let i = 0; i < memberRows.length; i++) {
    const row = memberRows[i] || [];
    const orgid = String(row[11] || '').trim();
    const type = String(row[21] || '').trim().toUpperCase();
    const status = String(row[24] || '').trim().toUpperCase();

    if (orgid && type === 'CADET' && status === 'ACTIVE') {
      out.add(orgid);
    }
  }

  return out;
}

function getCAWGCadetGroupsAutomationSpreadsheetId_() {
  if (typeof CONFIG === 'undefined' || !CONFIG) {
    throw new Error('CONFIG is not defined');
  }

  const id = String(CONFIG.AUTOMATION_SPREADSHEET_ID || '').trim();
  if (!id) {
    throw new Error('CONFIG.AUTOMATION_SPREADSHEET_ID is not set');
  }

  return id;
}

function getCAWGCadetGroupsUserAdditionsSheet_(automationSs) {
  if (!automationSs) return null;

  return automationSs.getSheetByName('User Additions') ||
    automationSs.getSheetByName('UserAdditions') ||
    automationSs.getSheetByName('USER ADDITIONS');
}

function getCAWGCadetGroupsDefinitionsSheet_(automationSs) {
  if (!automationSs) return null;

  return automationSs.getSheetByName('Groups') ||
    automationSs.getSheetByName('Automation') ||
    automationSs.getSheetByName('GROUPS');
}

function getCAWGCadetGroupsWingCode_() {
  if (typeof CONFIG !== 'undefined' && CONFIG && CONFIG.WING) {
    return String(CONFIG.WING).trim();
  }
  if (typeof WING !== 'undefined' && WING) {
    return String(WING).trim();
  }

  throw new Error('Wing code is not configured');
}

function getCAWGCadetGroupsWingLabel_() {
  const wing = getCAWGCadetGroupsWingCode_().toUpperCase();
  return wing.endsWith('WG') ? wing : `${wing}WG`;
}

function getCAWGCadetsTenantDomain_() {
  if (typeof CADETS_TENANT_DOMAIN !== 'undefined' && CADETS_TENANT_DOMAIN) {
    return String(CADETS_TENANT_DOMAIN).trim().toLowerCase();
  }
  if (typeof CONFIG !== 'undefined' && CONFIG && CONFIG.CADETS_TENANT_DOMAIN) {
    return String(CONFIG.CADETS_TENANT_DOMAIN).trim().toLowerCase();
  }
  if (typeof CONFIG !== 'undefined' && CONFIG && CONFIG.CADET_TENANT_DOMAIN) {
    return String(CONFIG.CADET_TENANT_DOMAIN).trim().toLowerCase();
  }

  const wing = getCAWGCadetGroupsWingCode_().toLowerCase();
  const wingPrefix = wing.endsWith('wg') ? wing : `${wing}wg`;
  return `${wingPrefix}cadets.org`;
}

function buildCAWGCadetTargets_(squadrons, activeCadetOrgIds, wingCode) {
  const targets = new Map();
  const wingOrg = Object.values(squadrons).find(org =>
    org &&
    String(org.scope || '').trim().toUpperCase() === 'WING' &&
    String(org.wing || '').trim().toLowerCase() === wingCode &&
    String(org.unit || '').trim() === '001'
  ) || null;

  targets.set(wingCode, {
    prefix: wingCode,
    scope: 'WING',
    org: wingOrg
  });

  activeCadetOrgIds.forEach(orgid => {
    const org = squadrons[orgid];
    if (!org) return;

    const scope = String(org.scope || '').trim().toUpperCase();
    const unit = String(org.unit || '').trim().padStart(3, '0');

    if (scope === 'UNIT' && unit && unit !== '000' && unit !== '001') {
      targets.set(`${wingCode}${unit}`, {
        prefix: `${wingCode}${unit}`,
        scope: 'UNIT',
        org: org
      });
    }

  });

  return Array.from(targets.values()).sort((a, b) => String(a.prefix || '').localeCompare(String(b.prefix || '')));
}

function getCAWGCadetTargetDisplayBase_(org, squadrons) {
  if (org && String(org.name || '').trim()) {
    return toCAWGCadetTitleCase_(String(org.name || '').trim());
  }

  const wingOrg = Object.values(squadrons || {}).find(item =>
    item &&
    String(item.scope || '').trim().toUpperCase() === 'WING' &&
    String(item.wing || '').trim().toLowerCase() === getCAWGCadetGroupsWingCode_().toLowerCase() &&
    String(item.unit || '').trim() === '001'
  );

  if (wingOrg && String(wingOrg.name || '').trim()) {
    return toCAWGCadetTitleCase_(String(wingOrg.name || '').trim());
  }

  return getCAWGCadetGroupsWingLabel_();
}

function buildCAWGCadetStandardGroupName_(org, squadrons, label) {
  const normalizedLabel = String(label || '').trim();
  if (typeof getSquadronGroupMetadata_ === 'function') {
    const metadata = getSquadronGroupMetadata_(org || null, normalizedLabel);
    const name = String((metadata && metadata.name) || '').trim();
    if (name) return name;
  }

  const baseLabel = getCAWGCadetTargetDisplayBase_(org, squadrons);
  return baseLabel && normalizedLabel ? `${baseLabel} - ${normalizedLabel}` : (baseLabel || normalizedLabel || '');
}

function buildCAWGCadetSourceGroupEmail_(prefix, scope, kind, cadetTenantDomain) {
  const normalizedPrefix = String(prefix || '').trim().toLowerCase();
  const domain = String(cadetTenantDomain || '').trim().toLowerCase();

  if (kind === 'cadets') {
    return `${normalizedPrefix}.cadets@${domain}`;
  }

  if (kind === 'parents') {
    return `${normalizedPrefix}.parents@${domain}`;
  }

  throw new Error('Unknown CAWG cadet source group kind: ' + kind);
}

function buildCAWGCadetGroupDefinition_(groupName, description, groupIds) {
  return {
    category: 'custom',
    groupName: String(groupName || '').trim().toLowerCase(),
    attribute: 'manualOnly',
    values: Array.from(new Set((groupIds || []).map(id => String(id || '').trim().toLowerCase()).filter(Boolean))).join(','),
    description: String(description || '').trim(),
    addExt: 'Y'
  };
}

function getCAWGParentGroupPrefix_(org, squadrons) {
  if (!org) return '';
  if (String(org.scope || '').trim().toUpperCase() !== 'UNIT') return '';

  const parentOrgId = String(org.nextLevel || '').trim();
  const parent = parentOrgId ? squadrons[parentOrgId] : null;
  if (!parent) return '';
  if (String(parent.scope || '').trim().toUpperCase() !== 'GROUP') return '';

  const wing = String(parent.wing || '').trim().toLowerCase();
  const unit = String(parent.unit || '').trim().padStart(3, '0');
  if (!wing || !unit || unit === '000' || unit === '001') return '';

  return `${wing}${unit}`;
}

function upsertCAWGCadetGroupDefinitionRows_(sheet, desiredRows) {
  const defaultHeader = ['Category', 'Group Name', 'Attribute', 'Values', 'Description', 'Add EXT'];
  const values = sheet.getDataRange().getValues();

  let headerRow = values && values.length ? values[0].slice() : defaultHeader.slice();
  if (!headerRow.length || headerRow.every(cell => !String(cell || '').trim())) {
    headerRow = defaultHeader.slice();
  }

  function ensureHeader_(label) {
    const normalized = String(label || '').trim().toLowerCase();
    for (let i = 0; i < headerRow.length; i++) {
      if (String(headerRow[i] || '').trim().toLowerCase() === normalized) return i;
    }
    headerRow.push(label);
    return headerRow.length - 1;
  }

  const idxCategory = ensureHeader_('Category');
  const idxGroupName = ensureHeader_('Group Name');
  const idxAttribute = ensureHeader_('Attribute');
  const idxValues = ensureHeader_('Values');
  const idxDescription = ensureHeader_('Description');
  const idxAddExt = ensureHeader_('Add EXT');
  const width = headerRow.length;

  const managedGroupNames = new Set(
    (Array.isArray(desiredRows) ? desiredRows : [])
      .map(row => String((row && row.groupName) || '').trim().toLowerCase())
      .filter(Boolean)
  );

  const out = [headerRow];
  for (let r = 1; r < values.length; r++) {
    const row = values[r] || [];
    const padded = row.slice();
    while (padded.length < width) padded.push('');

    const groupName = String(padded[idxGroupName] || '').trim().toLowerCase();
    if (groupName && managedGroupNames.has(groupName)) continue;
    out.push(padded);
  }

  (Array.isArray(desiredRows) ? desiredRows : []).forEach(entry => {
    const row = new Array(width).fill('');
    row[idxCategory] = String(entry.category || '').trim();
    row[idxGroupName] = String(entry.groupName || '').trim();
    row[idxAttribute] = String(entry.attribute || '').trim();
    row[idxValues] = String(entry.values || '').trim();
    row[idxDescription] = String(entry.description || '').trim();
    row[idxAddExt] = String(entry.addExt || '').trim();
    out.push(row);
  });

  sheet.clear({ contentsOnly: true });
  sheet.getRange(1, 1, out.length, width).setValues(out);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, width);

  return {
    preservedCount: Math.max(out.length - 1 - ((desiredRows || []).length), 0),
    rowCount: Math.max(out.length - 1, 0)
  };
}

function buildCAWGCadetNestedGroupEntry_(name, email, groups) {
  return {
    name: String(name || '').trim(),
    email: String(email || '').trim().toLowerCase(),
    role: 'MEMBER',
    groups: String(groups || '').trim().toLowerCase()
  };
}

function getCAWGCadetGroupDisplayName_(org, groupId) {
  const baseName = buildCAWGCadetGroupRowName_(org, groupId);
  return baseName;
}

function isManagedCAWGCadetGroupEmail_(email, cadetTenantDomain) {
  const normalizedEmail = String(email || '').trim().toLowerCase();
  const normalizedDomain = String(cadetTenantDomain || '').trim().toLowerCase();
  if (!normalizedEmail || !normalizedDomain) return false;

  const wingCode = getCAWGCadetGroupsWingCode_().trim().toLowerCase();
  const escapedWingCode = wingCode.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const escapedDomain = normalizedDomain.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const pattern = new RegExp(`^${escapedWingCode}(?:\\d{3})?\\.(?:cadets|parents|all|all\\.parents|all-cadets|all-parents)@${escapedDomain}$`);

  return pattern.test(normalizedEmail);
}

function buildCAWGCadetGroupRowName_(org, groupId) {
  if (!org) {
    return `${getCAWGCadetGroupsWingLabel_()} Cadets`;
  }

  const orgName = String((org && org.name) || '').trim();
  if (orgName) return `${toCAWGCadetTitleCase_(orgName)} Cadets`;

  const base = String(groupId || '').split('.')[0].toUpperCase();
  return `${base} Cadets`;
}

function toCAWGCadetTitleCase_(value) {
  const str = String(value || '').trim();
  if (!str) return '';

  const preserve = new Set([
    'CAP', 'USAF', 'FAA', 'DOT', 'TSA', 'ICAO', 'EASA', 'HQ', 'IT', 'ES', 'AEM', 'NCO'
  ]);

  function isWingAcronym_(token) {
    return /^[A-Z]{2,4}WG$/.test(token) || token === 'PCR';
  }

  function titleCore_(core) {
    if (!core) return core;
    if (/\d/.test(core)) return core;

    if (core.includes('-')) {
      return core.split('-').map(titleCore_).join('-');
    }
    if (core.includes('/')) {
      return core.split('/').map(titleCore_).join('/');
    }
    if (core.includes("'")) {
      return core.split("'").map(titleCore_).join("'");
    }

    const upper = core.toUpperCase();
    if (preserve.has(upper) || isWingAcronym_(upper)) return upper;

    return upper.charAt(0) + upper.slice(1).toLowerCase();
  }

  return str
    .split(/\s+/)
    .map(token => {
      const match = token.match(/^(.+?)([.,;:)]*)$/);
      const core = match ? match[1] : token;
      const punct = match ? match[2] : '';
      return titleCore_(core) + punct;
    })
    .join(' ');
}
