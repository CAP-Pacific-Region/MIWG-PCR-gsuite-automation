/**
 * -------------------------------------------------------------------------
 * File: UpdateCAWGCadetGroups.gs
 * Version: 1.3.0
 * Date: 2026-07-31
 * Author: Lt Col Noel Luneau, Pacific Region
 * Contributors: Maj Isaac Wilson IV, California Wing (1.2.0–1.3.0)
 * Changes: 1.3.0 — a WING-scope cadet source is now the cadet tenant's own
 *   all-hands (ca.all@cawgcadets.org), not ca.cadets@cawgcadets.org, which does
 *   not exist and never has: `.cadets` groups there are created by
 *   updateAllSquadronGroups(), which walks UNIT-scope orgs only.
 *   buildCAWGCadetSourceGroupEmail_() took `scope` and ignored it. So this
 *   function generated a row for a nonexistent address (the add 404s and is
 *   swallowed), left ca.cadets@cawgcap.org with no member that resolves, AND
 *   deleted the hand-added row for the address that does work — its
 *   managed-address pattern matches `.all@` while it generated none. Parents is
 *   deliberately unchanged; see buildCAWGCadetSourceGroupEmail_.
 *   1.2.0 — every Groups row this function nests INTO is now stamped
 *   "Add EXT = Y", not only the rows it creates. It set that column on its own
 *   "cadets"/"parents" rows while nesting them into ".all" rows it left alone,
 *   so it built a dependency it did not satisfy: with the column blank on "all",
 *   updateEmailGroups() declines the external add AND writes
 *   allowExternalMembers=false onto the group on every run. Found on
 *   ca.all@cawgcap.org, where the wing-wide list carried a nested cadet group it
 *   could no longer have re-added. Applies to any wing adopting this.
 * Description:
 * Writes CAWG split-tenant cadet rows into the automation spreadsheet:
 * - "Groups" tab rows for exact destination groups (manualOnly + Add EXT)
 * - "Groups" rows nested into (".all") get Add EXT = Y as well
 * - "User Additions" rows for nested cadet-tenant source groups
 *
 * Example nested mappings:
 *   ca007.cadets@cawgcadets.org   -> ca007.cadets, ca007.all
 *   ca007.parents@cawgcadets.org  -> ca007.parents
 *   ca.all@cawgcadets.org         -> ca.cadets, ca.all   (WING scope: the cadet
 *                                    tenant has no wing-level .cadets group)
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
  const groupsResult = upsertCAWGCadetGroupDefinitionRows_(
    groupsSheet, desired.groupDefinitions, desired.externalNestTargets);
  const additionsResult = upsertCAWGCadetGroupRows_(additionsSheet, desired.userAdditionsRows, cadetTenantDomain);

  Logger.info('CAWG cadet split-tenant rows updated', {
    userAdditionsRowsWritten: desired.userAdditionsRows.length,
    groupsRowsWritten: desired.groupDefinitions.length,
    externalNestTargets: desired.externalNestTargets,
    addExtStamped: groupsResult.addExtStamped,
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
    externalNestTargets: desired.externalNestTargets,
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
    externalNestTargets: collectCAWGCadetExternalNestTargets_(userAdditionsRows),
    groupDefinitions: [
      buildCAWGCadetGroupDefinition_('cadets', 'Cadets', cadetGroupIds),
      buildCAWGCadetGroupDefinition_('parents', 'Parents & Guardians', parentGroupIds)
    ]
  };
}

/**
 * The Groups-tab row names these nestings land in, which therefore have to allow
 * external members.
 *
 * Every address here is on the cadet tenant, so every group it is nested into is
 * holding an address outside this domain. updateEmailGroups() reads permission for
 * that from the Groups sheet ("Add EXT", or "Add Lite" which implies it) and does
 * two things with a blank column: it declines the add, and it writes
 * allowExternalMembers=false onto the group. Both silently, and the second one
 * undoes any repair made in the Admin console by the next morning.
 *
 * This function used to set that column only on the "cadets" and "parents" rows it
 * creates — while nesting them into ".all" rows written by a human. So a wing-wide
 * list could hold a nested cadet group it was no longer permitted to re-add, which
 * is exactly what ca.all@cawgcap.org was found doing.
 *
 * Names are the Groups-tab convention: everything after the first dot of the group
 * id, so "ca.all" and "ca007.all" are both the row named "all".
 *
 * @param {Array<Object>} userAdditionsRows - Rows from buildCAWGCadetManagedRows_
 * @returns {Array<string>} Distinct base group names, sorted
 */
function collectCAWGCadetExternalNestTargets_(userAdditionsRows) {
  const seen = {};

  (Array.isArray(userAdditionsRows) ? userAdditionsRows : []).forEach(row => {
    String((row && row.groups) || '').split(',').forEach(token => {
      const groupId = String(token || '').trim().toLowerCase();
      if (!groupId || groupId.indexOf('.') < 0) return;
      const baseName = groupId.split('.').slice(1).join('.');
      if (baseName) seen[baseName] = true;
    });
  });

  return Object.keys(seen).sort();
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

/**
 * The address on the CADET tenant to nest, for one target.
 *
 * SCOPE MATTERS, and used not to. This function has always taken `scope` and
 * ignored it, so a wing target produced "ca.cadets@cawgcadets.org" — an address
 * nothing creates. On the cadet tenant, `.cadets` groups come from
 * updateAllSquadronGroups(), which reaches an org only through
 * shouldCreateDistributionLists() and so walks UNIT scope only; no CAPWATCH org
 * is wing scope. Verified 2026-07-31: ca.cadets@cawgcadets.org returns 404 while
 * ca.all@cawgcadets.org holds 2,646 members and is named "CAWG - Cadets".
 *
 * What DOES exist wing-wide on that tenant is its own all-hands, created by its
 * updateEmailGroups() from the "all" row like any other wing-level group. On a
 * cadet-only tenant that list IS the wing's cadets, so nesting it is not a
 * substitution — it is the same population under the name that exists.
 *
 * The consequence of the old behavior was quiet: a User Additions row pointing at
 * a nonexistent address, a 404 swallowed as "cannot add external member", and a
 * wing-wide ca.cadets@cawgcap.org whose only intended member never resolved. It
 * also made this function delete a hand-added row for the address that does work,
 * since its managed-address pattern matches `.all@` while it generated none.
 *
 * PARENTS IS DELIBERATELY UNCHANGED. Whether a wing-level parents group exists on
 * the cadet tenant has not been established, and this tenant cannot check — a
 * Workspace tenant cannot read the other's directory. Changing it on the same
 * reasoning would be a guess. See docs/TROUBLESHOOTING.md.
 *
 * @param {string} prefix - Group id prefix, e.g. "ca" or "ca007"
 * @param {string} scope - CAPWATCH org scope: WING or UNIT
 * @param {string} kind - "cadets" or "parents"
 * @param {string} cadetTenantDomain - e.g. cawgcadets.org
 * @returns {string}
 */
function buildCAWGCadetSourceGroupEmail_(prefix, scope, kind, cadetTenantDomain) {
  const normalizedPrefix = String(prefix || '').trim().toLowerCase();
  const domain = String(cadetTenantDomain || '').trim().toLowerCase();
  const normalizedScope = String(scope || '').trim().toUpperCase();

  if (kind === 'cadets') {
    return normalizedScope === 'WING'
      ? `${normalizedPrefix}.all@${domain}`
      : `${normalizedPrefix}.cadets@${domain}`;
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

/**
 * Rewrites this function's own rows on the Groups tab, and guarantees the rows it
 * nests into permit external members.
 *
 * @param {SpreadsheetApp.Sheet} sheet - The "Groups" tab
 * @param {Array<Object>} desiredRows - Rows this function owns outright
 * @param {Array<string>=} externalNestTargets - Base group names that must allow
 *   external members; their Add EXT is set to Y, everything else on the row is
 *   left exactly as the human wrote it.
 * @returns {{preservedCount:number, rowCount:number, addExtStamped:Array<string>}}
 */
function upsertCAWGCadetGroupDefinitionRows_(sheet, desiredRows, externalNestTargets) {
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

  // Rows this function does not own, but nests an external address into. The only
  // cell touched is Add EXT — see collectCAWGCadetExternalNestTargets_ for why a
  // blank one silently breaks the nesting it is being asked to hold.
  const nestTargets = new Set(
    (Array.isArray(externalNestTargets) ? externalNestTargets : [])
      .map(n => String(n || '').trim().toLowerCase())
      .filter(Boolean)
  );
  const addExtStamped = [];
  const alreadyAllowsExternal_ = v =>
    ['y', 'yes', 'x', 'true'].indexOf(String(v || '').trim().toLowerCase()) > -1;

  const out = [headerRow];
  for (let r = 1; r < values.length; r++) {
    const row = values[r] || [];
    const padded = row.slice();
    while (padded.length < width) padded.push('');

    const groupName = String(padded[idxGroupName] || '').trim().toLowerCase();
    if (groupName && managedGroupNames.has(groupName)) continue;

    if (groupName && nestTargets.has(groupName) && !alreadyAllowsExternal_(padded[idxAddExt])) {
      padded[idxAddExt] = 'Y';
      addExtStamped.push(groupName);
    }

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
    rowCount: Math.max(out.length - 1, 0),
    addExtStamped: addExtStamped
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
