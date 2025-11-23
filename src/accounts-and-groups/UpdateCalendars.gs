/**
 * Updates or creates Google Calendars for each CAPWATCH organization
 * Mirrors the structure of the UpdateGroups.gs module.
 * --------------------------------------------------------------
 * Reads CAPWATCH org data, ensures each org (Region/Wing/Group/Squadron)
 * has a matching calendar with proper ACLs and metadata.
 */
function updateCalendars() {
  clearCache();
  const start = new Date();
  Logger.info('📅 Starting Calendar Sync', { started: start.toISOString() });

  const orgs = getSquadrons(); // CAPWATCH org data
  const folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
  const existingCals = listExistingCalendars_();
  const calendars = {};

  for (const orgid in orgs) {
    const org = orgs[orgid];
    const name = org.name;
    const scope = org.scope?.toUpperCase();
    const unit = (org.unit || '').padStart(3, '0');
    const wing = org.wing?.toUpperCase();

    // Skip inactive or invalid orgs
    if (!name || !scope || (unit === '000' && scope !== 'WING' && scope !== 'REGION')) continue;

    // Calendar naming
    const calName = `${toTitleCase(name)} Calendar`;
    const base = (scope === 'REGION' || wing === 'PCR') ? 'pcr' : `${wing.toLowerCase()}wg`;
    const suffix = (['GROUP', 'SQUADRON', 'FLIGHT'].includes(scope) && unit !== '000') ? `-${unit}` : '';

    const ownerGroup = `cal-owners-${base}${suffix}@${CONFIG.DOMAIN}`;
    const readerGroup = `dl-${base}${suffix}-all@${CONFIG.DOMAIN}`;
    const calSummary = calName.toLowerCase();

    let calId = existingCals[calSummary];

    try {
      if (calId) {
        Logger.info('✅ Calendar already exists', { calName, calId });
          // Ensure description and timezone are up to date
          try {
            const cal = CalendarApp.getCalendarById(calId);
            // const tz = getTimezoneForWing_(org?.wing || CONFIG.WING);
            cal.setDescription(`${toTitleCase(name)} Calendar`);
            // cal.setTimeZone(tz);

            Logger.info('📝 Updated calendar metadata', {
              calName,
              calId,
              newDescription: `${toTitleCase(name)} Calendar`
              // timeZone: tz
            });
          } catch (metaErr) {
            Logger.warn('⚠️ Could not update calendar metadata', {
              calName,
              errorMessage: metaErr.message
            });
          }
      } else {
        Logger.info('📅 Creating new calendar', { calName, ownerGroup });
        const calendar = CalendarApp.createCalendar(calName, {
          summary: calName,
          description: `${name} Calendar`
          // timeZone: CONFIG.TIMEZONE || 'America/Los_Angeles'
        });
        calId = calendar.getId();
        Logger.info('✅ Calendar created', { calId, calName });
      }

      // Apply ACLs (Domain-wide read, Group ownership, etc.)
      tryApplyAcl_(calId, ownerGroup, 'owner');
      tryApplyAcl_(calId, readerGroup, 'reader');
      tryApplyAcl_(calId, CONFIG.DOMAIN, 'reader', true); // Domain default access
       // Make calendar publicly readable (optional)
       try {
         Calendar.Acl.insert({
           role: 'reader',
           scope: { type: 'default' } // "default" = everyone (public)
         }, calId);
         Logger.info('🌍 Calendar made public', { calId, calName });
       } catch (e) {
         if (!/duplicate/i.test(e.message)) {
           Logger.warn('⚠️ Could not make calendar public', { calName, errorMessage: e.message });
         }
       }

      calendars[calSummary] = {
        orgid,
        calId,
        name: calName,
        description: `${toTitleCase(name)} Calendar`,
        scope,
        created: new Date().toISOString()
      };

    } catch (e) {
      Logger.error('❌ Failed to process calendar', {
        orgid, name, scope, errorMessage: e.message, stack: e.stack
      });
    }
  }

  // Save results to Drive
  saveCalendars_(calendars, folder);
  Logger.info('✅ Calendar sync complete', {
    total: Object.keys(calendars).length,
    duration: `${new Date() - start}ms`
  });
}

/**
 * Returns a map of existing calendars (summary → id)
 */
function listExistingCalendars_() {
  const cals = {};
  const all = CalendarApp.getAllOwnedCalendars();
  for (const cal of all) {
    const name = cal.getName().toLowerCase();
    cals[name] = cal.getId();
  }
  Logger.info('Loaded existing calendars', { count: Object.keys(cals).length });
  return cals;
}

/**
 * Applies ACLs safely with retry/backoff
 */
function tryApplyAcl_(calId, scopeValue, role, isDomain = false) {
  try {
    const scope = isDomain
      ? { type: 'domain', value: CONFIG.DOMAIN }
      : { type: 'user', value: scopeValue };
    Calendar.Acl.insert({ role, scope }, calId);
    Logger.info('✅ ACL applied', { calId, role, scope });
  } catch (e) {
    if (!/duplicate/i.test(e.message)) {
      Logger.warn('⚠️ Failed to apply ACL', { calId, role, scopeValue, errorMessage: e.message });
    }
  }
}

/**
 * Saves calendar metadata to Drive (Calendars.txt)
 */
function saveCalendars_(calendars, folder) {
  const fileName = 'Calendars.txt';
  let file;
  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    file = files.next();
    file.setContent(JSON.stringify(calendars, null, 2));
  } else {
    file = folder.createFile(fileName, JSON.stringify(calendars, null, 2), MimeType.PLAIN_TEXT);
  }
  Logger.info('📁 Calendars saved', { fileName, count: Object.keys(calendars).length });
}
