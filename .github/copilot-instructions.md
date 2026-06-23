# AI Coding Assistant Instructions for CAPWATCH Google Workspace Automation

## Project Overview
This is a Google Apps Script (GAS) project that automates Google Workspace account and group management based on Civil Air Patrol (CAP) membership data from CAPWATCH. It synchronizes user accounts, email groups, and manages license lifecycles entirely within Google Workspace.

## Architecture & Data Flow
```
CAPWATCH eServices → Google Drive (CSV files) → GAS Processing → Google Workspace Admin SDK
     ↓                        ↓                        ↓                        ↓
Source of Truth         Daily Downloads         Parse & Transform         API Updates
```

**Key Components:**
- `config.gs`: Centralized configuration constants (update first when deploying)
- `utils.gs`: Shared utilities (caching, retry logic, structured logging)
- `GetCapwatch.gs`: Downloads CAPWATCH data to Google Drive
- `UpdateMembers.gs`: Creates/updates/suspends user accounts
- `UpdateGroups.gs`: Syncs email distribution groups
- `ManageLicenses.gs`: Archives/deletes inactive accounts

## Essential Patterns & Conventions

### Configuration Structure
Always reference `CONFIG` object for settings. Update these when deploying:
```javascript
const CONFIG = {
  CAPWATCH_ORGID: 'your-wing-orgid',
  DOMAIN: 'yourdomain.cap.gov',
  EMAIL_DOMAIN: '@yourdomain.cap.gov',
  CAPWATCH_DATA_FOLDER_ID: 'drive-folder-id',
  AUTOMATION_FOLDER_ID: 'drive-folder-id',
  AUTOMATION_SPREADSHEET_ID: 'sheets-id'
};
```

### API Calls & Error Handling
Wrap all Admin SDK calls with `executeWithRetry()` for automatic exponential backoff:
```javascript
executeWithRetry(() => {
  return AdminDirectory.Users.update(updates, email);
});
```
- Retries on 403/500 errors, doesn't retry 400/404/409/401
- 400: Bad request (fix data)
- 404: User/group not found
- 409: Already exists/conflict

### File Parsing & Caching
Use `parseFile()` for CAPWATCH CSV data - automatically caches results:
```javascript
const members = parseFile('Member');  // Parses Member.txt, caches
// Subsequent calls return cached data
clearCache();  // Clear at start of major operations
```

### Structured Logging
Use `Logger` utility for consistent logging:
```javascript
Logger.info('Operation started', { count: 5 });
Logger.error('API call failed', error);
Logger.warn('Unexpected data', { field: 'missing' });
```

### Batch Processing
Process in small batches to avoid GAS timeouts (6-10 min limit):
```javascript
const BATCH_SIZE = 50;  // From CONFIG
for (let i = 0; i < members.length; i += BATCH_SIZE) {
  const batch = members.slice(i, i + BATCH_SIZE);
  // Process batch
  Utilities.sleep(250);  // Rate limiting
}
```

### Email & Account Naming
- **Username**: CAPID@domain (e.g., `123456@miwg.cap.gov`)
- **Alias**: firstname.lastname@domain (e.g., `john.doe@miwg.cap.gov`)
- **Groups**: [org-unit].[function]@domain (e.g., `miwg.cadets@miwg.cap.gov`)

### Organizational Units
Map CAPWATCH org IDs to Google Workspace OUs via `OrgPaths.txt`:
```
ORGID,OrgUnitPath
223,/MI-001
1984,/MI-001/MI-700
```

## Critical Workflows

### Member Synchronization
1. Load previous state from `CurrentMembers.txt`
2. Parse current CAPWATCH data
3. Compare for changes (rank, org, positions, status)
4. Update only changed accounts
5. Save new state to `CurrentMembers.txt`

### License Lifecycle
- **Suspend**: Expired members after 7-day grace
- **Archive**: Suspended >1 year (frees standard license)
- **Delete**: Archived >5 years
- **Reactivate**: Renewed members (even archived)

### Group Management
1. Read group config from automation spreadsheet
2. Build desired memberships based on member attributes
3. Calculate deltas vs current Google Groups
4. Apply changes in batches

## Development & Testing

### Testing Approach
- Use `previewLicenseLifecycle()` to see license changes before applying
- Test with small batches: `getMembers().slice(0, 10)`
- Run individual functions manually in GAS editor
- Check "Executions" log for errors

### Deployment Steps
1. Copy all `.gs` files to new GAS project
2. Update `config.gs` with your wing's IDs
3. Run `setAuthorization()` once with CAPWATCH credentials (then delete)
4. Set up time-driven triggers (see README for schedule)
5. Test with `testGetMember()` and `testGetSquadrons()`

### Common Issues
- **Timeouts**: Reduce `BATCH_SIZE` or `MAX_BATCH_SIZE`
- **Auth errors**: Re-run `setAuthorization()` or check permissions
- **Missing OUs**: Verify `OrgPaths.txt` mappings
- **Group errors**: Check automation spreadsheet configuration

## Key Files to Reference
- [`src/config.gs`](src/config.gs) - All configuration constants
- [`src/utils.gs`](src/utils.gs) - Core utilities and Logger
- [`src/accounts-and-groups/UpdateMembers.gs`](src/accounts-and-groups/UpdateMembers.gs) - Account lifecycle
- [`src/accounts-and-groups/UpdateGroups.gs`](src/accounts-and-groups/UpdateGroups.gs) - Group synchronization
- [`README.md`](README.md) - Setup and architecture overview
- [`docs/DEVELOPMENT.md`](docs/DEVELOPMENT.md) - Detailed development guide</content>
<parameter name="filePath">c:\Users\isaac\OneDrive\Documents\GitHub\MIWG-PCR-gsuite-automation\.github\copilot-instructions.md