# Troubleshooting Guide

> **📖 Multi-tenant note.** Inherited from the upstream single-wing project; most diagnostics apply
> unchanged. Two PCR differences to keep in mind: **(1)** each of the three tenants (seniors
> `cawgcap.org`, cadets `cawgcadets.org`, Pacific `pcr.cap.gov`) has its **own** Apps Script project,
> triggers, and Script Properties — diagnose on the affected project, and remember they can drift;
> **(2)** config values (folder IDs, ORGID, domain) live in **Script Properties** (`TENANT_*`), not
> literals in `config.gs`. On this Workspace-for-Nonprofits edition **archiving is a no-op** and
> seats are reclaimed by **deletion** (see the License Management section). Operational escalation:
> the **[Administrator & Successor Guide](ADMIN_GUIDE.md)** (§14 has a symptom→fix table).

This guide helps diagnose and resolve common issues with the CAPWATCH automation system.

## Table of Contents
- [General Troubleshooting](#general-troubleshooting)
- [CAPWATCH Download Issues](#capwatch-download-issues)
- [Member Sync Issues](#member-sync-issues)
- [Email Group Issues](#email-group-issues)
- [License Management Issues](#license-management-issues)
- [Display Name Updates](#display-name-updates)
- [Error Codes Reference](#error-codes-reference)

---

## General Troubleshooting

### Check Execution Logs

1. Open the Apps Script project
2. Click "Executions" in the left sidebar
3. Review recent executions for errors
4. Click on any execution to see detailed logs

### View Structured Logs

The system uses structured JSON logging. To view logs:

```javascript
// Get all logs
var logs = Logger.getAllLogs();
console.log(JSON.stringify(logs, null, 2));

// Get summary
var summary = Logger.getSummary();
console.log(summary);
```

### Common Issues Checklist

- [ ] Are credentials properly configured?
- [ ] Do folder IDs in `config.gs` point to correct locations?
- [ ] Are triggers still active and configured correctly?
- [ ] Has the authorization token expired?
- [ ] Are there any Google Workspace service outages?

---

## CAPWATCH Download Issues

### "CAPWATCH_AUTHORIZATION not set" Error

**Symptom:** Script fails with error about missing authorization

**Cause:** Authorization token not configured

**Solution:**
1. Open `GetCapwatch.gs`
2. Locate `setAuthorization()` function
3. Temporarily add your eServices credentials:
   ```javascript
   let username = 'your-username';
   let password = 'your-password';
   ```
4. Run `setAuthorization()` once
5. **IMMEDIATELY** clear the credentials from the code
6. Run `getCapwatch()` to verify it works

### Download Returns Empty Files

**Symptom:** Files are created but contain no data

**Causes & Solutions:**

1. **Invalid Organization ID**
   - Verify `CONFIG.CAPWATCH_ORGID` is correct (should be your Wing ORGID)
   - Check: `https://www.capnhq.gov/CAP.CapWatchAPI.Web/api/cw?ORGID=223&unitOnly=0`

2. **Expired Credentials**
   - Re-run `setAuthorization()` with fresh credentials
   - Verify credentials work on eServices website first

3. **Network Issues**
   - Check script execution logs for timeout errors
   - Try running during off-peak hours

### Files Not Updating in Drive

**Symptom:** Script runs successfully but files in Drive are outdated

**Causes & Solutions:**

1. **Wrong Folder ID**
   - Verify `CONFIG.CAPWATCH_DATA_FOLDER_ID` in `config.gs`
   - Make sure you have write permissions to the folder
   - Test: `DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID).getName()`

2. **File Permissions**
   - Ensure the script has permission to modify files
   - Re-authorize the script if needed

### API Rate Limiting

**Symptom:** Intermittent failures with 429 error codes

**Solution:**
The system includes automatic retry logic with exponential backoff. If you still see rate limiting:

1. Reduce batch sizes in `config.gs`:
   ```javascript
   BATCH_SIZE: 25  // Down from 50
   ```

2. Add delays between operations:
   ```javascript
   Utilities.sleep(2000);  // 2 second delay
   ```

---

## Member Sync Issues

### Members Not Being Created

**Symptom:** Members exist in CAPWATCH but not in Google Workspace

**Diagnostic Steps:**

1. **Check if member is being processed:**
   ```javascript
   function testGetMember() {
     var members = getMembers();
     var member = members['CAPID'];  // Replace with actual CAPID
     Logger.info('Member data', { member: member });
   }
   ```

2. **Check organization path:**
   ```javascript
   var missing = findMissingOrgPaths();
   console.log(missing);
   ```

3. **Verify member meets criteria:**
   - Status must be 'ACTIVE'
   - Type must be in `CONFIG.MEMBER_TYPES.ACTIVE`
   - Orgid cannot be 0 or 999
   - Organization must have valid `orgPath`

**Common Causes:**

1. **Missing OrgPath**
   - Squadron not in `OrgPaths.txt`
   - Solution: Add squadron to OrgPaths file with correct path

2. **Invalid Member Data**
   - Missing name or CAPID
   - Check `validateMember()` results

3. **API Errors**
   - Check logs for specific error codes
   - See [Error Codes Reference](#error-codes-reference)

### Member Never Received a Welcome Email

**Symptom:** The member has a working account, but never got credentials.

**Cause, almost always:** the account was **created out-of-band** — in the Admin console, by
GAM, or by another admin. `sendWelcomeEmail()` has a single call site, inside the *insert*
branch of `addOrUpdateUser()`. An account that already exists never reaches it: the next sync
finds it, takes the update path, and the member reads as an ordinary existing member. Nothing
detects or repairs this on its own.

**Tell-tale:** an account whose creation date **precedes** the member's Level I completion.
`REQUIRE_LEVEL_I_FOR_SENIORS` withholds new senior accounts until Level I is recorded, so
provisioning could not have created it — someone did it by hand.

**Confirm:**

1. Execution log around the account's creation date — you should find neither
   `New user created` nor `Welcome email sent to IT` for that CAPID, and (for a senior)
   `Senior skipped — Level I not complete` on the runs before it appeared.
2. Every welcome email CCs `ITSUPPORT_EMAIL`. No CC on file is direct proof none was sent.

**Fix:**

```javascript
previewWelcomeEmailResend(123456);  // read-only: would it send, and to whom
resendWelcomeEmail(123456);         // resets the password, then sends
```

This **resets the account password** — the original temp password is stored nowhere, so new
credentials are the only thing that can be delivered. See
[Admin Guide §9](ADMIN_GUIDE.md#9-entry-point-function-reference) for the guards; the refusals
that cannot be forced are described there.

**Stop it recurring:** `sendWelcomeEmail()` records every send in `WelcomeEmailLedger.txt`, so
accounts with no entry are reportable. Run `seedWelcomeLedger(false)` **once** to establish the
baseline, then `scanUnwelcomedAccounts()` (read-only) or arm
`installWelcomeAuditMonthlyTrigger()` to have IT mailed about it monthly. Until the seed runs,
the audit reports nothing — by design, so a missing ledger cannot accuse the whole wing.

> A member who *was* welcomed and simply never logged in is indistinguishable from one who was
> never welcomed. The audit reports those as **UNKNOWN** rather than pretending otherwise —
> review that list, don't bulk-resend against it.

### Members Not Being Updated

**Symptom:** Changes in CAPWATCH not reflected in Google Workspace

**Causes & Solutions:**

1. **Cache Not Cleared**
   ```javascript
   function manualUpdate() {
     clearCache();
     updateAllMembers();
   }
   ```

2. **No Detected Changes**
   The system only updates when it detects changes in:
   - Rank
   - Charter/Organization
   - Duty Positions
   - Status
   - Email

   To force update:
   ```javascript
   function forceUpdateMember() {
     clearCache();
     var members = getMembers();
     addOrUpdateUser(members['CAPID']);  // Replace CAPID
   }
   ```

3. **CurrentMembers.txt Corrupted**
   ```javascript
   function resetCurrentMembers() {
     saveCurrentMemberData({});
     Logger.info('CurrentMembers.txt reset');
   }
   ```

### Suspended Members Not Reactivating

**Symptom:** Member renewed but account still suspended

**Diagnostic Steps:**

1. **Verify member is active in CAPWATCH:**
   ```javascript
   var activeMembers = getActiveMembers();
   console.log(activeMembers['CAPID']);  // Should show join date
   ```

2. **Check if account is suspended or archived:**
   ```javascript
   var inactiveUsers = getInactiveUsers();
   console.log(JSON.stringify(inactiveUsers, null, 2));
   ```

3. **Manually reactivate:**
   ```javascript
   function manualReactivate() {
     reactivateMember('CAPID@miwg.cap.gov', false);
   }
   ```

4. **For archived users:**
   ```javascript
   function manualUnarchive() {
     manualReactivateArchivedUser('CAPID@miwg.cap.gov');
   }
   ```

### Aliases Not Being Created

**Symptom:** Users created but firstname.lastname alias missing

**Solution:**

```javascript
// Update all missing aliases
updateMissingAliases();

// Or manually add for specific user
function addMissingAlias() {
  var user = AdminDirectory.Users.get('CAPID@miwg.cap.gov');
  addAlias(user);
}
```

### Suspended by Mistake

**Symptom:** Active members are being suspended

**Causes:**

1. **In Excluded Organization**
   - Check if member's orgid is in `CONFIG.EXCLUDED_ORG_IDS`
   - MI-000 (744) and MI-999 (1920) are holding units and will be suspended

2. **Grace Period Not Expired**
   - Members get `CONFIG.SUSPENSION_GRACE_DAYS` (default: 7) after expiration
   - Check member's `lastUpdated` field

3. **Not Found in CAPWATCH**
   - Verify member appears in `Member.txt`
   - Check member status is 'ACTIVE'

**Solution:**
```javascript
function checkMemberStatus() {
  var members = getActiveMembers();
  var users = getActiveUsers();
  
  // Find specific user
  var user = users.find(u => u.capid === 'CAPID');
  console.log('User:', user);
  console.log('In CAPWATCH:', members['CAPID']);
}
```

---

## Email Group Issues

### Members Not Added to Groups

**Symptom:** Members should be in group but aren't

**Diagnostic Steps:**

1. **Check group configuration:**
   - Open automation spreadsheet
   - Review 'Groups' sheet
   - Verify attribute and values are correct

2. **Check member qualifications:**
   ```javascript
   function checkGroupMembership() {
     var members = getMembers();
     var member = members['CAPID'];
     
     // Check attributes
     console.log('Type:', member.type);
     console.log('Rank:', member.rank);
     console.log('Duty Positions:', member.dutyPositionIds);
     console.log('Email:', member.email);
   }
   ```

3. **Review error emails:**
   - Check 'Error Emails' sheet in automation spreadsheet
   - Look for the member's email
   - Review error codes and messages

**Common Causes:**

1. **Missing Email**
   - Member must have email in CAPWATCH
   - Check MbrContact.txt for PRIMARY EMAIL

2. **Invalid Email Format**
   - Email fails validation
   - Check logs for "Invalid email format" warnings

3. **Group Doesn't Exist**
   - System will auto-create groups
   - Check for 404 errors in logs

4. **External Email Issues**
   - Group settings may prevent external members
   - Check group settings in Admin Console
   - Squadron distribution lists auto-apply `allowExternalMembers=true`,
     `whoCanPostMessage=ANYONE_CAN_POST` and `spamModerationLevel=MODERATE` via
     `applyGroupSettings()` (SquadronGroups.gs) on each `updateAllSquadronGroups()`
     run. If a cross-tenant nested group (e.g. `ca###.cadets@cawgcadets.org`) still
     won't add, confirm the `AdminGroupsSettings` advanced service is enabled and
     check the log for "Group settings applied".

### Cadet-lite members are missing from a group, or vanish and come back

**Symptom:** a wing- or unit-level `.all` list holds only the members who have accounts. Or the
membership count changes depending on when you look.

**Cause:** two paths manage `.all` groups and only one of them can see cadet-lite members.
`SquadronGroups.gs` adds them by personal CAPWATCH address; `UpdateGroups.gs` builds its desired
set from `getMembers()`, which on a `CADET_LITE=true` tenant filters them out — so it marks every
one of those addresses for removal. On CAWG that was 1,643 removals at 05:24, undone at 06:01.

**Fix:** set **`Add Lite`** = `Y` on the Groups-sheet row (column beside `Add EXT`). That row's
groups then include cadet-lite members, and both paths agree. `Add Lite` also implies
`allowExternalMembers=true`, so the flag stops being flipped nightly.

**Check it before and after:**

```javascript
// Read-only. Lists what the sheet path currently wants to remove.
var d = getEmailGroupDeltas();
Object.keys(d).forEach(c => Object.keys(d[c]).forEach(g => {
  var r = Object.keys(d[c][g]).filter(e => d[c][g][e] === -1);
  if (r.length && /\.all$/.test(g)) console.log(g + ' would remove ' + r.length);
}));
```

Zero removals on `.all` groups is the healthy state.

### A member keeps disappearing from their unit list

**Symptom:** a member is on the list one day and gone the next, then back again. Or the log
shows `Member already exists` (409) on a group that plainly does not contain them.

**Cause:** two spellings of one Google account. On `gmail.com` dots carry no meaning and
everything after a `+` is a tag, so `first.last@`, `firstlast@` and `firstlast+cap@` are one
mailbox. Before SquadronGroups.gs 1.8.0 membership was compared as strings, so a group holding
one spelling while CAPWATCH supplied another read as a member to add **and** a stranger to
remove — the add 409'd and was swallowed, the remove succeeded.

**Fix:** already in place — `googleAccountKey()` (utils.gs) keys both sides of the diff on the
account rather than the string. If you still see it, check whether the address is a **Workspace
alias** rather than a Gmail variant; those are folded by Google too but not by this key, and the
409 log line now names the member so you can tell.

### Addresses Google will not accept as members

**Symptom:** `Failed to add member ... Resource Not Found: <address>` (404), and the member never
appears on the list. At the end of a run, `Addresses Google would not accept as group members`.

**Cause:** Google verifies `gmail.com` addresses against real accounts, so a typo or a closed
account is refused rather than accepted blindly the way an arbitrary domain would be. Two
patterns are impossible regardless of whether the account exists:

- **plus-addressing** (`name+tag@gmail.com`) — the Directory API refuses it for group membership
- **underscores** in a Gmail username — Gmail permits only letters, digits and dots

**Fix:** none available in code — the address is wrong in CAPWATCH. `sanitizeEmail()` cannot
catch these; they are all well-formed. Take the run's worklist to the unit and have the contact
corrected in eServices.
### A member on the other tenant cannot post to a list

**Symptom:** a senior on `@cawgcap.org` mails a cadet-side list such as
`ca.all@cawgcadets.org` and the message bounces or is held for moderation. Members of
the list's own domain can post fine.

**Cause:** the two tenants are separate Workspace accounts, so **every sender on the
other one is external.** A list at `ALL_IN_DOMAIN_CAN_POST` (or
`ALL_MEMBERS_CAN_POST`, for a sender who is not a member) rejects them. This is also
what blocks cross-tenant fan-out, where the forwarded message still carries the
*original* wing sender.

**Fix:** `whoCanPostMessage` must be `ANYONE_CAN_POST` — Google has no value meaning
"members plus my other domain". Since SquadronGroups.gs 1.6.0 the sync enforces this
on every managed list, so the repair is to run `updateAllSquadronGroups()` on the
tenant that owns the list. Verify first, and after, with
`groupAdministration_auditReceiveListPosting()` (read-only).

**On the openness:** `ANYONE_CAN_POST` accepts mail from anywhere on the internet, not
just the sibling tenant. It is paired with `spamModerationLevel=MODERATE` for that
reason. A list that genuinely must stay closed should be left out of the managed set
rather than hand-set in the console, where the next sync will overwrite it.

### Too Many Members Removed

**Symptom:** Mass removal of members from groups

**Causes:**

1. **CAPWATCH Data Issues**
   - Incomplete download
   - Corrupted files
   - Missing data

2. **Configuration Changes**
   - Group criteria changed in spreadsheet
   - Attribute values modified

**Prevention:**
```javascript
// Preview changes before applying
function previewGroupChanges() {
  var deltas = getEmailGroupDeltas();
  
  // Count changes
  var adds = 0, removes = 0;
  for (var category in deltas) {
    for (var group in deltas[category]) {
      for (var email in deltas[category][group]) {
        if (deltas[category][group][email] === 1) adds++;
        if (deltas[category][group][email] === -1) removes++;
      }
    }
  }
  
  console.log('Adds:', adds, 'Removes:', removes);
  
  if (removes > 100) {
    throw new Error('Too many removes - review before proceeding');
  }
}
```

### Groups Not Being Created

**Symptom:** Script runs but new groups don't appear

**Causes & Solutions:**

1. **Insufficient Permissions**
   - Ensure script has Groups Admin API access
   - Re-authorize if needed

2. **Domain Restrictions**
   - Check domain allows group creation
   - Verify naming conventions

3. **API Errors**
   - Check logs for specific error codes
   - Review creation attempts

### Error Emails Not Clearing

**Symptom:** Same emails appear in Error Emails sheet repeatedly

**Analysis:**

This is by design - the sheet tracks persistent issues. To resolve:

1. **For External Emails (404 errors):**
   - These are likely parent/guardian emails or external contacts
   - Verify the email exists and is accessible
   - Check if group allows external members

2. **For Invalid Emails (400 errors):**
   - Email format is invalid
   - Check CAPWATCH data for typos
   - May need manual correction in eServices

3. **For Duplicate Errors (409):**
   - Member already in group
   - This is informational only
   - No action needed

**Manual Cleanup:**
```javascript
// Clear error sheet after resolving issues
function clearErrorEmails() {
  var sheet = SpreadsheetApp.openById(CONFIG.AUTOMATION_SPREADSHEET_ID)
    .getSheetByName('Error Emails');
  
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  
  Logger.info('Error Emails sheet cleared');
}
```

### Additional Group Members Not Added

**Symptom:** Manual additions from 'User Additions' sheet not working

**Diagnostic Steps:**

1. **Check spreadsheet format:**
   - Column A: Name
   - Column B: Email
   - Column C: Role (MEMBER, MANAGER, or OWNER)
   - Column D: Comma-separated group IDs (without @miwg.cap.gov)

2. **Verify role capitalization:**
   ```javascript
   // Invalid
   member, manager, owner
   
   // Valid
   MEMBER, MANAGER, OWNER
   ```

3. **Check group IDs:**
   ```javascript
   // Invalid
   miwg.commanders@miwg.cap.gov
   
   // Valid
   miwg.commanders
   ```

**Solution:**
```javascript
// Test specific addition
function testAdditionalMember() {
  var groupEmail = 'test-group@miwg.cap.gov';
  var email = 'test@example.com';
  var role = 'MEMBER';
  
  try {
    AdminDirectory.Members.insert({
      email: email,
      role: role
    }, groupEmail);
    Logger.info('Test successful');
  } catch (e) {
    Logger.error('Test failed', e);
  }
}
```

---

## License Management Issues

### "Archiving" isn't reclaiming seats

**Symptom:** You expected `manageLicenseLifecycle()` to archive old accounts and free licenses, but
nothing gets archived and seat usage doesn't drop.

**Cause:** This is **expected** on the Workspace-for-Nonprofits edition. Archived-User licenses are
not provisioned, so `AdminDirectory.Users.update({archived:true})` returns **412** and
`archiveLongSuspendedUsers()` is effectively a no-op. Suspension alone also does **not** free a seat
against the 2,000-user cap — **only deletion does.**

**What actually reclaims seats:** `deleteIneligibleSuspendedUsers()` (called by
`manageLicenseLifecycle()`) deletes accounts that are **suspended in Workspace AND ineligible in
CAPWATCH** after a 30-day grace (`LICENSE_CONFIG.DAYS_BEFORE_DELETE_INELIGIBLE`). Renewed members are
rescued by `reactivateRenewedMembers()` first, so they are skipped.

**Preview before trusting it:**
```javascript
var preview = previewLicenseLifecycle();   // or previewIneligibleMembers()
// Review the ineligible-deletion candidates before the monthly run does it for real.
```

### Ineligible accounts not being deleted

**Symptom:** Suspended, long-ineligible users remain and seats stay exhausted.

**Diagnostic steps:**

1. **Confirm the account is actually ineligible.** It must be **suspended** in Workspace and **not**
   an eligible active CAPWATCH member. A member who renewed will be reactivated and skipped.
2. **Check the grace window.** Deletion only applies after `DAYS_BEFORE_DELETE_INELIGIBLE` (30 days).
3. **Manual member protection.** Accounts added via the `User Additions` / Manual Members sheet
   (e.g. PCR/NHQ staff) are **never** auto-deleted by design.
4. **Preview and run:**
   ```javascript
   var preview = previewIneligibleMembers();   // read-only
   console.log('Would delete:', preview.length);
   // If the list looks right:
   manageLicenseLifecycle();                    // performs the deletion step
   ```

> `deleteLongArchivedUsers()` (the old archive→delete path) is intentionally commented out in
> `manageLicenseLifecycle()` because archiving never happens on this edition. Don't re-enable it
> expecting seat recovery — use the ineligible-deletion path above.

### No License Management Report Received

**Symptom:** Script runs but no email received

**Diagnostic Steps:**

1. **Check email addresses:**
   - Verify `LICENSE_CONFIG.NOTIFICATION_EMAILS` in `config.gs`
   - Ensure addresses are valid

2. **Check email quotas:**
   - Apps Script has daily email quotas
   - See: https://developers.google.com/apps-script/guides/services/quotas

3. **Check script logs:**
   ```javascript
   // Look for email sending errors
   var logs = Logger.getAllLogs();
   var emailLogs = logs.filter(l => l.message.includes('email'));
   console.log(emailLogs);
   ```

**Manual Test:**
```javascript
function testLicenseReport() {
  var summary = {
    archived: [],
    deleted: [],
    errors: [],
    startTime: new Date().toISOString(),
    endTime: new Date().toISOString(),
    duration: 1000
  };
  
  sendLicenseManagementReport(summary);
}
```

### Reactivated User Still Has No License

**Symptom:** User reactivated but can't access services

**Cause:** Moving from archived to active requires license reassignment

**Solution:**
1. In Google Admin Console
2. Go to Users
3. Find the user
4. Click "Licenses"
5. Assign appropriate license

Or via script:
```javascript
// Note: License management requires Admin SDK License API
// This is typically done through Admin Console
function assignLicense(email) {
  // This requires additional API setup
  Logger.warn('License assignment should be done via Admin Console');
}
```

---

## Display Name Updates

### Display Names Not Updating

**Symptom:** SendAs display names don't match CAPWATCH data

**Diagnostic Steps:**

1. **Check GAM is installed and configured:**
   ```bash
   gam version
   ```

2. **Verify cron job is running:**
   ```bash
   crontab -l
   ```

3. **Check log files:**
   ```bash
   ls -la *_gam-job.log
   tail -50 $(ls -t *_gam-job.log | head -1)
   ```

4. **Test manually:**
   ```bash
   gam print users primaryEmail aliases lastname firstname custom all query "orgUnitPath=/MI-001 isSuspended=False" > test_users.csv
   
   # Check if CSV has data
   head test_users.csv
   ```

**Common Causes:**

1. **GAM Configuration Issue**
   - Re-run GAM authorization
   - Verify OAuth credentials are valid

2. **Query Filter Issue**
   - Check orgUnitPath matches your structure
   - Verify suspended filter is correct

3. **Custom Schema Not Populated**
   - Ensure UpdateMembers runs before display name update
   - Verify custom fields exist in users

### SendAs Not Being Created

**Symptom:** Display name update fails because SendAs doesn't exist

**Solution:**

The create command runs first, then update. If it still fails:

```bash
# Manually create SendAs for specific user
gam user CAPID@miwg.cap.gov sendas firstname.lastname@miwg.cap.gov name "Last, First Rank CAP GLR-MI-XXX" default treatasalias True
```

### Format Issues

**Symptom:** Display names have wrong format or missing data

**Check Custom Schema:**
```javascript
function checkCustomData() {
  var user = AdminDirectory.Users.get('CAPID@miwg.cap.gov', {
    projection: 'custom',
    customFieldMask: 'MemberData'
  });
  
  console.log('Rank:', user.customSchemas.MemberData.Rank);
  console.log('Org:', user.customSchemas.MemberData.Organization);
}
```

**Verify GAM Template:**
The format is: `Last, First Rank CAP GLR-MI-XXX`

- `~~name.familyName~~` = Last name
- `~~name.givenName~~` = First name
- `~~customSchemas.MemberData.Rank~~` = Rank
- `~~customSchemas.MemberData.Organization~~` = Charter (NER-MI-100)

---

## Error Codes Reference

### HTTP Error Codes

| Code | Meaning | Common Causes | Solution |
|------|---------|---------------|----------|
| 400 | Bad Request | Invalid email format, malformed data | Validate input data, check email format |
| 401 | Unauthorized | Invalid or expired credentials | Re-run setAuthorization(), verify eServices credentials |
| 403 | Forbidden | Insufficient permissions | Check API enablement, verify admin role |
| 404 | Not Found | User/group doesn't exist, external email not found | Create resource, verify email exists |
| 409 | Conflict | Duplicate resource (alias/member already exists) | Check if already exists, use different identifier |
| 429 | Rate Limited | Too many API requests | Reduce batch size, add delays, use retry logic |
| 500 | Server Error | Google server issue | Retry operation, check Google Workspace status |
| 503 | Service Unavailable | Temporary outage | Wait and retry, check status dashboard |

### CAPWATCH-Specific Issues

| Issue | Symptom | Solution |
|-------|---------|----------|
| Empty download | Files created but no content | Verify ORGID, check credentials, check API URL |
| Incomplete data | Some members missing | Check CAPWATCH download completed, verify filters |
| Stale data | Old data in files | Clear cache, re-run getCapwatch() |
| Parse errors | Script fails reading files | Verify CSV format, check for corruption |

### Common Log Messages

**"Invalid email format - skipping"**
- Member's email in CAPWATCH doesn't match email regex
- Check MbrContact.txt for malformed emails

**"Member already in group"**
- Informational only
- No action needed

**"Cannot add external member - not found"**
- External email (parent/guardian) doesn't exist or group settings prevent external members
- Verify email exists, check group settings
- For squadron lists, `allowExternalMembers` is applied automatically by
  `applyGroupSettings()`; a persistent failure here usually means the target
  external address/group doesn't exist, not a settings problem

**"File not found"**
- CAPWATCH file missing from Drive folder
- Re-run getCapwatch(), verify folder ID

**"Max retries exceeded"**
- Operation failed after 3 attempts
- Check error details, may indicate persistent issue

---

## Getting Help

### Before Contacting Support

1. **Check logs:**
   - Execution logs in Apps Script
   - GAM job logs on Linux server
   - Error Emails sheet in automation spreadsheet

2. **Gather information:**
   - What operation were you performing?
   - When did it last work correctly?
   - What error messages did you see?
   - Have you made any recent configuration changes?

3. **Try common fixes:**
   - Clear cache and re-run
   - Verify configuration in config.gs
   - Check triggers are still active
   - Re-authorize if needed

### Log Information to Provide

When reporting issues, include:

```javascript
// Run this to get comprehensive diagnostic info
function getDiagnostics() {
  var info = {
    // Configuration
    config: {
      domain: CONFIG.DOMAIN,
      wingOrgid: CONFIG.CAPWATCH_ORGID,
      suspensionGraceDays: CONFIG.SUSPENSION_GRACE_DAYS
    },
    
    // Recent execution summary
    logSummary: Logger.getSummary(),
    
    // File status
    files: {},
    
    // API status
    apiTests: {}
  };
  
  // Check CAPWATCH files
  try {
    var folder = DriveApp.getFolderById(CONFIG.CAPWATCH_DATA_FOLDER_ID);
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      info.files[file.getName()] = {
        size: file.getSize(),
        lastUpdated: file.getLastUpdated()
      };
    }
  } catch (e) {
    info.files = { error: e.message };
  }
  
  // Test API access
  try {
    AdminDirectory.Users.list({
      domain: CONFIG.DOMAIN,
      maxResults: 1
    });
    info.apiTests.adminDirectory = 'OK';
  } catch (e) {
    info.apiTests.adminDirectory = e.message;
  }
  
  console.log(JSON.stringify(info, null, 2));
  return info;
}
```

### Support Contacts

- **IT Support:** ITSUPPORT_EMAIL (configured in config.gs)
- **Developer:** Check repository README for current developer contact
- **Project Manager:** Check repository README for current PM contact

### Useful Resources

- [Google Apps Script Documentation](https://developers.google.com/apps-script)
- [Admin SDK Directory API](https://developers.google.com/admin-sdk/directory)
- [GAM Documentation](https://github.com/jay0lee/GAM/wiki)
- [Google Workspace Status Dashboard](https://www.google.com/appsstatus)
