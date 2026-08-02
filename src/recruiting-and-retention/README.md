# Retention Email Automation

Automated email system for CAP member retention, sending personalized emails to members at key lifecycle points.

## Overview

This module automatically identifies and emails CAP members who are:
- **Turning 18** - Cadets transitioning to senior member eligibility
- **Turning 21** - Cadets aging out of the cadet program
- **Expiring** - Members whose membership expires this month

Emails are personalized with member rank/name and include squadron commanders on cadet emails for awareness and follow-up support. Commanders are CC'd at their **CAP account**, not the personal address on their CAPWATCH record — see [Commander's CC Bounced](#commanders-cc-bounced) for how that address is chosen.

## Features

- ✅ Automated member identification based on CAPWATCH data
- ✅ Personalized email templates with rank and name
- ✅ Squadron commander CC for cadet emails
- ✅ Every send recorded in the retention log spreadsheet for tracking
- ✅ Comprehensive logging to spreadsheet
- ✅ Summary report emailed to retention team
- ✅ Error tracking and retry logic
- ✅ Rate limiting to prevent Gmail quota issues

## Email Types

### Turning 18 Email
**Template:** `Turning18Email.html`  
**Recipients:** Active cadets turning 18 this month  
**CC:** Squadron Commander + **Deputy Commander for Cadets**  
**Purpose:** Inform cadet about transition to senior member opportunities

### Turning 21 Email
**Template:** `Turning21Email.html`  
**Recipients:** Active cadets turning 21 this month  
**CC:** Squadron Commander + **Deputy Commander for Cadets**  
**Purpose:** Inform cadet about aging out of cadet program and senior membership

### Expiring Membership Email
**Template:** `ExpiringEmail.html`  
**Recipients:** Active cadets and seniors expiring this month  
**CC:** Squadron Commander + Recruiting Officer — **cadets only**. A senior's renewal carries no
unit CC at all.  
**Purpose:** Remind member to renew membership before expiration, and invite feedback by reply

### Unit Renewal Digest

**Recipients:** Squadron Commander (addressee) + Recruiting Officer (copied)  
**Content:** Every member under their command expiring this month — **cadets and seniors**  
**Purpose:** Give the unit a worklist for follow-up

This is how a unit hears about its **seniors**. A senior's renewal notice is between them and the
wing, so it is not copied to anyone; the unit gets an addressed worklist instead of a blind copy of
someone else's mail. Cadets appear in the digest as well, so the unit sees one complete list — and
separately keep the CC on their own notice, which is a cadet protection matter rather than a
retention one.

Addressed to the commander with the recruiting officer copied. A unit with only one of the two gets
it addressed to whichever exists; a unit with **neither** is reported in the run summary, since
those units hear nothing.

### How the unit CC is built

| Recipient | On which mail |
|---|---|
| Squadron Commander | **Cadet** mail only — turning 18/21, and cadet renewals |
| Deputy Commander for Cadets | Turning 18/21 |
| Recruiting Officer | Cadet renewals (CC), and every unit digest |

Duty titles come from `RETENTION_CONFIG.CC_DUTY_TITLES` and are matched through
`formatDutyTitle_()`, so legacy `Recruiting & Retention Officer` rows match the current title and
the trailing whitespace the CAPWATCH feed ships on duty values is irrelevant.

- **Each recipient is independent.** A unit with no reachable commander still reaches its duty
  holders. The only empty CC is one where nobody resolved — which for a senior renewal means the
  unit has no Recruiting Officer.
- **A duty nobody holds is simply absent** from the CC.
- **Primary beats assistant**, so "the unit's recruiting officer" resolves to one person rather
  than a unit's entire staff.
- Addresses resolve through the same chain as the commander's, and the list is **deduplicated** —
  in a small unit one person often holds several of these.

## Setup Instructions

### Step 1: Create Email Templates

Create three HTML email templates in your Google Apps Script project:

1. **Turning18Email.html** - Template for cadets turning 18
2. **Turning21Email.html** - Template for cadets turning 21  
3. **ExpiringEmail.html** - Template for expiring members

All substitution goes through `retentionRenderTemplate_()`, so these placeholders work in
any of the three templates:

| Placeholder | Source | Notes |
|-------------|--------|-------|
| `{{rank}}` | member | Member's rank |
| `{{lastName}}` | member | Member's last name |
| `{{expiration}}` | member | Expiration date (populated for `ExpiringEmail`) |
| `{{wingName}}` | `CONFIG.WING_NAME` | Proper name — "California Wing", "Hawaii Wing" |
| `{{orgLabel}}` | `CONFIG.ORG_LABEL` | Abbreviation — "CAWG", "HIWG", "PCR" |
| `{{signature}}` | `retentionSignatureHtml_()` | Full closing block, name + office + wing |

**Never hard-code a wing name, abbreviation, or role holder in a template.** These are
member-facing, and the whole point of the placeholders is that another wing can adopt this
module by setting Script Properties. An unrecognized placeholder is left in the rendered
output rather than blanked, so a typo shows up in the test email.

**Example template structure:**
```html
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; line-height: 1.6; }
    .header { background-color: #003366; color: white; padding: 20px; }
    .header h1 { text-transform: uppercase; }
    .content { padding: 20px; }
  </style>
</head>
<body>
  <div class="header">
    <h1>Civil Air Patrol<br>{{wingName}}</h1>
  </div>
  <div class="content">
    <p>Dear {{rank}} {{lastName}},</p>

    <!-- Email content here -->

    <p>Sincerely,</p>
    <p>{{signature}}</p>
  </div>
</body>
</html>
```

> **Filenames include the folder.** A file at `src/recruiting-and-retention/ExpiringEmail.html`
> deploys into Apps Script under the literal name `recruiting-and-retention/ExpiringEmail`.
> `retentionRenderTemplate_()` adds that prefix for you — pass it `'ExpiringEmail'`.

### Step 2: Configure Settings

These are **Script Properties**, not `config.gs` constants — `clasp push` overwrites
`config.gs` with the shared copy, so no tenant value may live there. Set them in
Project Settings → Script Properties, or via `setupTenantConfig()`.

| Script Property | Purpose |
|-----------------|---------|
| `TENANT_RETENTION_LOG_SPREADSHEET_ID` | Retention tracking spreadsheet |
| `TENANT_RETENTION_EMAIL` | Retention role group — receives the **run summary only** (not a copy of each send; the Log sheet is the per-send record) |
| `TENANT_DIRECTOR_RECRUITING_EMAIL` | `replyTo` on every send. **Required** — blank makes every send fail |
| `TENANT_DIRECTOR_RECRUITING_NAME` | Signature name. Blank signs with the office title alone |
| `TENANT_AUTOMATION_SENDER_EMAIL` | `from` address. Requires a matching verified Send-As alias |
| `TENANT_SENDER_NAME` | Sender display name |
| `TENANT_TEST_EMAIL` | Recipient for the test functions |
| `TENANT_ITSUPPORT_EMAIL` | Contact shown in the summary report footer |

The director's **address and name are individuals**, so they are deliberately blank in
`config-tenants/*.json` and must be set per project. Update them there when the role changes —
the name goes out on member-facing mail.

The wing labels in the templates come from `TENANT_WING` (or the `TENANT_WING_NAME` /
`TENANT_WING_ABBREVIATION` overrides) and need no retention-specific setup.

### Step 3: Create Retention Log Spreadsheet

1. Create a new Google Spreadsheet
2. Name it "Retention Email Log" (or similar)
3. Copy the spreadsheet ID from the URL
4. Update `RETENTION_LOG_SPREADSHEET_ID` in config.gs

The script will automatically create a "Log" sheet with the following columns:
- Timestamp
- Email Type
- CAPID
- Name
- Email
- Commander CAPID
- Commander Name
- Commander Email

### Step 4: Verify CAPWATCH Data Files

Ensure the following CAPWATCH files are available in your configured folder:
- `Member.txt` - Member data
- `MbrContact.txt` - Contact information
- `Commanders.txt` - Squadron commander assignments

### Step 5: Test the System

Run the test functions to verify everything works:

```javascript
// Test 1: Preview member counts + the resolved unit CC for every unit. Sends nothing.
testRetentionEmail();

// Test 2: Send a single test email with sample data
testSendSingleEmail();

// Test 3: Send test emails using real member data (sent to TEST_EMAIL)
testAllRetentionEmails();

// Just the CC dump, without the member sampling
previewRetentionCcLists();
```

### Recovering missing Log rows

If a run sent mail but failed to log it, `backfillRetentionLogFromSentMail()` reconstructs the
rows from the automation account's **sent mail** — the record of what actually happened, carrying
the real send timestamps.

```javascript
// Preview. Sends nothing, writes nothing.
backfillRetentionLogFromSentMail({ period: '2026-08' });

// Then, once the numbers look right:
backfillRetentionLogFromSentMail({ period: '2026-08', write: true });
```

**Run it signed in as the automation account** — it reads that mailbox. Under any other identity
it finds nothing, which the preview states plainly.

Sent mail is used rather than re-deriving from CAPWATCH because those answer different questions:
CAPWATCH says who *would be selected now*, not who *was mailed*. A member who has renewed since
drops out, and one whose expiration has since come into range appears — and writing an
"already sent" row for that second member would silently suppress a mail they never received. A
row is written only where a matching message exists.

Since a sent message carries no CAPID, and the dedupe key is `(type, CAPID)`, CAPWATCH supplies
the CAPID for each address. Families share a primary address, so where one address has fewer
messages than candidates, only as many as were actually sent are written and the remainder are
reported as ambiguous rather than guessed at.

`previewRetentionCcLists()` is the one to read before a first real run. It prints each unit's
commander and CC'd duty holders, the exact CC string each email type would carry, and then three
lists worth acting on:

- **Units with no reachable commander** — these send with no unit CC at all.
- **Unfilled CC duty positions** — the commander is CC'd alone.
- **Derived addresses** — reconstructed as `first.last@<command domain>` because the directory had
  no account for that CAPID. These are **not verified to exist**. A member whose account was
  renamed, created by hand, or is a `.2` duplicate gets an address that Gmail accepts and then
  bounces, per recipient. Spot-check a few in the Admin console. The member's own send is
  unaffected either way.

Review test emails at `TEST_EMAIL` to verify:
- ✅ Templates render correctly
- ✅ Placeholders are replaced with actual data
- ✅ Email formatting looks professional
- ✅ Reply-to and sender settings are correct

### Step 5a: Confirm this is the right tenant

`sendRetentionEmails()` only runs where `PROFILE_.RUN_RETENTION_EMAILS` is true — the **seniors**
profile. It is a no-op elsewhere and returns `{ skipped: true }`.

This is deliberate and is not about scope. The module hardcodes `'CADET'`/`'SENIOR'` rather than
reading `MEMBER_TYPES.ACTIVE`, and both wing tenants download the *same* wing-wide CAPWATCH
extract, so it addresses the whole wing from wherever it runs. Arming it on both tenants does not
split the work — it delivers two copies to every member.

### Step 6: Set Up Trigger

> ⚠️ **Sign in as the automation account first.** A time-driven trigger runs as whoever created
> it, and only that account owns the `AUTOMATION_SENDER_EMAIL` Send-As alias every retention email
> is sent with. Created under any other identity, every send fails with "Invalid argument" — and
> this module fails *worse* than the notification digests do, because the summary is sent the same
> way and its catch only logs. A wrong identity produces **no member mail and no failure summary**.

Signed in as the automation account, run:

```javascript
installRetentionMonthlyTrigger();
```

It removes any existing `sendRetentionEmails` triggers first, so re-running never stacks
duplicates, and leaves other functions' triggers alone. It **refuses to install** on a tenant
whose profile does not run retention, rather than creating a trigger that fires into a no-op and
looks like the feature is running.

Then confirm in the Triggers panel that the listed owner is the automation account.

<details>
<summary>Equivalent manual setup</summary>

1. Open the Apps Script editor
2. Click on the clock icon (Triggers) in the left sidebar
3. Click "+ Add Trigger" in the bottom right
4. Configure the trigger:
   - **Function:** `sendRetentionEmails`
   - **Deployment:** Head
   - **Event source:** Time-driven
   - **Type:** Month timer
   - **Day of month:** 1
   - **Time of day:** 10am to 11am
5. Click "Save"

</details>

**Why 10am?** Emails arrive mid-morning when members are likely to check email, maximizing engagement and response rates.

**Why the 1st?** Ensures CAPWATCH data is updated after month-end processing, and gives members with expiring memberships advance notice.

## Usage

### Automatic Execution

Once the trigger is configured, the system runs automatically on the 1st of each month at 10am. It will:

1. Download fresh CAPWATCH data
2. Identify members in each category
3. Send personalized emails
4. Log all sends to spreadsheet
5. Email summary report to retention team

### Manual Execution

You can also run the system manually:

```javascript
// Run full retention email process
sendRetentionEmails();
```

This is useful for:
- Testing after configuration changes
- Sending emails on a different schedule
- Re-sending if there was an issue

## Monitoring

### Execution Logs

View execution logs in Google Apps Script:
1. Open the script editor
2. Click "Executions" in the left sidebar
3. Review status, duration, and any errors

### Email Log Spreadsheet

Track all sent emails in the retention log spreadsheet:
- View who received emails and when
- See which commander was CC'd
- Identify patterns in member lifecycles
- Export data for retention metrics

### Summary Email

The retention team receives a summary email after each run with:
- Total emails sent by category
- Failed sends (if any)
- Processing duration
- Breakdown by email type

## Troubleshooting

### No Members Found

**Symptom:** "Found 0 members" in logs

**Possible Causes:**
- CAPWATCH data not updated
- No members match criteria this month
- Date parsing issues

**Solution:**
```javascript
// Run test to see member data
testRetentionEmail();

// Check logs for warnings about invalid dates
```

### Template Not Found

**Symptom:** "Template file not found" error

**Possible Causes:**
- HTML file not created in Apps Script project
- Filename doesn't match exactly (case-sensitive)

**Solution:**
1. Verify files exist: `Turning18Email.html`, `Turning21Email.html`, `ExpiringEmail.html`
2. Check for typos in filenames
3. Ensure files are in the same project

### Emails Not Sending

**Symptom:** Members not receiving emails

**Possible Causes:**
- Invalid email addresses in CAPWATCH
- Gmail quota exceeded
- Sender authorization issues

**Solution:**
```javascript
// Check for email validation warnings in logs
// Look for "No valid email" messages

// Verify email quotas (100/day for personal, 1500/day for Workspace)
// https://support.google.com/mail/answer/22839
```

### Commander Not CC'd

**Symptom:** Squadron commander not included on email

**Possible Causes:**
- Commander not listed in Commanders.txt
- Commander has no usable name in Commanders.txt **and** no email in MbrContact.txt — with
  neither, no address can be derived or looked up, and the CC is dropped
- Wrong orgid assignment

### Commander's CC Bounced

**Symptom:** Member received the email; the commander's copy bounced

**Cause:** The CC resolves to the commander's **CAP account**, preferring the real Workspace
address, then the derived `first.last@<command domain>`. Derived addresses are never verified,
so a commander whose account does not follow the default naming — a rename, a manual creation,
a `.2` duplicate — gets an address that does not exist. The member's own send is unaffected.

**Solution:** Confirm the commander has a Workspace account this tenant can read, and that
`TENANT_COMMAND_EMAIL_DOMAIN` names the domain their account is actually on (on the cadets
tenant that is the **senior** domain, since command staff are senior members).

**Solution:**
```javascript
// Test commander lookup
let commander = getCommanderInfo('2503'); // Use actual orgid
console.log(JSON.stringify(commander, null, 2));
```

### Rate Limiting Issues

**Symptom:** "Rate limit exceeded" errors

**Possible Causes:**
- Too many emails sent too quickly
- Other scripts also sending emails

**Solution:**
- Default delay is 100ms between emails
- Increase delay in `RETENTION_CONFIG.EMAIL_DELAY_MS`
- Spread execution across multiple hours

## Configuration Reference

### RETENTION_CONFIG Object

```javascript
const RETENTION_CONFIG = {
  // Email subject lines
  SUBJECTS: {
    TURNING_18: 'Important Membership Update - Turning 18',
    TURNING_21: 'Important Membership Update - Turning 21',
    EXPIRING: 'Your CAP Membership Expires Soon'
  },
  
  // Age thresholds for triggers
  AGE_THRESHOLDS: {
    TRANSITION_TO_SENIOR: 18,  // Cadets turning 18
    CADET_AGE_OUT: 21          // Cadets turning 21
  },
  
  // Rate limiting (milliseconds between emails)
  EMAIL_DELAY_MS: 100,
  
  // Progress logging (log every N emails)
  PROGRESS_LOG_INTERVAL: 10
};
```

## Best Practices

### Template Design
- ✅ Keep emails concise and actionable
- ✅ Include clear next steps for member
- ✅ Provide contact information for questions
- ✅ Use professional CAP branding
- ✅ Test on mobile devices

### Timing
- ✅ Run on the 1st after CAPWATCH updates
- ✅ Send at 10am for optimal open rates
- ✅ Avoid holidays/weekends if possible
- ✅ Give advance notice for expirations

### Monitoring
- ✅ Review logs monthly for errors
- ✅ Check email delivery rates
- ✅ Monitor member response/engagement
- ✅ Update templates based on feedback

### Testing
- ✅ Test after any template changes
- ✅ Verify data accuracy before production
- ✅ Send test emails to yourself first
- ✅ Review all three email types

## Gmail Quotas

Be aware of Google Workspace email quotas:

| Account Type | Daily Limit |
|--------------|-------------|
| Personal Gmail | 100 emails/day |
| Google Workspace | 1,500 emails/day |
| Google Workspace (high-volume) | 10,000 emails/day |

With 100ms delay between emails:
- 10 emails = ~1 second
- 100 emails = ~10 seconds
- 1,000 emails = ~100 seconds (~1.7 minutes)

The script should complete well within execution time limits.

## Data Privacy

This system processes member personal information:
- ✅ Email addresses stored securely in CAPWATCH
- ✅ Logs stored in private spreadsheet (restricted access)
- ✅ No data shared outside organization
- ✅ Members can opt out via email preferences
- ✅ Comply with CAP privacy policies

## Support

### Documentation
- Module code: `SendRetentionEmail.gs`
- Utility functions: `utils.gs`
- Configuration: `config.gs`

### Contacts
- **IT Support:** [IT Support Email from config]
- **Retention Director:** [Director Email from config]
- **Project Maintainers:** Listed in code comments

### Reporting Issues

When reporting issues, include:
1. Execution timestamp
2. Error message from logs
3. Function that failed
4. Member counts from test function
5. Relevant log entries

## Version History

### SendRetentionEmail.gs 1.7.0 (July 2026)
- ✅ **Per-unit renewal digest** to the commander and recruiting officer, listing every expiring
  member under their command — cadets and seniors. Deduplicated per unit per month
- ✅ **Senior renewals carry no unit CC**; the digest is how their unit hears. Cadet renewals
  unchanged (commander + recruiting officer, for cadet protection)
- ✅ Duty holders decoupled from the commander — a unit with no reachable commander still reaches
  its DCC / Recruiting Officer

### SendRetentionEmail.gs 1.6.0 (July 2026)
- ✅ `previewRetentionCcLists()` — resolved unit CC per unit, plus units with no reachable
  commander, unfilled CC duty positions, and every derived (unverified, bounce-prone) address
- ✅ `installRetentionMonthlyTrigger()` — idempotent, 1st at ~10:00, refuses to install on a
  tenant that does not run retention

### SendRetentionEmail.gs 1.5.0 (July 2026)
- ✅ Turning 18/21 also CC the unit's **Deputy Commander for Cadets**; renewals also CC the unit's
  **Recruiting Officer**, where one is assigned
- ✅ Extra staff ride on the commander CC and never replace it; primary beats assistant; list
  deduplicated
- ✅ Legacy `Recruiting & Retention Officer` rows match via `formatDutyTitle_()`

### SendRetentionEmail.gs 1.4.0 (July 2026)
- ✅ **Already-sent guard** — sends filtered against `(email type, CAPID)` for the current
  calendar month, reading the Log sheet the module previously only wrote. Fails open on an
  unreadable log, with a warning in the execution log and a banner on the summary
- ✅ **Tenant guard** — `PROFILE_.RUN_RETENTION_EMAILS`, true on seniors only; arming both wing
  tenants previously mailed every member twice rather than splitting the work
- ✅ Summary email and `testRetentionEmail()` report skipped counts

### SendRetentionEmail.gs 1.3.0 (July 2026)
- ✅ Retention group receives the **run summary only** — dropped `bcc: RETENTION_EMAIL` from all
  three member-facing sends, which at wing scale meant a few hundred messages a month into one
  mailbox, duplicating what the Log sheet already records per send
- ✅ Corrected the retention/reply-to addresses: the previous `recruiting@cawgcap.org` does not
  exist as a group, so the summary had nowhere to land and the failure was swallowed by a catch

### SendRetentionEmail.gs 1.2.0 (July 2026)
- ✅ Commander CC goes to the CAP account (real Workspace address → derived
  `first.last@<command domain>` → CAPWATCH primary as last resort), reusing the resolver
  from `RecoveryEmailNotify.gs` so the two modules stay in agreement
- ✅ `getCommanderInfo()` backed by an index built once per run, instead of re-parsing
  `Commanders.txt` and rebuilding the CAPWATCH email map on every call
- ✅ First test coverage for the module (`test/SendRetentionEmail.commander.test.js`)

### SendRetentionEmail.gs 1.1.0 (July 2026)
- ✅ Templates genericized — `{{wingName}}` / `{{orgLabel}}` / `{{signature}}` replace the
  hard-coded California Wing masthead, footer, and named role holder
- ✅ New optional Script Property `TENANT_DIRECTOR_RECRUITING_NAME` (blank signs with the
  office title alone)
- ✅ `retentionRenderTemplate_()` centralizes substitution across all seven render sites
- ✅ Removed the feedback survey block from `ExpiringEmail.html`, which shipped with an
  unfilled `LINK TO FORM HERE` placeholder — feedback now routes to `replyTo`
- ✅ Removed the commented-out MIWG "Phoenix Senior Flight" block and orphaned CSS
- ✅ Fixed unbalanced `<p>` tags in all three signature blocks

### v1.5 (Public Release) (November 2025)
- ✅ Structured logging with Logger utility
- ✅ Comprehensive error tracking
- ✅ Summary email to retention team
- ✅ Email sanitization and validation
- ✅ Retry logic for transient failures
- ✅ Progress tracking during execution
- ✅ Rate limiting between sends
- ✅ Full JSDoc documentation
- ✅ Fixed template bug (Turning18Email)

## Contributing

When modifying this module:

1. **Test thoroughly** using test functions
2. **Update documentation** if adding features
3. **Follow logging standards** from other modules
4. **Use executeWithRetry()** for API calls
5. **Add JSDoc comments** to new functions
6. **Log errors with context** for troubleshooting
7. **Maintain backward compatibility** with templates

## License

This module is part of the MIWG CAPWATCH Automation system. Internal CAP use only.
