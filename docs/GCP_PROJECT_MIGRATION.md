# Migrating an Apps Script project to a standard GCP project

## Why

Each tenant's Apps Script project runs on an **Apps Script-managed "default" GCP
project**. Default projects auto-enable the manifest's *advanced services* (Admin SDK,
Drive, Gmail, Calendar, Chat, Groups Settings) but **do not let you enable additional
APIs** — you get `serviceusage.services.enable` permission denied.

Domain Shared Contacts is written through the legacy m8/Contacts feed, which is served by
`contacts.googleapis.com`. That API is **not** an advanced service and **must be enabled
manually**, which requires a **standard (self-owned) GCP project**. This affects every
shared-contacts feature: `cross-tenant-contacts` and the existing
`accounts-and-groups/SharedContacts.gs` ("External Contacts").

> ⚠️ **The switch is one-way.** Apps Script will not let you revert to a default project
> afterward (you can only move to another standard project). Do every prep step **before**
> switching so the live automation is only unauthorized for the minute it takes to
> re-consent.

## What is NOT affected

- **Service accounts / domain-wide delegation** (`SA_*`, `XT_PEER_SA_*`) are independent of
  the Apps Script project's GCP project — DWD is keyed by client ID in the Workspace admin
  console. The JWT-bearer flows keep working across the switch.
- **Script Properties** (`TENANT_*`, `XT_*`) are untouched.
- **Source / manifest** — no code change; the advanced-service list in `appsscript.json`
  stays as-is (you just re-enable the corresponding APIs in the new project).

## Prerequisites

- A Google Cloud user in the **same Workspace org as the tenant** with rights to create a
  project (or an existing standard project in that org). Create the seniors project under
  the `cawgcap.org` org, the cadets project under `cawgcadets.org` — this is what makes the
  **Internal** OAuth consent screen available (no verification needed).
- Do this per tenant; each tenant gets its **own** standard GCP project.

## Procedure (per tenant)

### 1. Create / choose the standard GCP project
In the correct Workspace org, create a project (e.g. `cawg-seniors-automation`). Note its
**project number**.

### 2. Enable all required APIs (APIs & Services → Library)
Enable **all** of these in the new project — the six the automation already uses plus
Contacts. Missing one breaks that feature after the switch:

1. **Admin SDK API** (`admin.googleapis.com`) — directory + the cross-tenant peer read
2. **Google Drive API**
3. **Gmail API**
4. **Google Calendar API**
5. **Google Chat API**
6. **Groups Settings API**
7. **Contacts API** (`contacts.googleapis.com`) — the reason for this migration

### 3. Configure the OAuth consent screen
APIs & Services → OAuth consent screen → **User type: Internal** → fill app name, user
support email, developer contact → Save. Internal apps skip verification and scope review.

### 4. Switch the Apps Script project
Apps Script editor → **Project Settings → Google Cloud Platform (GCP) Project → Change** →
enter the new project number → **Set project**. This revokes existing authorizations
(expected).

### 5. Re-authorize immediately
Open the editor and run any simple function once (e.g. `validateTenantConfig`), completing
the OAuth consent under the new project. This re-authorizes the script.

### 6. Verify the existing automation still works
Before trusting triggers, manually run the primary entry points and confirm the advanced
services respond (no `PERMISSION_DENIED` / `API disabled`):

- `getCapwatch` (Drive), a member-sync entry (Admin SDK), a groups entry (Groups Settings),
  and any Gmail/Calendar/Chat entry point in normal use.
- Check **Triggers** are still listed; if a trigger's owner differs, that owner must open
  the project once and re-authorize so time-based triggers resume.

### 7. Verify Contacts now works
Run `syncCrossTenantContacts`. `Managed shared-contact list failed … Contacts API disabled`
should be gone; you should reach `Cross-tenant reconcile plan`.

## Rollback

There is **no revert to the default project**. If something is wrong, the only move is to
switch to a *different* correctly-configured standard project. Mitigation is prevention:
complete steps 1–3 fully before step 4, and keep step 5 immediately after step 4.

## Per-tenant status

| Tenant | Workspace org for the GCP project | Needs migration? |
| --- | --- | --- |
| Seniors (`cawgcap.org`) | `cawgcap.org` | Yes — required for cross-tenant contacts |
| Cadets (`cawgcadets.org`) | `cawgcadets.org` | Yes — same, before enabling there |
| Pacific (`pcr.cap.gov`) | `pcr.cap.gov` | Yes *if/when* External Contacts (`RUN_SHARED_CONTACTS`) is exercised |
