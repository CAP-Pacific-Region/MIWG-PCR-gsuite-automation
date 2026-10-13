# Oregon Wing — first-run checklist

The Oregon project (`clasp-targets/orwg.clasp.json`) was moved from a hand-configured older
fork onto the shared `src/` on **2026-09-28**. It is a **composite** tenant (`orwgcap.org`,
cadets and seniors together, `TENANT_PROFILE=composite`), which is why some steps differ from
the California tenants. This is the order to bring it up in. Nothing here arms a trigger until
step 7.

> **There is no dry run for the member sync.** `updateAllMembers()` writes to the directory
> the first time it runs. Step 3 holds the read-only looks that exist; step 5 is a real run
> and should be watched.

## 0. Open the editor cleanly

- Close the Apps Script tab or hard-reload it (**Ctrl+Shift+R**). A tab left open from before
  the push can write its stale copy back over the new code or delete the new files.
- Confirm you are on the new code: `config.gs` is **1.17.0** and its profile list includes
  `composite`.
- The first run of anything asks for **re-consent** (new scopes).

## 1. Configuration

1. Run `validateTenantConfig()`. Expect `DOMAIN=orwgcap.org, ORGID=930, WING=OR`.
2. In Project Settings → Script Properties confirm `TENANT_PROFILE = composite`. A blank one
   silently means `seniors`, under which **cadets are not provisioned and existing cadet
   accounts read as ineligible**.
3. Confirm the secrets exist: `SA_IMPERSONATION_EMAIL`, `SA_PRIVATE_KEY`
   (and `SA_PRIVATE_KEY_ID` if used).
4. Optional and blank by default: `TENANT_HELP_GUIDE_URL`, `TENANT_SUPPORT_URL` (welcome-email
   links; blank omits the help link and mails `TENANT_ITSUPPORT_EMAIL` for support),
   `TENANT_DIRECTOR_RECRUITING_NAME` (retention mail signature), `TENANT_2SV_SETUP_GROUP`
   (blank disables the 2SV prune).

## 2. CAPWATCH

1. `testGetCapwatch()` — confirms the stored credential works. (`CAPWATCH_AUTHORIZATION` is a
   **User** property: only the account that ran `setAuthorization()` can download, and it must
   own the `getCapwatch` trigger. Make sure the eServices password is not left in source.)
2. `getCapwatch()`. It also runs `syncOrgPaths()`; new units are mapped as **`OR-`** OUs and a
   summary is emailed to `TENANT_ITSUPPORT_EMAIL`. Anything under "needs manual" means a unit
   whose parent is not in `OrgPaths.txt`.
3. Check the data folder now holds fresh `Member.txt`, `Organization.txt`, etc., and that
   `OrgPaths.txt` is present.

## 3. Read-only checks before any write

Run each and read the log — none changes anything:

| Function | What to look for |
|---|---|
| `findMissingOrgPaths()` | Units with no OU mapping; members there would land in the wrong OU. |
| `previewIneligibleMembers()` | Who the new rules consider ineligible. **Cadets appearing here means the profile is wrong.** |
| `previewLicenseLifecycle()` | Should report 0 / 0 / 0 destructive candidates. Do not run `manageLicenseLifecycle()` yet. |
| `previewEmailGroupRows()` | The wing / duty group rows the sync will manage. |

## 4. Sanity-check one account

`testaddOrUpdateUser()` / `testGetMember()` in `UpdateMembers.gs` work on a single member. Use a
member you know and confirm the org, grade and OU come out as expected.

## 5. First real member sync — watch it

`updateAllMembers()` compares against the previously saved member data. Because this tenant
was previously run by older code, expect the first run to touch **many** existing accounts
(org path, external IDs, display names, Send-As names). Run it from the editor, not a trigger,
and watch the log:

- Most changes should be benign updates. A large number of **suspensions** or
  **deletions/restores** is a stop-and-investigate signal.
- Accounts for **CADET SPONSOR** members are new behavior for Oregon and will be created.
- New accounts trigger a welcome email; check its help links and IT-support line look right.
- There is **no Level I wait** on Oregon: new senior accounts are provisioned right away
  (`REQUIRE_LEVEL_I_FOR_SENIORS` is off in the `composite` profile, `config.gs` 1.17.1).

## 6. Groups

1. `updateEmailGroups()` — wing/duty/specialty distribution groups. If it times out on a big day,
   `updateEmailGroupsBatch()` resumes.
2. **Squadron groups are new to Oregon.** The profile creates access and public-contact groups
   and the all-hands / cadets / seniors / parents lists automatically. Run
   `previewSquadronGroups()` first and decide whether Oregon wants them before arming
   `updateAllSquadronGroupsBatch()`.

## 7. Arm triggers — one at a time

Use the schedule in [ADMIN_GUIDE.md §8](ADMIN_GUIDE.md#8-what-runs-when-the-automation-schedule),
created **as the account that holds the CAPWATCH credential**. Suggested order: `getCapwatch`,
`updateAllMembers`, `suspendExpiredMembers`, `updateEmailGroups`, then the rest as each has
been previewed.

Inherited from the `seniors` profile and **on** for Oregon — each mails unit commanders, so
preview and confirm the recipients before arming: `notifyLSCodeChanges` (`previewLSCodeChanges`),
`notifyRecoveryEmailCompliance` (`previewRecoveryEmailCompliance`), the parent-email digest
(`previewBadParentEmails`), and the monthly retention mail (`SendRetentionEmail.gs`).

## Do not arm on Oregon

- Cadet→senior **transition** functions and **cross-tenant contacts** — there is no peer tenant
  (`TRANSITION_ROLE` and cross-tenant are off in the `composite` profile).
- `updateCAWGCadetGroups` / `previewCAWGCadetGroups` — they nest a separate cadet tenant's groups.
- `updateResources`, mission provisioning, the shared-contacts and region features — not part of
  this deployment.

## Known open items

- Existing time-driven triggers on the project were checked by the Oregon admin before the push;
  they now run the new code.
- Squadron groups, the commander digests and retention mail have never run on Oregon.
- A standard GCP project is needed only if shared contacts is ever used (Contacts API).
