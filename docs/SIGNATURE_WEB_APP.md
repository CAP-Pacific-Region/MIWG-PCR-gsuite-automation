# Member self-service email signature web app

A member opens a link, signs in as themselves, and sees the CAP email signature their CAPWATCH
record says they should have. If they approve it, it is written to their own CAP addresses. Two
things are theirs to decide — whether the phone row appears, and which of their own duty
assignments do. Not what any of it says, and not the order.

Source: [`signature-webapp/`](../signature-webapp/README.md) — **a separate Apps Script project**,
not part of `src/`. Tests: `test/SignatureWebApp.test.js`, run by `npm test`.

---

## 1. Why it exists

Signatures reach members three ways today, and none of them is a member self-serving:

| Path | When | Where |
|---|---|---|
| `runDelayedGmailSetup()` | five minutes after an account is created | `src/accounts-and-groups/UpdateMembers.gs` |
| `pushAllSignatures()` | **manually**, over every member at once | same file |
| CAP's own generator | a member builds and pastes one by hand | civilairpatrol.frontify.com |

So a member who joined before signatures were automated, was promoted, took a new duty
assignment, or simply cleared their signature has had to ask IT or hand-build one. This app closes
that gap without giving anyone a bulk-write button.

## 2. What a member may and may not change

**May:**

- include or omit the **phone row**;
- choose **which of their own duty assignments** appear — any of them, assistants included, up to
  the two-line cap.

**Cannot, by anyone:** the signature's **name** in Gmail. Gmail's multiple-signatures feature gives
each signature a label, and the Gmail API does not expose it — the `SendAs` resource carries
`sendAsEmail`, `displayName`, `replyToAddress`, `signature`, `isPrimary`, `isDefault`,
`treatAsAlias`, `smtpMsa` and `verificationStatus`, and `signature` is the HTML and nothing else.
No API client can set that name, so the page says so after a successful apply and points the member
at **Gmail → Settings → See all settings → General → Signature**, which is the only place it can be
changed. The name is shown only to the member; it never appears on mail they send. (`displayName`
*is* settable and is a different thing — the "From:" name — set at provisioning by
`updateGmailSendAsDisplayName()` in `src/`. This app does not touch it.)

**May not:** anything else, and in particular not what a duty *says* or what **order** the block is
printed in. Grade, name, duty titles and unit come from CAPWATCH; the layout and the ordering come
from the CAP brand style guide. The browser never sends HTML, never names an account, and never
supplies a name, grade or duty *text* — it sends one boolean and a list of keys naming duties the
member already holds. The server rebuilds the whole signature from CAPWATCH on every call,
**including on apply**, so the preview is a rendering of the record and not a document the client
can hand back edited.

### How the duty choice is bounded

The selection **filters the member's own duty list**; it is then handed to the same generator that
builds every other signature. So the guarantees are structural rather than a matter of the page
behaving:

| | |
|---|---|
| **Order** | Highest organizational level first, always. A wing *assistant* still precedes a squadron *principal* — the assistant rule orders duties **within** one echelon, never across them. |
| **Cap** | Two lines. Enforced in the generator, and again in the API, and shown in the page by greying the unticked boxes once two are chosen. |
| **Provenance** | Every key is checked against the member's own CAPWATCH record. A key naming a duty they do not hold is **refused, not dropped** — quietly ignoring it would publish a signature they never saw. Likewise three duties are refused, not truncated. |
| **Default** | Touch nothing and you get exactly what the automated path produces: assistants excluded, at most one duty per echelon. |

`getDutyBlock()` in `src/` gained the optional `member.selectedDutyKeys` for this, and **nothing in
`src/` sets it** — provisioning and `pushAllSignatures()` pass a member straight from
`getMembers()`, which has no such field, so their output is byte-identical to before the parameter
existed. That is asserted, along with the port's agreement, in `test/SignatureWebApp.test.js`.

The remedy for a wrong grade or a missing duty is a correction in eServices, which reaches the app
with the next CAPWATCH extract. The page says so.

Two consequences worth knowing before members ask:

- **Cadets never have a phone row at all.** `addContactInfo()` in `src/` refuses to publish a
  cadet phone number to the directory or to a signature, and this app mirrors that rule. On the
  cadets tenant the toggle is replaced by a sentence explaining why there is nothing to toggle.
- **A senior with no cell number in eServices** gets the same treatment: no phone row, and a
  pointer to eServices.

## 3. Access and trust model

| | |
|---|---|
| Deployment | `executeAs: USER_DEPLOYING`, `access: DOMAIN` |
| Who may open it | any authenticated user on the tenant, unless `SIGNATURE_WEBAPP_ALLOWED_GROUP` is set |
| Whose account it acts on | **only the caller's own**, always |
| What it can write | the Gmail signature on Send-As identities whose domain the tenant owns |

`access: DOMAIN` is load-bearing, not a preference: under anonymous access `getActiveUser()`
returns an empty address and every visitor would be indistinguishable. `resolveActor_()` fails
closed if it is ever blank.

There is deliberately **no admin-acts-for-a-member mode**. `pushAllSignatures()` in `src/` already
covers that ground, and adding it here would turn a member-facing page into a way to write to
anybody's mailbox.

`SIGNATURE_WEBAPP_ALLOWED_GROUP` is the pilot switch: set it to a group address and only that
group may use the app; unset it and every member may. Note this is the **reverse** of
`WEBAPP_ALIAS_ADMIN_GROUP` in `webapp/`, which locks that app to nobody when blank — that app
hands out addresses, this one lets people format their own name.

### The service-account key

Gmail settings have no admin-on-behalf-of endpoint: the only way to write a signature is to hold a
token for that user. The app therefore impersonates the caller through the same service account
`src/` uses, which means **`SA_IMPERSONATION_EMAIL` and `SA_PRIVATE_KEY` must exist in this
project's Script Properties too** — Script Properties are per project and there is no shared store.

That is a second copy of a credential that can act as any user in the tenant. It lives only in
Script Properties, never in this repo, and **a key rotation must update both projects** (three, if
`webapp/` is deployed). No new domain-wide-delegation grant is needed as long as the same service
account is used — the `gmail.settings.basic` / `gmail.settings.sharing` scopes are already granted
for the main project. If you use a different service account, grant those two scopes to its client
ID in Admin console → Security → API controls → Domain-wide delegation.

## 4. Setting it up on a tenant

**One project per tenant.** The same `signature-webapp/` source is pushed to a separate Apps
Script project in each Workspace, exactly as `src/` is — a project lives inside one tenant, and a
cadet cannot sign in to a script project that lives on the senior domain. Each has its own clasp
target and its own Script Properties:

| | Seniors | Cadets |
|---|---|---|
| Target | `clasp-targets/signature-webapp-seniors.clasp.json` | `clasp-targets/signature-webapp-cadets.clasp.json` |
| Push | `npm run push:signature:seniors` | `npm run push:signature:cadets` |
| Status / pull / open | `status:` / `pull:` / `open:signature:seniors` | …`:signature:cadets` |

1. **Create the script project** in the tenant (script.new, signed in as the automation account for
   that domain), copy its Script ID into that tenant's clasp target, and push.

   Push to the wrong target and the code lands in the wrong tenant's project — harmless in itself
   (it is the same source), but it will then be configured with the wrong domain, so check
   `scriptId` before the first push of a new project.

   > **The `push:signature:*` scripts carry `--force`, and that is deliberate.** clasp asks before
   > overwriting a project's manifest, and Apps Script normalizes `appsscript.json` server-side, so
   > it sees a manifest difference on *every* push of this project — not just the first. Run
   > non-interactively, it answers its own prompt: the run prints **`Skipping push.`**, exits 0, and
   > changes nothing. That reads exactly like success and will have you testing code you never
   > uploaded. The manifest here is version-controlled and nobody edits it in the console, so
   > overwriting it is what we want anyway. **If you push by hand, pass `--force` yourself.**

2. **Set Script Properties** (Project Settings → Script Properties). Copy the `TENANT_*` values
   straight from `config-tenants/<tenant>.json` — the names are identical to the main project's on
   purpose:

   | Property | Required | Seniors (`cawgcap.org`) | Cadets (`cawgcadets.org`) |
   |---|---|---|---|
   | `TENANT_EMAIL_DOMAIN` | yes | `@cawgcap.org` | `@cawgcadets.org` |
   | `TENANT_SECONDARY_EMAIL_DOMAIN` | if the tenant has one | `@cawg.cap.gov` | **leave unset** — the cadet tenant has none |
   | `TENANT_CAPWATCH_DATA_FOLDER_ID` | yes | that tenant's extract folder | **different folder** — copy from `config-tenants/cadets.json` |
   | `TENANT_WING` | yes | `CA` | `CA` (same wing) |
   | `TENANT_WING_ABBREVIATION` | no | blank → `CAWG` | blank → `CAWG` |
   | `TENANT_ITSUPPORT_EMAIL` | recommended | `it@cawgcap.org` | `it@cawgcap.org` (senior-domain mailbox, as elsewhere) |
   | `SIGNATURE_WEBAPP_ALLOWED_GROUP` | no | blank = every member | blank = every cadet |
   | `SA_IMPERSONATION_EMAIL` / `SA_PRIVATE_KEY` | yes | **that tenant's own** service account | **the cadet tenant's own** — not the senior one |

   ⚠️ Each tenant has its **own** service account with its own domain-wide-delegation grant. Copying
   the seniors key into the cadets project produces `unauthorized_client` on every apply.

   **The Apps Script UI will not store a blank value, so "leave blank" means "do not create the
   property".** That is fine and expected: an absent property reads exactly the same as an empty one
   (`getProperty()` returns null, which the config reader turns into `''`). Omitting
   `SIGNATURE_WEBAPP_ALLOWED_GROUP` means every member may use the app; omitting
   `TENANT_WING_ABBREVIATION` derives `CAWG`; omitting `TENANT_ITSUPPORT_EMAIL` falls back to "your
   wing IT director". The four **required** rows are different — if any is missing, both the page and
   every API call refuse with *"This page is not set up yet"* and name the missing property in the
   execution log. They deliberately do **not** tell the member their account is the problem, which is
   what a missing `TENANT_EMAIL_DOMAIN` used to look like.

   Note that omitting `TENANT_SECONDARY_EMAIL_DOMAIN` on the **seniors** tenant is silent but wrong:
   `@cawg.cap.gov` Send-As identities would simply never be written. The tenant has one — set it.

   `TENANT_PROFILE` is deliberately **not** in this list. Nothing here branches on seniors-vs-cadets
   behavior; the differences that matter (cadet phone numbers, cadet duty positions) come from the
   CAPWATCH record itself. See §7.

3. **Check the deploying account's access.** It must be able to read that tenant's CAPWATCH folder
   and hold admin rights for `admin.directory.user.readonly` on that tenant. Deploy as the
   automation account for the domain, not as a person who might leave.

4. **Deploy — from the Apps Script editor, never from clasp.** Deploy → New deployment → Web app;
   *Execute as* **Me**, *Who has access* **Anyone within \<domain\>**. Give members the `/exec`
   URL — a different URL per tenant. To publish a later change: Deploy → **Manage deployments** →
   pencil → *Version: New version* → Deploy, which keeps the same URL.

   > ⚠️ **`clasp create-deployment` / `clasp update-deployment` break this app's deployment.**
   > Observed twice on the seniors tenant: after a clasp redeploy, visitors get Google's access
   > page instead of the app, and it takes an editor redeploy to clear. The Apps Script API
   > replaces a deployment's config wholesale, and the web-app entry point does not survive it
   > intact — the manifest's `webapp` block is only a default applied when a deployment is
   > *created*, so the access setting a human chose in the editor is not restored from it.
   >
   > `clasp push` is fine and is how code gets to the project. Only the deploy step must be done
   > by a human in the editor. `clasp create-version` is also harmless on its own, but pointless
   > if the editor is going to make its own version anyway.

5. **Smoke-test as a member**, not as the deployer: open the URL from an ordinary member account on
   that domain, confirm the preview matches what `previewSignatureForMember()` logs for that CAPID
   in the same tenant's main project, then apply and check Gmail settings. On the cadets tenant,
   test with a cadet who holds a duty position — that is the case where the record assembly does the
   most work.

## 5. What the apply actually writes

For the caller's mailbox only, it lists Send-As identities and patches `signature` on each one
**whose domain the tenant owns**. Personal identities members have added for their own mail are
listed as skipped and never touched — stamping a CAP signature on somebody's private mail would be
a real intrusion, which is the same reason `updateSignatureForAllAliases()` in `src/` carries this
check. Domain comparison is whole-domain, so a lookalike (`…@cawgcap.org.example.com`) is refused
rather than matched as a suffix.

If a mailbox has no org-owned identity at all, the apply reports an error rather than a cheerful
success that changed nothing.

## 6. Interaction with `pushAllSignatures()`

`pushAllSignatures()` in `src/` is **manual** and rebuilds every member's signature from CAPWATCH
with the phone row included. So:

- Running it over the wing **undoes a member's decision to omit their phone**, and re-applies the
  same content otherwise. The member can simply come back to the app and untick the box again.
- It cannot produce a *different-looking* signature: `test/SignatureWebApp.test.js` renders every
  fixture member through both generators and fails on any byte of difference. Keep that test green
  and the two paths remain interchangeable.

If the phone opt-out ever needs to survive a bulk push, that means a shared store both projects can
read, and it belongs in `src/` — not here.

## 7. Region and national duties (the second CAPWATCH extract)

**CAPWATCH scopes an extract to the echelon it was downloaded as.** A wing pull (ORGID 188)
contains the wing's members and the duties they hold *at wing or below*. A member's region or
national billet is **not in that file at all** — not filtered, not blank, absent. So a member
holding a PCR assignment saw a signature that silently omitted it, with nothing in any log to
explain why, because there was nothing in the data to omit.

The region tenant already downloads a region-wide extract, so the fix is to read that folder as a
**second, read-only source** — the wing tenant keeps operating as the wing in every other respect.

### Setting it up

1. **Share the folder.** The region tenant's region-wide CAPWATCH folder (`PCR CapWATCH Data`,
   `1lU9yWHPf1Eij3AEQPmMR8ki7EpgslV9z`) must be shared **read-only** with the tenant's automation
   account. It lives in a different Workspace, so this is an external share and must be permitted
   between the two domains.
2. **Set the property** on the tenant that needs it — remembering that the web app has **its own**
   Script Properties, separate from the main automation project:

   ```
   TENANT_REGION_CAPWATCH_DATA_FOLDER_ID = 1lU9yWHPf1Eij3AEQPmMR8ki7EpgslV9z
   ```

   On the web app project it feeds the signature page. On the main project it feeds `getMembers()`,
   and therefore provisioning and `pushAllSignatures()`. They are independent — set it on both if
   you want both.
3. **Blank is a valid, complete configuration.** Unset property, missing folder, unreadable folder
   and empty file all degrade to "no extra duties" and never to an error.

### What it may and may not do

| | |
|---|---|
| **Never widens the roster** | Rows are read only for CAPIDs the tenant's own extract already has. The region extract holds every member of every wing in the region; treating it as a source of members would provision Nevada. |
| **Nothing else can see it** | The duties land in their own array, `member.outOfWingDutyPositions`, which only the signature generator reads (via `allSignatureDuties_()`). This was first written to add them to `dutyPositions` while staying clear of `dutyPositionIds` — wrong, because `dutyPositions` is *also* read for the Workspace **directory job title**, for duty-id and duty-level matching in `UpdateGroups.gs`, and for unit distribution lists in `SquadronGroups.gs`. A member's region billet would have reached their GAL title and any wing duty group whose title matched. A separate array makes the guarantee structural instead of a promise each consumer has to keep. |
| **Only orgs outside this wing** | The region extract also carries this wing's own duties; our own pull is authoritative for those. Filtering on the org rather than de-duplicating keeps the rule legible — "duties our own pull cannot see" — instead of depending on which extract downloaded last. |
| **Org names come from the supplement** | It is the only place "Pacific Region" and "National Headquarters" exist, so a region billet renders as *"Pacific Region Assistant Director of IT"* rather than falling back to the member's home unit. `DUTY_LEVEL_ORDER` already ranks NAT/NHQ above REGION above WING. |

### When it doesn't work

Run **`previewOutOfWingDuties`** from the web app project's editor (set your CAPID in
`SIGNATURE_DIAGNOSTIC_RUN_INPUTS` in `MemberRecord.gs` first). It writes nothing, reads no Gmail or
Directory data, and walks the same path the app takes, reporting where it stops: property unset,
folder unreadable, file absent, no rows for that CAPID, or rows found — and for each row whether it
was TAKEN or SKIPPED as ours.

Run it **from the editor**, which executes the project's current code. The deployed app serves
whatever version was last published, and "pushed but not re-deployed" is a common cause that this
distinction isolates immediately.

Two things that look like failure and are not:

- **An assistant billet is not in the default pick.** It appears as an *unticked* box in the duty
  chooser, not in the preview, until the member ticks it. Most region assignments are assistant
  ones, so this is the usual answer.
- **The record cache is 10 minutes.** A page load from before the property was set keeps serving the
  old record. A cache miss logs `Supplementary CAPWATCH consulted for out-of-wing duties` with the
  counts; if that line is absent from the run, you got a cache hit.

## 8. On the cadets tenant

The **same source** runs on both tenants and there is no `TENANT_PROFILE` branch anywhere in it.
Everything that differs for cadets differs because their CAPWATCH record differs, or because the
tenant's own Script Properties differ. What that means in practice:

- **A cadet never gets a phone row.** `addContactInfo()` in `src/` refuses to publish a cadet phone
  number to the directory or a signature, and `MemberRecord.gs` mirrors that rule — a cadet with a
  perfectly good `CELL PHONE` / `PRIMARY` / not-DoNotContact row in `MbrContact.txt` still resolves
  to no number. The page replaces the toggle with a sentence saying so, and asking for the phone
  through the API anyway produces the identical signature.
- **Cadet duty positions come from a different file.** `CadetDutyPositions.txt`, not
  `DutyPosition.txt`. The app reads both, unconditionally, exactly as `getMembers()` does, so a
  cadet commander's line is built the same way a senior's is.
- **No secondary domain**, so only `@cawgcadets.org` Send-As identities are written. A cadet who
  also holds a senior-domain address (staff, or mid-transition) has that address left alone by this
  deployment — it belongs to the seniors tenant's own project, which has its own record for it.
- **Cadet-lite members are covered here but not by `pushAllSignatures()`.** Grades below C/SSgt
  (`CADET`, `C/Amn`, `C/A1C`, `C/SrA`) are filtered out of `getMembers()` because they get no
  account. This app is reached *by a signed-in account*, so provisioning eligibility is not the
  question — if such a member does hold an account, they get a correct signature rather than a
  refusal. Same reasoning for the `EXCLUDED_ORG_IDS` holding units.
- **Two projects, two URLs.** Publish the cadet URL to cadets and the senior URL to seniors; a
  cadet signing in to the seniors deployment is refused by `access: DOMAIN` before any of this code
  runs.
- **Duplicate accounts.** The `duplicate_retired_capid` marker is honored, so a dead twin's CAPID
  never resolves. An account carrying two *live* CAPIDs is refused rather than guessed at.

`test/SignatureWebApp.test.js` §8 pins each of these.

## 9. Troubleshooting

| Symptom | Cause |
|---|---|
| Google's "You need access" page instead of the app | Most often a deployment published with `clasp` — see the warning in §4 step 4. Redeploy from the editor. Otherwise: the visitor opened a `…/edit` link (that is the code, not the app), or the deployment's *Who has access* is narrower than the domain. |
| "This page is not set up yet" | A required Script Property is missing. The execution log names which — see §4 step 2. |
| "This page is for … members signed in to their … account" | Signed in as the wrong Google account, or `SIGNATURE_WEBAPP_ALLOWED_GROUP` excludes them. |
| "Your Workspace account is not linked to a CAPID" | No usable `externalId`/`employeeId` on the account. A `duplicate_retired_capid` marker is ignored by design. |
| "linked to more than one CAPID" | Two live CAPIDs on one account. Fix the directory; the app refuses to guess whose record to print. |
| "not an active member in today's CAPWATCH extract" | Expired member, or the extract has not downloaded. |
| Preview loads, applying fails | Delegation. Check `SA_*` properties in **this** project and the DWD grant for the service account's client ID. |
| `unauthorized_client` on apply, cadets tenant | The seniors service account's key was copied into the cadets project. Each tenant has its own SA and its own DWD grant. |
| A cadet sees "not available" | They opened the **seniors** deployment's URL. `access: DOMAIN` refuses cross-domain callers before any of this code runs. |
| A cadet expects a phone row | There is never one. See §8 — this matches what `src/` publishes for cadets. |
| A member's region or national duty is missing | See §7 — usually an assistant billet sitting unticked in the chooser, or the 10-minute record cache. `previewOutOfWingDuties` settles it. |
| "Can we set the signature's name in Gmail?" | No — the Gmail API has no field for it (§2). The member renames it themselves in Gmail → Settings → General → Signature; the page tells them so after applying. |
| Send-as addresses could not be read | Same as above — the app shows the preview anyway rather than blocking on it. |
| Signature applied but Gmail still shows the old one | Gmail caches settings briefly; reload. Check the member is looking at the same Send-As identity. |

Every request logs a structured line to Cloud Logging (`Signature applied from the self-service
app`, with the caller and the counts). There is no separate audit sheet — the execution log is the
record.
