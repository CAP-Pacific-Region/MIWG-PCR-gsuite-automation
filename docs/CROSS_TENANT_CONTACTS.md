# Cross-Tenant Shared Contacts

**Module:** `src/cross-tenant-contacts/CrossTenantContacts.gs` (v0.1.0, draft)
**Status:** Drafted, **not yet deployed.** Replaces the two legacy wing projects
(cadets `1fJRqo…`, seniors `1b2JSIB…`) that did the same job via export spreadsheets.

## What it does

The wing runs two Workspace tenants — seniors (`cawgcap.org`) and cadets
(`cawgcadets.org`) — but they are **one CAP wing (CAWG) in CAPWATCH**. Members of one
tenant don't appear in the other's Global Address List, so a senior can't just type a
cadet's name in Gmail. This module fixes that by publishing the **peer** tenant's members
into **this** tenant's Domain Shared Contacts:

- On the **seniors** project it publishes **cadets** into the senior GAL.
- On the **cadets** project it publishes **seniors** into the cadet GAL.
- On **pacific** it is off (no peer tenant).

It is one file, byte-identical on every project (`clasp push` stays a verbatim fan-out).
Nothing tenant-specific is a literal: **behavior** is chosen by `TENANT_PROFILE`
(`PROFILE_.CROSS_TENANT`), **identity/secrets** by Script Properties. The code only ever
talks about *self* and *peer* — each project's properties define what those mean for it.

## Why there are no spreadsheets anymore

The legacy design had each tenant *export* its directory to a Google Sheet, which the
other tenant *imported*. That sheet existed for one reason: a script bound to the seniors
tenant cannot list the cadets tenant's directory with its own identity.

This module removes the sheet by reading the two data sources each tenant already has:

| Data | Source |
| --- | --- |
| Peer roster + name/grade/unit/duty/type — **every peer member, incl. cadet-lite** | one wing CAPWATCH pull via the existing `getMembers(peerTypes)` |
| Authoritative peer **Workspace email** (and whether an account exists) | a **live read of the peer directory** using a read-only peer-tenant service account (domain-wide delegation) |
| Fallback personal email for members with no Workspace account | CAPWATCH `MbrContact` (already attached to `member.email`, `doNotContact`-respected) |

### The email waterfall (`xtResolvePeerEmail_`)

For each peer member from CAPWATCH, the address is resolved in priority order:

1. **Peer Workspace `primaryEmail`** (by CAPID, from the live peer-directory read).
   Authoritative — this is what also corrects name-collision suffixes (`.2`) and manual
   renames, which a derived `firstname.lastname@` address would get wrong.
2. **CAPWATCH personal primary email.** Used for members with **no Workspace account** —
   most importantly the **cadet-lite** cadets below C/SSgt, who don't get a
   `@cawgcadets.org` account until they promote (`CONFIG.CADET_LITE_EXCLUDED_GRADES`).
3. **`do.not.contact+<CAPID>@<selfDomain>` sentinel.** Presence-only: the member appears
   in the GAL for name/grade/unit but isn't emailable (opted out or no email). Controlled
   by `EMIT_PLACEHOLDERS`.

> **Why the roster is complete in both directions.** `getMembers` applies the cadet-lite
> grade filter only when `CONFIG.CADET_LITE` is true *and* the row is a cadet grade. The
> seniors project (`CADET_LITE=false`) fetching `['CADET']` returns **all** cadets incl.
> sub-C/SSgt; the cadets project fetching senior types never trips the cadet filter.

## Reconciliation is stateless

There is no sync-state sheet. Managed contacts are self-describing:

- `gd:orgName` carries a **marker** — `CONFIG.WING` for members, `CONFIG.WING + '_PARENTS'`
  for parent groups — so `xtListManagedContacts_` can find exactly the contacts this module
  owns and never touch anything else in the GAL.
- A content hash rides in a `gContact:userDefinedField` (`xtSyncHash`), so the next run
  diffs create/update/delete by comparing hashes — no external state.

Writes to *this* tenant's shared contacts use the script's own token; the peer service
account is **read-only** and only ever reads the peer directory.

## Parent-group sync

`syncCrossTenantParentContacts` (gated by `RUN_PARENTS`, on for seniors) publishes the
peer tenant's parent/guardian distribution groups into the GAL. It finds them by the
squadron-groups naming convention — email localpart ending in `.parents`
(`ca346.parents@cawgcadets.org`) — via the peer Directory **Groups** API, and manages them
under the separate `<WING>_PARENTS` marker so they never overlap the member sync. This
replaces the legacy `cadetsParentsSync` (and fixes its undefined `SHEET_ID_PARENTS` bug by
removing the sheet entirely).

## Configuration

### Behavior — `PROFILE_.CROSS_TENANT` (version-controlled, `config.gs`)

| Key | seniors | cadets | pacific |
| --- | --- | --- | --- |
| `RUN_INBOUND` | `true` | `true` | `false` |
| `RUN_PARENTS` | `true` | `false` | `false` |
| `PEER_TYPES` | `['CADET']` | `['SENIOR','FIFTY YEAR','INDEFINITE','CADET SPONSOR']` | `[]` |
| `PEER_LABEL` | `'CADET'` | `'SENIOR'` | `''` |
| `EMIT_PLACEHOLDERS` | `true` | `true` | `false` |

### Identity/secrets — Script Properties (per project)

| Key | Secret? | Set via | seniors | cadets |
| --- | --- | --- | --- | --- |
| `XT_PEER_DOMAIN` | no | `setupCrossTenantConfig()` / JSON | `cawgcadets.org` | `cawgcap.org` |
| `XT_PEER_SA_EMAIL` | no | `setupCrossTenantConfig()` | cadets SA `client_email` | seniors SA `client_email` |
| `XT_PEER_SA_SUBJECT` | no | `setupCrossTenantConfig()` | a cadets super admin | a seniors super admin |
| `XT_PEER_SA_KEY` | **yes** | by hand in Script Properties | cadets SA private key | seniors SA private key |

Canonical non-secret values live in `config-tenants/seniors.json` / `cadets.json`.

## Deployment runbook

Do this **once per wing project** (seniors and cadets). Pacific needs none of it. The two
projects are mirror images — each reads the *other* tenant. Nothing here touches the
`pacific` project.

### 0. Repo changes that ship with `clasp push` (already in `src/`)

These are code/manifest changes, identical on every project; they arrive when you push the
shared `src/`. No per-project action beyond the re-consent in step 3.

| File | Change |
| --- | --- |
| `src/cross-tenant-contacts/CrossTenantContacts.gs` | new module (entry points, engine, setup/validate) |
| `src/config.gs` | `PROFILE_.CROSS_TENANT` block on each profile |
| `src/appsscript.json` | adds `https://www.google.com/m8/feeds` scope |

Repo-only (not pushed): `config-tenants/{seniors,cadets}.json` gain `XT_PEER_DOMAIN`;
`config-tenants/README.md` and this doc.

### 1. Create a read-only peer service account (in the PEER tenant)

For the **seniors** project the peer is **cadets**, and vice-versa. In a GCP project you
control:

1. Create a **service account** (or reuse one) and generate a **JSON key**. Note its
   `client_email` and its numeric **OAuth client ID**.
2. In the **peer tenant's** Admin console → **Security → Access and data control → API
   controls → Domain-wide delegation**, add the SA's client ID with **exactly** these
   scopes (read-only, comma-separated):
   ```
   https://www.googleapis.com/auth/admin.directory.user.readonly,
   https://www.googleapis.com/auth/admin.directory.group.readonly
   ```
   (`group.readonly` is only needed where `RUN_PARENTS` is on — i.e. the seniors project —
   but granting both on both projects is harmless and simpler.)
3. Pick a **subject**: any super admin in the peer tenant for the SA to impersonate.

> This is a credential *into the peer tenant*. Keep it a **dedicated, read-only** SA — do
> not reuse the peer tenant's Gmail/Calendar delegation SA, which has broader scopes. The
> SA only ever **reads** the peer directory; all writes are to this tenant's own contacts
> with the script's own token.

### 2. Set Script Properties (on THIS project)

Run `setupCrossTenantConfig()` from the editor to write the three non-secret keys (fill the
blanks in the function first), then add the secret key by hand. Final state per project:

**Seniors project (`cawgcap.org`) — Project Settings → Script Properties**

| Property | Value |
| --- | --- |
| `XT_PEER_DOMAIN` | `cawgcadets.org` |
| `XT_PEER_SA_EMAIL` | *cadets* read-only SA `client_email` |
| `XT_PEER_SA_SUBJECT` | a *cadets* super-admin address |
| `XT_PEER_SA_KEY` | *cadets* SA private key PEM (secret; `\n` or real newlines OK) |

**Cadets project (`cawgcadets.org`) — Project Settings → Script Properties**

| Property | Value |
| --- | --- |
| `XT_PEER_DOMAIN` | `cawgcap.org` |
| `XT_PEER_SA_EMAIL` | *seniors* read-only SA `client_email` |
| `XT_PEER_SA_SUBJECT` | a *seniors* super-admin address |
| `XT_PEER_SA_KEY` | *seniors* SA private key PEM (secret) |

No other new properties are required — roster, duty, and fallback email all come from the
project's existing CAPWATCH pull (ORGID 188 already includes both seniors and cadets) and
`TENANT_*` config. `TENANT_PROFILE` must already be set (`seniors` / `cadets`).

Then run **`validateCrossTenantConfig()`** — it must print `✅ Cross-tenant config OK`.

### 3. Re-consent for the new scope

After the shared `src/` is pushed, the added `m8/feeds` scope requires one re-authorization.
Open the project editor and run any function once (e.g. `validateCrossTenantConfig`), or run
`syncCrossTenantContacts` manually, and complete the OAuth consent when prompted.

### 4. Add triggers (Apps Script → Triggers → Add trigger)

| Project | Function | Type | Cadence |
| --- | --- | --- | --- |
| Seniors | `syncCrossTenantContacts` | Time-driven | Daily |
| Seniors | `syncCrossTenantParentContacts` | Time-driven | Daily |
| Cadets | `syncCrossTenantContacts` | Time-driven | Daily |

Sequencing: run these **after** the peer tenant's own member provisioning and CAPWATCH
refresh for the day, so the peer directory (authoritative emails) and the local CAPWATCH
roster are both current. Daily is sufficient — the GAL isn't latency-sensitive. There is no
cadets parent-group trigger because `RUN_PARENTS` is off for the cadets profile.

### 5. First-run verification

Run each function once by hand and check the execution log:

- `Cross-tenant contact sync complete` with an `emailSources` breakdown
  (`workspace` / `capwatch` / `placeholder`) and a plan (`creates/updates/deletes`).
- `Peer Workspace email map built` with a non-zero `count` (proves the peer SA works). A
  `Peer directory list failed (403…)` means the DWD grant/scope/subject is wrong.
- In the tenant's GAL (or Contacts → Directory), confirm peer members now resolve by name.
- Re-run: the second run should report mostly `unchanged` and near-zero writes (proves the
  stateless hash diff works).

## Migration off the legacy projects

The two legacy projects tag their managed contacts `orgName = CAWG` (members) and
`CAWG_CADET_PARENTS_GROUPS` (parents). This module uses `CONFIG.WING` (`CA`) and
`CA_PARENTS`. Because the markers differ, the new module will not see the old contacts and
would create duplicates alongside them. Before enabling in production:

1. **Decommission** the legacy cadets (`1fJRqo…`) and seniors (`1b2JSIB…`) projects
   (remove their triggers).
2. **Clean up** their managed contacts in each tenant (delete `orgName=CAWG` /
   `CAWG_CADET_PARENTS_GROUPS`), **or** run a one-time re-tag to the new markers.
3. Then enable the module's triggers so it repopulates from live data.

The old export spreadsheets (`CAWG_seniors`, `CAWG_cadets`, `CAWG_cadets_groups`) become
unused and can be archived.

## Limitations

- The peer directory read needs a cross-tenant credential (read-only SA). This is the
  established pattern in this repo (`getImpersonatedToken_`), but it is a credential into
  the peer tenant — keep the DWD scope read-only and the key in Script Properties only.
- Suspended peer accounts are excluded from the authoritative map, so a suspended member
  falls back to their CAPWATCH email or the sentinel.
- One run is capped at `XT_MAX_WRITES_PER_RUN` and a soft time limit; a large first-time
  backfill simply finishes over subsequent runs (the stateless diff resumes cleanly).
