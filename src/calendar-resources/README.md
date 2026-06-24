# Calendar Resources Module

> **Syncs CAP aircraft, vehicles, and squadron locations to Google Workspace Calendar resources**

## Overview

This module reads CAPWATCH aircraft and vehicle data and creates corresponding bookable resources in Google Calendar, organized under buildings that represent airports and squadron locations. It also adds every squadron in the wing as a building, so members can select their unit's meeting location when creating calendar events — even if the unit has no assigned equipment.

### What It Creates

| Resource Type | Building | Name Format | Contact |
|---------------|----------|-------------|---------|
| Aircraft | Airport (ICAO) | `N445CP - Cessna 182T Skylane (2012)` | Maintenance Officer |
| Vehicle | Squadron (charter) | `04186 - 2019 FORD Transit 150` | Transportation Officer |
| *(none)* | Squadron (charter) | All wing squadrons with addresses | — |

## Files

| File | Location | Purpose |
|------|----------|---------|
| `UpdateResources.gs` | `src/calendar-resources/` | Main sync script |
| `AirportOverrides.json` | CAPWATCH Data folder (Drive) | FAA LIDs that the AWC API doesn't resolve |
| `AircraftOverrides.json` | CAPWATCH Data folder (Drive) | Aircraft with placeholder airport IDs in CAPWATCH |
| `AirportCache.json` | CAPWATCH Data folder (Drive) | Auto-generated cache of AWC API results (do not edit) |

## Installation

### Step 1: Add the Script

Copy `UpdateResources.gs` into your Google Apps Script project alongside the other `.gs` files.

### Step 2: Enable the Admin SDK

In your Apps Script project, go to **Services** (the `+` icon in the left panel) and add:
- **Admin SDK Directory API**

### Step 3: Upload Override Files

Upload both JSON files to your **CAPWATCH Data folder** in Google Drive (the same folder specified by `CONFIG.CAPWATCH_DATA_FOLDER_ID`):

- `AirportOverrides.json` — pre-populated with a template; add any FAA LIDs specific to your wing that don't resolve via the AWC API
- `AircraftOverrides.json` — pre-populated with a template; add any aircraft whose `AirPortID` field in CAPWATCH is set to a placeholder like `ASSIGN`

Both files are optional — the script runs cleanly if either is missing.

### Step 4: Set Up a Trigger

Add a time-driven trigger for `syncResources()`. Weekly is typically sufficient since aircraft and vehicle assignments change infrequently:

| Function | Frequency | Suggested Time |
|----------|-----------|----------------|
| `syncResources()` | Weekly | Sunday, 6:00 AM |

## Configuration

This module uses the same `CONFIG` object as the rest of the project. No additional configuration constants are required. The wing is determined automatically from `CONFIG.WING`, so only squadrons belonging to your wing are synced.

### AirportOverrides.json

Add entries for any FAA LID airport identifiers that the Aviation Weather Center API doesn't return (typically smaller general aviation fields with no ICAO code):

```json
{
  "C83": { "name": "Byron Airport",               "city": "Byron",     "state": "CA", "lat": 37.8284, "lng": -121.6258 },
  "L18": { "name": "Fallbrook Community Airpark", "city": "Fallbrook", "state": "CA", "lat": 33.3542, "lng": -117.2508 }
}
```

Keys are the FAA LID as it appears in CAPWATCH's `AirPortID` field (case-insensitive). All fields are optional except `name`.

### AircraftOverrides.json

Add entries for aircraft whose CAPWATCH `AirPortID` field is set to a placeholder value such as `ASSIGN`:

```json
{
  "N28FA": { "airportId": "KWHP", "orgid": "663" }
}
```

| Field | Description |
|-------|-------------|
| Key | Tail number, uppercase, as it appears in CAPWATCH |
| `airportId` | ICAO code of the aircraft's home airport |
| `orgid` | CAPWATCH orgid of the owning squadron (for Maintenance Officer lookup) |

## How It Works

### Buildings

The script maintains two types of buildings in Google Workspace:

**Airport buildings** (for aircraft) use the ICAO/FAA identifier as the building ID (e.g. `KSNA`, `C83`). Names and coordinates are resolved from the [Aviation Weather Center stations API](https://aviationweather.gov/api/doc/), with results cached in `AirportCache.json` to minimize API calls on subsequent runs. Airports that don't resolve via the API can be added to `AirportOverrides.json`.

**Squadron buildings** (for vehicles and event locations) use the CAP charter number as the building ID (e.g. `PCR-CA-138`). Addresses come from the `MEETING` address type in CAPWATCH's `OrgAddresses.txt`, falling back to a non-PO-box `MAIL` address if no meeting address is on file. Every squadron in the wing with a known address gets a building, regardless of whether it has any vehicles assigned.

The following unit numbers are excluded from building creation as they have no regular meeting location:
- `x-000` — Wing/Region HQ administrative placeholder
- `x-999` — Legislative Squadron (state legislators and congressional staff)

### Aircraft Resources

Only aircraft matching all of the following criteria are synced:
- Tail number matches the FAA N-number format (`N` + 1–5 digits + 0–2 letters, excluding I and O)
- Status is `active` (case-insensitive)
- Airport ID is a valid identifier (not a placeholder like `ASSIGN` or `NONE`)

The resource's `userVisibleDescription` field is populated with the Maintenance Officer's name and email from CAPWATCH DutyPosition data, making it easy for Calendar users to contact the right person when booking.

### Vehicle Resources

All rows in `Vehicles.txt` are processed. The resource name uses the CAP vehicle ID, year, make, and vehicle type. The Transportation Officer is populated in the same way as the Maintenance Officer for aircraft.

### Idempotency

The script is safe to run repeatedly. On each run it:
1. Loads the current list of buildings and resources from the Admin SDK
2. Creates anything missing
3. Updates anything that already exists (name, address, contact info)
4. Never deletes anything — removals must be done manually in the Admin Console

## Usage

### Run a Sync

```javascript
syncResources();
```

### Preview Aircraft (no API calls)

```javascript
previewAircraftResources();
```

This logs what would be created without making any changes to Google Workspace.

### Force Refresh Airport Names

Delete `AirportCache.json` from your CAPWATCH data folder, then run `syncResources()`. The script will re-fetch all airport names from the AWC API and rebuild the cache.

## Troubleshooting

**Airport building shows the identifier instead of a name** (e.g. building name is `C83` instead of `Byron Airport`)

The AWC API didn't return data for that identifier. Delete `AirportCache.json` and re-run. If it still doesn't resolve, add the airport to `AirportOverrides.json`.

**Aircraft is missing from Calendar**

Check the execution log for a `Skipping aircraft resource - invalid airportId` warning. The aircraft's `AirPortID` in CAPWATCH is likely set to a placeholder. Add an entry to `AircraftOverrides.json` with the correct airport and orgid.

**Vehicle is missing from Calendar**

Check for a `Skipping vehicle resource - no valid building` warning. The squadron has no address on file in CAPWATCH — either no `MEETING` address or all `MAIL` addresses are PO boxes. Update the address in eServices.

**`Entity Already Exists` / `Resource Not Found` errors on buildings**

This can happen if buildings were previously created with different IDs (e.g. `ORG-123` instead of `PCR-CA-138`). Delete all buildings and resources in the Admin Console under **Directory > Buildings** and **Calendar Resources**, then re-run from a clean slate.

**Admin SDK errors (403 Forbidden)**

Verify that the account running the script has super admin privileges in Google Workspace and that the Admin SDK Directory API is enabled in the Apps Script project's Services.
