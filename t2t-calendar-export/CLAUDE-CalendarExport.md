# M365 Cross-Tenant Migration – Source Calendar JSON Export

## Project Overview

Standalone PowerShell script that connects to the **source tenant** Microsoft Graph API and exports the full calendar contents of all mailboxes scheduled for migration to JSON files. This is a pre-migration backup tool — read-only, no changes are made to any mailbox. The JSON export serves as a safety net in case calendar data needs to be referenced or reconstructed after the cross-tenant migration.

## Goals

- Connect to the source tenant using a dedicated app registration
- Enumerate all mailboxes to be migrated from an input CSV
- Export all calendar events (past and future) for each mailbox to a per-user JSON file
- Produce a summary report of what was exported
- Make no changes to any mailbox — purely read-only

## Architecture

```
/
├── CLAUDE.md                           # This file
├── README.md                           # Setup and usage guide
├── config/
│   └── config.psd1                     # Source tenant credentials and settings
├── data/
│   └── mailboxes.csv                   # Input: list of mailboxes to export
├── src/
│   ├── Connect-GraphAPI.ps1            # Auth: Client credentials flow, token management
│   ├── Get-MailboxList.ps1             # Load and validate the input mailbox list
│   ├── Export-UserCalendar.ps1         # Graph: pull all events and write to JSON
│   └── Write-SummaryReport.ps1         # Console + CSV summary of export results
├── exports/
│   └── (generated at runtime)          # Per-user JSON files
│       ├── john.doe_calendar.json
│       ├── jane.smith_calendar.json
│       └── ...
└── Start-CalendarExport.ps1            # Main entry point / orchestrator
```

## Tech Stack

- **Language:** PowerShell 7.x (cross-platform)
- **API:** Microsoft Graph API v1.0 — **source tenant**
- **Auth:** OAuth 2.0 – Client Credentials flow (app-only, no user interaction)
- **App Registration:** Source tenant Azure AD app with the following **Application** permissions:
  - `Calendars.Read`
  - `User.Read.All`
- **Config format:** PowerShell Data File (`.psd1`)
- **Output format:** UTF-8 JSON, one file per user

## Graph API Endpoints Used

| Purpose | Endpoint |
|---|---|
| List all calendar events | `GET /users/{id}/events` |
| Paginate results | follow `@odata.nextLink` until exhausted |

Use `/events` rather than `/calendarView` — `/events` returns all events including recurring series masters without expanding occurrences, which gives full fidelity backup of the original data structure. Recurring series are exported as a master event with its recurrence pattern intact, not as individual expanded occurrences.

## Coding Conventions

- All functions must have comment-based help (`<# .SYNOPSIS .DESCRIPTION .PARAMETER .EXAMPLE #>`)
- All Graph calls go through a single wrapper function `Invoke-GraphRequest` that handles token refresh, retry on 429 (throttling with `Retry-After`), and error logging
- Never hardcode credentials, tenant IDs, or secrets — all come from `config/config.psd1` or environment variables
- Use `$ErrorActionPreference = 'Stop'` and wrap all Graph calls in try/catch
- Verbose logging via `-Verbose` switch; use `Write-Verbose` throughout
- This tool is **read-only** — no POST, PATCH, or DELETE calls anywhere; if a function attempts a write operation it is a bug
- Rate limit awareness: implement exponential backoff and `Retry-After` header handling on HTTP 429

## JSON Output Format

One file per user: `exports/{sanitised-upn}_calendar.json`

The file is the raw Graph API response for all events, with a metadata wrapper added:

```json
{
  "exportMetadata": {
    "userPrincipalName": "john.doe@contoso.com",
    "displayName": "John Doe",
    "exportedAtUtc": "2025-10-01T08:00:00Z",
    "sourceTenantId": "aaaa-1111-...",
    "totalEventsExported": 142
  },
  "events": [
    {
      "id": "AAMkAA...",
      "subject": "Weekly Team Standup",
      "start": { "dateTime": "2025-10-06T09:00:00", "timeZone": "Europe/Prague" },
      "end":   { "dateTime": "2025-10-06T09:30:00", "timeZone": "Europe/Prague" },
      "isOrganizer": true,
      "isOnlineMeeting": true,
      "onlineMeeting": { "joinUrl": "https://teams.microsoft.com/..." },
      "recurrence": { ... },
      "attendees": [ ... ],
      "body": { "contentType": "html", "content": "..." },
      "location": { ... },
      "type": "seriesMaster",
      "createdDateTime": "2024-03-01T10:00:00Z",
      "lastModifiedDateTime": "2025-09-15T14:22:00Z"
    },
    ...
  ]
}
```

All fields returned by Graph are preserved as-is — do not filter or truncate any properties. The purpose is full fidelity backup.

## Configuration (`config/config.psd1`)

```powershell
@{
    # Source tenant — where mailboxes currently live before migration
    SourceTenantId          = ''
    ClientId                = ''
    ClientSecret            = ''        # Or use CertificateThumbprint

    # Export scope
    # Set both to $null to export all events with no date boundary
    ExportStartDate         = $null     # e.g. '2020-01-01' or $null for no lower bound
    ExportEndDate           = $null     # e.g. '2026-12-31' or $null for no upper bound

    # Paths
    MailboxListCsv          = './data/mailboxes.csv'
    ExportDir               = './exports/'
    SummaryReportPath       = './exports/summary.csv'
}
```

Setting `ExportStartDate` and `ExportEndDate` to `$null` exports everything with no date filter — recommended for a full backup. Dates can be set to limit scope if needed.

## Input CSV Format (`data/mailboxes.csv`)

```csv
UserPrincipalName,DisplayName
john.doe@contoso.com,John Doe
jane.smith@contoso.com,Jane Smith
```

- **UserPrincipalName** — source tenant UPN, used for Graph API calls (`/users/{UPN}/events`)
- **DisplayName** — used in the export metadata and summary report

## Main Orchestrator Usage (`Start-CalendarExport.ps1`)

```powershell
# Export all mailboxes in CSV
.\Start-CalendarExport.ps1

# Export a single mailbox
.\Start-CalendarExport.ps1 -UserPrincipalName john.doe@contoso.com

# Export with verbose logging
.\Start-CalendarExport.ps1 -Verbose
```

There is no `-DryRun` or `-Confirm` flag — this tool is read-only and safe to run at any time.

## Summary Report (`exports/summary.csv`)

One row per mailbox processed:

```
Timestamp, UserPrincipalName, DisplayName, EventsExported, OutputFile, Status, ErrorMessage
2025-10-01T08:00:00Z, john.doe@contoso.com, John Doe, 142, exports/john.doe_calendar.json, Success,
2025-10-01T08:01:14Z, jane.smith@contoso.com, Jane Smith, 87, exports/jane.smith_calendar.json, Success,
2025-10-01T08:02:03Z, bob.jones@contoso.com, Bob Jones, 0, , Failed, Mailbox not found
```

## Important Rules

- **Read-only** — this tool must never write to any mailbox or make any Graph API call other than GET
- **No date filter by default** — export everything; let config limit scope if needed, not the script logic
- **Do not expand recurring occurrences** — use `/events` not `/calendarView`; series masters must be exported with their recurrence pattern intact, not exploded into individual occurrences
- **Preserve all Graph properties** — do not filter, rename, or truncate any fields in the JSON output
- **Always write the metadata wrapper** — `exportMetadata` must be present in every output file so the file is self-describing without needing the CSV
- **Handle pagination** — `/events` pages at 10 items by default; always follow `@odata.nextLink` until exhausted
- **Overwrite existing files** — if a file already exists for a user, overwrite it; re-running the export should always produce a fresh complete file
- **Throttling** — implement `Retry-After` header handling on HTTP 429; add `Start-Sleep -Milliseconds 100` between per-mailbox calls
- **Certificate auth preferred** over client secret

## Recommended Pre-Migration Timing

Run this export **before** the migration batch starts and store the output files in a safe location (network share, SharePoint, Azure Blob). Once the source tenant mailboxes are migrated and deprovisioned the export data is the only recovery option if calendar items are found to be missing or corrupted post-migration.

## Out of Scope

- Exporting mail, contacts, tasks, or any folder other than Calendar
- Writing exported data back to any mailbox
- Converting JSON to PST — use eDiscovery export in the compliance portal if PST format is needed
- The target tenant — this script operates on the source tenant only
