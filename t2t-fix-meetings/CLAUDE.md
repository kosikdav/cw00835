# M365 Cross-Tenant Migration – Meeting Remediation Tool

## Project Overview

PowerShell toolset that connects to Microsoft Graph API to remediate broken calendar meetings after a cross-tenant M365 mailbox migration. After migration, users lose organizer rights over their own meetings and Teams meeting join links become invalid. This tool automates the discovery, remediation, user communication, and reporting for all migrated mailboxes.

## Goals

- Enumerate all migrated mailboxes from an input list (CSV)
- For each mailbox, retrieve all future calendar events where `isOrganizer = true`
- Automatically remediate meetings that are safe to recreate
- Identify and flag meetings that are too complex to automate safely (leave to user)
- Generate a per-user report (PDF + TXT) in the user's language summarising what was fixed and what the user must do manually
- Email each user their report as an attachment
- BCC the IT admin/migration team mailbox on every user email
- Produce a full admin-level audit log of every action taken

## Architecture

```
/
├── CLAUDE.md                           # This file
├── README.md                           # Setup and usage guide
├── config/
│   └── config.psd1                     # Environment config (tenant, app reg, switches)
├── data/
│   └── migrated-mailboxes.csv          # Input: list of migrated users
├── templates/
│   ├── cs-CZ.psd1                      # Czech user-facing strings (editable by comms team)
│   └── en-US.psd1                      # English user-facing strings (editable by comms team)
├── src/
│   ├── Connect-GraphAPI.ps1            # Auth: Client credentials flow, token management
│   ├── Get-MigratedMailboxes.ps1       # Load and validate the input mailbox list
│   ├── Get-OrganizedMeetings.ps1       # Graph: get future events where isOrganizer = true
│   ├── Invoke-MeetingRemediation.ps1   # Remediation logic per meeting (auto vs skip decision)
│   ├── New-ReplacementMeeting.ps1      # Graph: recreate meeting under new identity
│   ├── Remove-BrokenMeeting.ps1        # Graph: cancel the old broken meeting entry
│   ├── New-UserReport.ps1              # Generate per-user PDF and TXT report from template
│   ├── Send-UserNotification.ps1       # Graph: email user with report attached, BCC admin
│   └── Write-AuditLog.ps1             # Structured admin audit log (CSV)
├── reports/
│   ├── audit.csv                       # Admin audit log (generated at runtime)
│   └── users/                          # Per-user PDF and TXT reports (generated at runtime)
│       ├── john.doe_report.pdf
│       ├── john.doe_report.txt
│       └── ...
└── Start-MeetingRemediation.ps1        # Main entry point / orchestrator
```

## Tech Stack

- **Language:** PowerShell 7.x (cross-platform)
- **API:** Microsoft Graph API v1.0
- **Auth:** OAuth 2.0 – Client Credentials flow (app-only, no user interaction)
- **PDF generation:** `PdfSharpCore` NuGet package (`Install-Package PdfSharpCore`)
- **App Registration:** Target tenant Azure AD app with the following **Application** permissions:
  - `Calendars.ReadWrite`
  - `OnlineMeetings.ReadWrite.All`
  - `Mail.Send`
  - `User.Read.All`
- **Config format:** PowerShell Data File (`.psd1`)
- **Logging:** Structured CSV audit log + color-coded console output

## Graph API Endpoints Used

| Purpose | Endpoint |
|---|---|
| List calendar events | `GET /users/{id}/calendarView?startDateTime=&endDateTime=` |
| Get event detail | `GET /users/{id}/events/{eventId}` |
| Create new event | `POST /users/{id}/events` |
| Cancel/delete event | `DELETE /users/{id}/events/{eventId}` |
| Send mail with attachment | `POST /users/{id}/sendMail` |

## Coding Conventions

- All functions must have comment-based help (`<# .SYNOPSIS .DESCRIPTION .PARAMETER .EXAMPLE #>`)
- All Graph calls go through a single wrapper function `Invoke-GraphRequest` that handles token refresh, retry on 429 (throttling with `Retry-After`), and error logging
- Never hardcode credentials, tenant IDs, or secrets — all come from `config/config.psd1` or environment variables
- Use `$ErrorActionPreference = 'Stop'` and wrap all Graph calls in try/catch
- All functions return typed PSCustomObjects, not raw API responses
- Verbose logging via `-Verbose` switch; use `Write-Verbose` throughout
- Always run in dry-run mode unless `-Confirm` is explicitly passed
- Rate limit awareness: Graph has per-app throttling limits; implement exponential backoff

## Identity Model After Migration

This is critical context for all remediation logic:

| Attribute | Source Tenant | Target Tenant |
|---|---|---|
| Primary SMTP / mail | `john.doe@contoso.com` | `john.doe@contoso.com` (**same**) |
| UPN | `john.doe@contoso.com` | `john.doe@contoso.onmicrosoft.com` (or new UPN domain) |
| AAD Object ID | `aaaa-1111-...` | `bbbb-2222-...` (**different**) |
| Exchange mailbox GUID | source GUID | **new GUID** in target |
| Teams identity | tied to source AAD OID | tied to **new** target AAD OID |

All email domains transfer to the target tenant — primary SMTP addresses are identical before and after migration. UPN is the only address attribute that differs.

## Meeting Remediation Logic

There is no detection step. Every future meeting where `isOrganizer = true` is broken by definition:

- The meeting was created under the user's old AAD Object ID in the source tenant
- After migration the user has a new AAD Object ID — Exchange no longer recognises them as the organizer
- Any Teams join link embedded in the meeting is tied to the source tenant and is invalid
- External meetings where the user is only an attendee will have `isOrganizer = false` and are skipped automatically by the Graph filter

**The Graph query filter for `Get-OrganizedMeetings.ps1`:**

```
GET /users/{TargetUPN}/calendarView
    ?startDateTime={today}
    &endDateTime={today + ScanWindowDays}
    &$filter=isOrganizer eq true
    &$select=id,subject,start,end,type,seriesMasterId,isOrganizer,isOnlineMeeting,
             onlineMeeting,recurrence,attendees,body,location,isAllDay,sensitivity
```

### What Gets Automated vs Left to User

| Meeting Type | Action | Reason |
|---|---|---|
| Single occurrence | ✅ Auto-recreate | Straightforward, no edge cases |
| Recurring series — no exceptions, no cancelled occurrences | ✅ Auto-recreate master | Safe, clean series |
| Recurring series — has modified exceptions | ❌ Skip, notify user | Exceptions would be lost; user knows what they changed |
| Recurring series — has cancelled occurrences | ❌ Skip, notify user | Cancelled occurrences would reappear |
| Room/resource bookings | ✅ Include as attendees in recreated meeting | Room auto-accept handles rebooking |

### Detecting Clean vs Complex Recurring Series

The `calendarView` endpoint returns both the series master event AND individual exceptions as separate items in the same response. Use this to avoid an extra API call per series:

- Group all returned events by `seriesMasterId`
- If any event in the group has `type = exception` or `type = exceptionDeleted` → complex → skip and flag for user
- If all events in the group have `type = occurrence` → clean → auto-recreate the master only

The instances endpoint (`/events/{id}/instances`) is **not needed** — cleanliness is determined entirely from the `calendarView` results already in memory.

When recreating a clean recurring series, recreate the **master event only** (the item where `type = seriesMaster`). Do not recreate individual occurrences — they are generated automatically from the master's recurrence pattern.

## User Communication

User communication is a first-class deliverable. Users must clearly understand what happened to their calendar and exactly what they need to do. All user-facing text lives exclusively in the language template files — no user-facing strings are hardcoded in scripts.

### Language Templates (`templates/`)

Two template files, one per language. Each is a PowerShell Data File (`.psd1`) containing all user-facing strings as named keys. The comms team can edit wording freely without touching any script.

The per-user language is determined by the `Language` column in the input CSV (`cs-CZ` or `en-US`). The script loads the matching template at runtime.

**Template structure — both files must contain identical keys:**

```powershell
# templates/cs-CZ.psd1  (Czech)
@{
    EmailSubject            = 'Vaše schůzky v kalendáři po migraci Microsoft 365'

    Greeting                = 'Vážená/ý {DisplayName},'

    Intro                   = @'
V rámci migrace Microsoft 365 bylo nutné aktualizovat některé schůzky ve vašem
kalendáři. Tento přehled vysvětluje, co jsme provedli automaticky a co případně
musíte udělat sami.
'@

    FixedHeader             = 'CO JSME OPRAVILI AUTOMATICKY'
    FixedIntro              = @'
Následující schůzky byly obnoveny vaším jménem. Účastníci obdrželi aktualizované
pozvánky. Z vaší strany není potřeba žádná akce.
'@
    FixedItem               = '✓ {Subject} ({Schedule})'
    FixedNone               = 'Všechny vaše schůzky byly aktualizovány automaticky. Není třeba nic dělat.'

    ManualHeader            = 'CO MUSÍTE UDĚLAT SAMI'
    ManualIntro             = @'
Následující schůzky nebylo možné aktualizovat automaticky, protože obsahují
upravené nebo zrušené výskyty. Prosíme, vytvořte tyto schůzky znovu sami
v aplikaci Outlook nebo Teams.
'@
    ManualItem              = '✗ {Subject} ({Schedule})'
    ManualAttendees         = '   Účastníci: {Attendees}'
    ManualAction            = '   Postup: Otevřete Outlook → Nová schůzka Teams → vytvořte tuto sérii znovu'
    ManualNone              = 'Žádné — všechny vaše schůzky byly aktualizovány automaticky.'

    ExpectHeader            = 'CO OČEKÁVAT'
    ExpectBody              = @'
• Po krátkou dobu se mohou v kalendáři zobrazovat duplicitní položky, dokud
  účastníci nepřijmou nebo neodmítnou nové pozvánky. To je normální.
• Staré položky schůzek byly zrušeny — účastníci obdrží oznámení o zrušení
  spolu s novou pozvánkou.
• Historie chatu ze schůzek Teams před migrací není přenositelná. Předchozí
  vlákna jsou stále dostupná v Teams v části Historie chatu.
• V případě dotazů kontaktujte IT helpdesk: {HelpDeskEmail}
'@
}

# templates/en-US.psd1  (English)
@{
    EmailSubject            = 'Your calendar meetings after the M365 migration'

    Greeting                = 'Dear {DisplayName},'

    Intro                   = @'
As part of the Microsoft 365 migration, some of your calendar meetings needed
to be updated. This report explains what we did automatically and what — if
anything — you need to do yourself.
'@

    FixedHeader             = 'WHAT WE FIXED AUTOMATICALLY'
    FixedIntro              = @'
The following meetings have been recreated on your behalf. Attendees have
received updated invitations. No action is needed from you for these.
'@
    FixedItem               = '✓ {Subject} ({Schedule})'
    FixedNone               = 'Nothing — all your meetings have been updated automatically.'

    ManualHeader            = 'WHAT YOU NEED TO DO YOURSELF'
    ManualIntro             = @'
The following meetings could not be updated automatically because they have
customised or cancelled occurrences that cannot be safely reproduced. Please
recreate these meetings yourself in Outlook or Teams.
'@
    ManualItem              = '✗ {Subject} ({Schedule})'
    ManualAttendees         = '   Attendees: {Attendees}'
    ManualAction            = '   Action: Open Outlook → New Teams Meeting → recreate this series'
    ManualNone              = 'Nothing — all your meetings have been updated automatically.'

    ExpectHeader            = 'WHAT TO EXPECT'
    ExpectBody              = @'
• You may see duplicate calendar entries for a short period while attendees
  accept or decline the new invitations. This is normal.
• Old meeting entries have been cancelled — attendees will receive a cancellation
  notice alongside the new invitation.
• Teams meeting chat history from before the migration cannot be recovered.
  Previous chat threads are still accessible in Teams under Chat history.
• If you have questions, contact the IT helpdesk at {HelpDeskEmail}
'@
}
```

**Template variable substitution:** the script replaces `{DisplayName}`, `{Subject}`, `{Schedule}`, `{Attendees}`, and `{HelpDeskEmail}` at render time. All other text is owned by the comms team. Template keys must never be removed or renamed without updating the script.

### Report Generation (`New-UserReport.ps1`)

- Loads the correct language template based on the user's `Language` value from the CSV
- Builds the report content by substituting variables into template strings
- Renders identical content to both PDF (`PdfSharpCore`) and TXT (plain UTF-8)
- Saves to `reports/users/{sanitised-upn}_report.pdf` and `reports/users/{sanitised-upn}_report.txt`
- Generated for every user on every run including dry-run (based on what would be done), so output can be reviewed before committing

### Email (`Send-UserNotification.ps1`)

- **To:** user's `PrimarySMTP` from the CSV
- **BCC:** `AdminNotificationEmail` from config
- **Subject:** `EmailSubject` from the language template
- **Body:** same content as the report (do not just say "see attached")
- **Attachments:** both PDF and TXT files, Base64-encoded via Graph `sendMail`
- **Sent from:** `SenderMailbox` in config — not the user's own mailbox
- One email per user, sent after all their meetings have been processed — never one email per meeting

## Configuration (`config/config.psd1`)

```powershell
@{
    # Target tenant (where mailboxes now live)
    TargetTenantId          = ''
    ClientId                = ''
    ClientSecret            = ''        # Or use CertificateThumbprint

    # How far ahead to scan for future meetings (days)
    ScanWindowDays          = 180

    # Communication
    SenderMailbox           = 'migration@contoso.com'
    AdminNotificationEmail  = 'it-migration@contoso.com'
    HelpDeskEmail           = 'helpdesk@contoso.com'

    # Behaviour switches
    SendNotifications       = $false
    DeleteOldMeetings       = $false
    DryRun                  = $true

    # Paths
    MailboxListCsv          = './data/migrated-mailboxes.csv'
    AuditLogPath            = './reports/audit.csv'
    UserReportDir           = './reports/users/'
    TemplateDir             = './templates/'
}
```

## Input CSV Format (`data/migrated-mailboxes.csv`)

```csv
TargetUPN,DisplayName,PrimarySMTP,Language,MigrationDate
john.doe@contoso.onmicrosoft.com,Jan Novák,john.doe@contoso.com,cs-CZ,2025-10-15
jane.smith@contoso.onmicrosoft.com,Jane Smith,jane.smith@contoso.com,en-US,2025-10-15
```

- **TargetUPN** — used for all Graph API calls (`/users/{TargetUPN}/...`)
- **PrimarySMTP** — used as the To address for the user email
- **DisplayName** — used in the report greeting (`{DisplayName}` substitution)
- **Language** — must match a template filename (`cs-CZ` or `en-US`); script throws if no matching template exists

## Main Orchestrator Usage (`Start-MeetingRemediation.ps1`)

```powershell
# Dry run – generate reports only, no calendar changes, no emails sent
.\Start-MeetingRemediation.ps1 -DryRun

# Apply remediation to specific user, send notification
.\Start-MeetingRemediation.ps1 -UserPrincipalName john.doe@contoso.onmicrosoft.com -Confirm -SendNotifications

# Apply remediation to all mailboxes in CSV, send notifications
.\Start-MeetingRemediation.ps1 -Confirm -SendNotifications

# Apply remediation only, hold notifications for later
.\Start-MeetingRemediation.ps1 -Confirm
```

## Admin Audit Log Format (`reports/audit.csv`)

```
Timestamp, TargetUPN, EventId, EventSubject, EventType, Action, Status, ErrorMessage
2025-10-20T09:15:00Z, john.doe@contoso.onmicrosoft.com, AAMkAA..., Weekly Sync, singleInstance, Recreated, Success,
2025-10-20T09:15:01Z, john.doe@contoso.onmicrosoft.com, AAMkAA..., Weekly Sync, singleInstance, OldCancelled, Success,
2025-10-20T09:15:02Z, john.doe@contoso.onmicrosoft.com, BBMkAA..., 1:1 with Sarah, seriesMaster, Skipped, Success, Has exceptions
2025-10-20T09:15:10Z, john.doe@contoso.onmicrosoft.com, -, -, -, EmailSent, Success,
```

## Important Rules

- **NEVER cancel/delete a meeting without first successfully creating its replacement** — verify the new event ID exists before touching the old one
- **Dry-run is the default** — all write operations (create, cancel, email) require explicit `-Confirm` flag
- **Dry-run still generates reports** — so the comms team can review exact user output before committing
- **Recreate master events only** for clean recurring series — never recreate individual occurrences
- **One email per user** — aggregate all results before sending; never send mid-processing
- **No user-facing strings in scripts** — all user-visible text comes from template files only
- **Template validation on startup** — verify all required keys exist in the loaded template before processing any mailbox; fail fast with a clear error if a key is missing
- **Throttling:** implement `Retry-After` header handling on HTTP 429; add `Start-Sleep -Milliseconds 100` between per-event calls as a baseline
- **Large tenants:** process mailboxes using `ForEach-Object -Parallel` with a `-ThrottleLimit` of 5-10
- **Certificate auth preferred** over client secret for production runs

## Out of Scope

- Historical (past) meetings — only future meetings (`start >= today`) are remediated
- Meetings where `isOrganizer = false` — user is an attendee only, nothing is broken
- Recurring series with modified exceptions or cancelled occurrences — flagged for user to recreate manually
- Teams meeting chat history — cannot be migrated, acknowledged in user report
- Adding additional languages — new language requires a new `{locale}.psd1` template file with all required keys, no script changes needed
- SharePoint / OneDrive migration
- Mail flow or distribution group remediation
