# AD Domain Release Script — Claude Code Prompt

## Context

This script assists with releasing one or more domains from a source Microsoft 365 tenant in a hybrid Exchange/Entra ID environment where objects are synced from on-premises Active Directory via Azure AD Connect.

Before a domain can be removed from an M365 tenant, all AD objects that reference that domain in relevant attributes must be updated. This script automates discovery, planning, and execution of those changes, and supports rollback.

---

## What to Build

A single PowerShell script (`Invoke-ADDomainRelease.ps1`) with two primary operating modes and a rollback mode.

---

## Configuration Block

At the top of the script, define all configuration as variables so they are easy to find and modify. No parameters — this is designed to be edited and run directly.

```powershell
# ============================================================
# CONFIGURATION — edit these before running
# ============================================================

# Domains being released from the source M365 tenant
$SourceDomains = @(
    "contoso.com",
    "fabrikam.com"
)

# Domain to replace source domain addresses with
$TargetDomain = "northwind.com"

# OUs to search — script processes each individually to prevent timeout
$SearchBases = @(
    "OU=Users,DC=contoso,DC=com",
    "OU=Groups,DC=contoso,DC=com",
    "OU=Contacts,DC=contoso,DC=com",
    "OU=ServiceAccounts,DC=contoso,DC=com"
)

# Operating mode: "Analyze", "Apply", or "Rollback"
$Mode = "Analyze"

# Path to the analysis CSV (output in Analyze mode, input in Apply/Rollback mode)
$PlanCsvPath = ".\DomainReleasePlan.csv"

# Log file path (auto-timestamped if left as $null)
$LogPath = $null

# Optional: target a specific DC. Leave $null to use the default.
$DomainController = $null
```

---

## AD Attributes to Check

The following attributes must be checked on all `user`, `group`, and `contact` objects. Include all of them:

| Attribute | Notes |
|---|---|
| `userPrincipalName` | Users only |
| `mail` | All object types |
| `proxyAddresses` | All object types — check every entry in the multi-value array |
| `targetAddress` | Users and contacts (MEU/mail routing) |
| `msExchShadowProxyAddresses` | Hybrid Exchange shadow copies — may drift from proxyAddresses |
| `msExchTargetAddress` | Exchange-specific companion to targetAddress |
| `msExchArchiveAddress` | Archive mailbox addresses |

---

## Mode 1 — Analyze

### Discovery

- Use `-LDAPFilter` (not `-Filter`) to push filtering to the DC for performance. Build the LDAP filter to match any of the source domains across all relevant attributes combined with an objectClass filter for `user`, `group`, `contact`.
- Process each OU in `$SearchBases` individually in a loop. Catch and log per-OU errors without aborting the whole run.
- After retrieving objects via LDAP filter, apply a secondary client-side regex filter (compiled pattern from all source domains joined as alternation) to catch any near-matches the LDAP wildcard may have returned.
- Retrieve all attributes listed above plus `DistinguishedName`, `ObjectClass`, `sAMAccountName`.

### Collision Detection

When multiple source domains are being replaced by a single target domain, the local part (left of `@`) of two different addresses could become identical after domain substitution. The script must detect this.

Build a registry of planned values globally across all objects before writing the CSV:

- For each address being transformed (mail, UPN, each proxyAddress entry), compute the planned new value by substituting the source domain with `$TargetDomain`.
- Track all planned values in a hashtable keyed by the new value, with the list of source DNs that would produce it.
- Any new value that appears more than once is a collision.
- Where a collision is detected, **mark the row** in the CSV with a collision flag and leave the planned new value blank or set it to a sentinel (e.g. `COLLISION — MANUAL REVIEW REQUIRED`) rather than proposing an automatic value. The operator must resolve these manually by editing the CSV before running Apply mode.

### CSV Output

Write one row per attribute value that needs to change (not one row per object). Columns:

| Column | Description |
|---|---|
| `DistinguishedName` | Object DN |
| `ObjectClass` | user / group / contact |
| `sAMAccountName` | Account name for reference |
| `Attribute` | Which attribute this row covers |
| `CurrentValue` | Current value of this attribute entry |
| `PlannedValue` | Proposed new value after domain substitution — blank if collision |
| `CollisionFlag` | `TRUE` if the planned value conflicts with another object, else `FALSE` |
| `Action` | `Replace`, `Remove`, or `NoChange` — pre-populated by the script |
| `ApplyChange` | `TRUE` or `FALSE` — operator can set to FALSE to skip a specific row |
| `Notes` | Free text for operator annotations |

After writing the CSV, print a summary to the console and log:
- Total objects scanned per OU
- Total attribute values flagged
- Count of collisions requiring manual review
- Path to output CSV

---

## Mode 2 — Apply

- Read the CSV from `$PlanCsvPath`.
- Validate that the CSV has all expected columns before proceeding.
- Skip any row where `ApplyChange` is not `TRUE`.
- Skip any row where `PlannedValue` is blank or contains the collision sentinel — log a warning for each.
- For each qualifying row, apply the change to the object in AD:
  - `userPrincipalName`: `Set-ADUser -UserPrincipalName`
  - `mail`: `Set-ADObject -Replace @{mail = ...}` (works for all object types)
  - `proxyAddresses`: remove old value and add new value as separate operations using `-Remove` and `-Add` on `Set-ADObject`. Never replace the full array — only touch the affected entry to preserve other addresses.
  - `targetAddress`, `msExchTargetAddress`, `msExchArchiveAddress`, `msExchShadowProxyAddresses`: `Set-ADObject -Replace` or `-Clear` as appropriate.
  - For `Action = "Remove"` rows, remove the attribute value without adding a replacement.
- Before making each change, write the current value and the intended new value to the log.
- After making each change, verify by re-reading the attribute from AD and confirming the old value is gone and the new value (if any) is present. Log the verification result.
- Track per-object success/failure counts. On error, log the full exception and continue to the next row — do not abort.
- At the end, print and log a summary: rows processed, succeeded, skipped, failed.

### Rollback Data

Before applying any change, write a rollback CSV alongside the plan CSV (e.g. `DomainReleasePlan_Rollback.csv`) using the same schema but with `CurrentValue` and `PlannedValue` swapped. This file is used by Rollback mode to undo changes.

---

## Mode 3 — Rollback

- Read the rollback CSV produced by Apply mode.
- Apply the same change logic as Apply mode but in reverse: the `PlannedValue` column contains the value to restore, and `CurrentValue` contains the value that was set during Apply.
- Verify each change after applying.
- Log everything with the same verbosity as Apply mode.
- Rollback only rows where `ApplyChange` is `TRUE` in the rollback CSV. The operator can set rows to `FALSE` in the rollback CSV to skip specific items.

---

## Logging

All modes must log extensively. Requirements:

- Auto-generate a timestamped log file path if `$LogPath` is null: e.g. `.\DomainRelease_Analyze_20250614_143022.log`
- Use a `Write-Log` helper function that writes to both the console and the log file simultaneously.
- Log levels: `INFO`, `WARN`, `ERROR`, `SUCCESS`, `DEBUG` — include the level as a prefix on every line.
- Timestamps on every log line in `yyyy-MM-dd HH:mm:ss` format.
- Log section headers (e.g. `=== Scanning OU: OU=Users,DC=... ===`) to make the log readable when scanning for specific events.
- Log the full configuration block at script start (with domain list, search bases, target domain, mode, CSV path).
- In Apply and Rollback mode, log before and after values for every attribute change, and the verification result.
- Log per-OU timing (start time, end time, duration) to help identify slow OUs.
- On any exception, log the full `$_.Exception.Message` and `$_.ScriptStackTrace`.

---

## Implementation Notes

- Use `[System.Collections.Generic.List[PSCustomObject]]` for result accumulation — not `+=` on arrays.
- Use `[System.Collections.Generic.HashSet[string]]` for collision detection registry (case-insensitive, use `[System.StringComparer]::OrdinalIgnoreCase`).
- The LDAP filter must use `-LDAPFilter` on `Get-ADObject`, not `-Filter`, so evaluation happens on the DC.
- Use `$adCommonParams` hashtable with `Server` pre-populated if `$DomainController` is set, splat it onto all AD cmdlet calls.
- `proxyAddresses` entries are case-sensitive in AD (`SMTP:` vs `smtp:`). Preserve case when comparing and when removing/adding. Primary SMTP address has uppercase `SMTP:` prefix — if the primary address is being replaced, the replacement must also use uppercase `SMTP:`.
- When substituting domains in proxy addresses, only replace the domain portion (right of `@`). Never mutate the local part unless the operator has manually edited the `PlannedValue` in the CSV to resolve a collision.
- Use `Export-Csv -NoTypeInformation -Encoding UTF8` for all CSV output.
- Wrap all OU-level processing in `try/catch` with typed catches for `ADServerDownException` and `ADIdentityNotFoundException` at minimum.
- Include a `#Requires -Modules ActiveDirectory` statement at the top.

---

## Script Structure (suggested)

```
#Requires -Modules ActiveDirectory

# --- Configuration ---
# --- Helper Functions: Write-Log, Build-LDAPFilter, Get-DomainPattern, Test-CollisionRegistry ---
# --- Mode dispatch: switch ($Mode) { Analyze / Apply / Rollback } ---
# --- Analyze: Invoke-AnalyzeMode ---
# --- Apply: Invoke-ApplyMode ---
# --- Rollback: Invoke-RollbackMode ---
```

Keep all mode logic in dedicated functions to keep the script readable. The top-level switch just calls the right function after validating configuration.
