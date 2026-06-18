# AD Domain Release Script — Claude Code Prompt

## Context

This script assists with releasing one or more domains from a source Microsoft 365 tenant in a hybrid Exchange/Entra ID environment where objects are synced from on-premises Active Directory via Azure AD Connect.

Before a domain can be removed from an M365 tenant, all AD objects that reference that domain in relevant attributes must be updated. This script discovers every affected object and attribute value, then generates two ready-to-review PowerShell scripts: one that applies the changes and one that reverts them. The script itself never modifies AD — it only reads.

---

## What to Build

A single PowerShell script (`Invoke-ADDomainRelease.ps1`) that:

1. Scans AD for objects referencing the source domains.
2. Detects collisions (two objects that would end up with the same address).
3. Writes a timestamped log file with full discovery detail.
4. Generates two timestamped PS1 scripts: an Apply script and a Revert script.

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

# Folder where the log file and the two generated scripts are written
$OutputFolder = "."

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

## Discovery

- Use `-LDAPFilter` (not `-Filter`) to push filtering to the DC for performance. Build the LDAP filter to match any of the source domains across all relevant attributes combined with an objectClass filter for `user`, `group`, `contact`.
- Process each OU in `$SearchBases` individually in a loop. Catch and log per-OU errors without aborting the whole run.
- After retrieving objects via LDAP filter, apply a secondary client-side regex filter (compiled pattern from all source domains joined as alternation) to catch any near-matches the LDAP wildcard may have returned.
- Retrieve all attributes listed above plus `DistinguishedName`, `ObjectClass`, `sAMAccountName`.
- Produce one in-memory row per attribute value that needs to change (not one row per object): DistinguishedName, ObjectClass, sAMAccountName, Attribute, CurrentValue, PlannedValue, Action (`Replace` or `Remove`).

---

## Collision Detection

When multiple source domains are being replaced by a single target domain, the local part (left of `@`) of two different addresses could become identical after domain substitution. The script must detect this.

Build a registry of planned values globally across all objects before generating the scripts:

- For each address being transformed (mail, UPN, each proxyAddress entry), compute the planned new value by substituting the source domain with `$TargetDomain`.
- Track all planned values in a dictionary keyed by the new value, with the list of rows that would produce it.
- Any new value produced by more than one row is a collision.
- For every row in a collision group, generate a unique non-conflicting alternative address (append an incrementing numeric suffix to the local part, e.g. `john.doe2@target.com`, checked against all other planned and already-generated values) and mark the row as a collision.
- Collision rows are still written to both generated scripts, but as **commented-out lines** carrying the generated alternative, with a comment noting the original natural value. The operator reviews and manually uncomments after resolving the conflict.
- Log every collision group (natural value, member count, affected accounts) and every generated alternative to the log file.

---

## Script Generation

Generate two PowerShell scripts instead of executing changes or writing a CSV plan:

- **Apply script** (`DomainRelease_Apply_<timestamp>.ps1`) — one AD cmdlet per line that performs the change.
- **Revert script** (`DomainRelease_Revert_<timestamp>.ps1`) — one AD cmdlet per line that undoes the corresponding Apply line (old/new values swapped).

Rules:

- One cmdlet invocation per line. No CSV, no execution loop — these are the artifacts the operator reviews and runs themselves.
- Group lines into blocks per attribute (e.g. all `mail` changes together, then all `proxyAddresses` changes, etc.), each preceded by a `# ===== <attribute> =====` header, in the order attributes are listed above.
- Within a block, sort rows by `sAMAccountName` then `DistinguishedName` for readable, stable output.
- Collision rows: comment out the cmdlet line and add a `# COLLISION: ...` note above it explaining the natural value and that a generated alternative is in use.
- Cmdlet mapping per attribute:
  - `userPrincipalName`: `Set-ADUser -Identity <DN> -UserPrincipalName <new>`
  - `mail`: `Set-ADObject -Identity <DN> -Replace @{mail=<new>}` (or `-Clear mail` for a `Remove` action)
  - `proxyAddresses`, `msExchShadowProxyAddresses` (multi-value): `Set-ADObject -Identity <DN> -Remove @{<attr>=<old>} -Add @{<attr>=<new>}` as a single call
  - `targetAddress`, `msExchTargetAddress`, `msExchArchiveAddress`: `Set-ADObject -Identity <DN> -Replace @{<attr>=<new>}` (or `-Clear <attr>` for a `Remove` action)
  - Skip and log a warning for any `userPrincipalName` row with `Action = Remove` — UPN cannot be cleared without a replacement.
- If `$DomainController` is set, append `-Server <dc>` to every generated line.
- Both DN and attribute values must be emitted as single-quoted PowerShell string literals with embedded single quotes doubled (`'` → `''`).
- Each generated script's header comment must reference its counterpart file by name (Apply references Revert, and vice versa) plus the source/target domains and generation timestamp.

---

## Logging

All phases must log extensively. Requirements:

- Auto-generate a timestamped log file: `DomainRelease_<timestamp>.log`, sharing the same timestamp as the two generated PS1 scripts so the three files are easy to pair up.
- Use a `Write-Log` helper function that writes to both the console and the log file simultaneously.
- Log levels: `INFO`, `WARN`, `ERROR`, `SUCCESS`, `DEBUG` — include the level as a prefix on every line.
- Timestamps on every log line in `yyyy-MM-dd HH:mm:ss` format.
- Log section headers (e.g. `=== Scanning OU: OU=Users,DC=... ===`) to make the log readable when scanning for specific events.
- Log the full configuration at script start (domain list, search bases, target domain, output folder, generated file paths).
- Log per-OU timing (start time, end time, duration) to help identify slow OUs.
- Log every collision group and every generated alternative value.
- On any exception, log the full `$_.Exception.Message` and `$_.ScriptStackTrace`.

---

## Implementation Notes

- Use `[System.Collections.Generic.List[PSCustomObject]]` for result accumulation — not `+=` on arrays.
- Use `[System.Collections.Generic.HashSet[string]]` for collision/uniqueness registries (case-insensitive, use `[System.StringComparer]::OrdinalIgnoreCase`).
- The LDAP filter must use `-LDAPFilter` on `Get-ADObject`, not `-Filter`, so evaluation happens on the DC.
- Use an `$adCommonParams` hashtable with `Server` pre-populated if `$DomainController` is set, splat it onto `Get-ADObject` calls during discovery.
- `proxyAddresses` entries are case-sensitive in AD (`SMTP:` vs `smtp:`). Preserve case when generating cmdlet lines. Primary SMTP address has uppercase `SMTP:` prefix.
- When substituting domains, only replace the domain portion (right of the last `@`). Never mutate the local part except when generating a collision-resolution suffix.
- Wrap all OU-level processing in `try/catch` with typed catches for `ADServerDownException` and `ADIdentityNotFoundException` at minimum.
- Include a `#Requires -Modules ActiveDirectory` statement at the top.
- The script only ever reads from AD (`Get-ADObject` for discovery). All `Set-AD*` cmdlets appear only as generated text in the output scripts, never executed directly.

---

## Script Structure (suggested)

```
#Requires -Modules ActiveDirectory

# --- Configuration ---
# --- Helper Functions: Write-Log, Build-LDAPFilter, Get-DomainPattern, Get-PlannedValue, Get-UniqueValue, ConvertTo-PSLiteral, Get-CmdletLine ---
# --- Invoke-Discover: scan AD, detect collisions, return rows ---
# --- New-CmdletScripts: write Apply.ps1 and Revert.ps1 from rows ---
# --- Main: validate config, run discovery, generate scripts ---
```

Keep discovery and generation logic in dedicated functions to keep the script readable.
