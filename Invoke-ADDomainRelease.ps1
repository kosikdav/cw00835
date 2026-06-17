#Requires -Modules ActiveDirectory

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

# ============================================================
# INTERNALS — do not edit below this line
# ============================================================

$script:LogFile = $null

$AttributesToCheck = @(
    'userPrincipalName',
    'mail',
    'proxyAddresses',
    'targetAddress',
    'msExchTargetAddress',
    'msExchArchiveAddress',
    'msExchShadowProxyAddresses'
)

$MultiValueAttributes = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@('proxyAddresses', 'msExchShadowProxyAddresses'),
    [System.StringComparer]::OrdinalIgnoreCase
)

$CsvColumns = @(
    'DistinguishedName', 'ObjectClass', 'sAMAccountName',
    'Attribute', 'CurrentValue', 'PlannedValue',
    'CollisionFlag', 'Action', 'ApplyChange', 'Notes'
)

# ============================================================
# HELPER FUNCTIONS
# ============================================================

function Initialize-Log {
    param([string]$Path, [string]$ModeLabel)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $Path = ".\DomainRelease_${ModeLabel}_${ts}.log"
    }
    $script:LogFile = $Path
    [string]::Empty | Set-Content -Path $Path -Encoding UTF8
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'DEBUG')]
        [string]$Level = 'INFO'
    )
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "$ts [$Level] $Message"

    switch ($Level) {
        'WARN'    { Write-Host $line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $line -ForegroundColor Red }
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        'DEBUG'   { Write-Host $line -ForegroundColor DarkGray }
        default   { Write-Host $line }
    }

    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $line -Encoding UTF8
    }
}

function Write-LogSection {
    param([string]$Title)
    $sep = '=' * 70
    Write-Log $sep
    Write-Log "=== $Title"
    Write-Log $sep
}

function Build-LDAPFilter {
    param([string[]]$Domains)

    $clauses = foreach ($domain in $Domains) {
        $d = $domain.Replace('\', '\5c').Replace('*', '\2a').Replace('(', '\28').Replace(')', '\29')
        "(userPrincipalName=*@$d)"
        "(mail=*@$d)"
        "(proxyAddresses=*@$d)"
        "(targetAddress=*@$d)"
        "(msExchTargetAddress=*@$d)"
        "(msExchArchiveAddress=*@$d)"
        "(msExchShadowProxyAddresses=*@$d)"
    }

    $orBlock    = "(|$($clauses -join ''))"
    $classBlock = "(|(objectClass=user)(objectClass=group)(objectClass=contact))"
    return "(&$classBlock$orBlock)"
}

function Get-DomainPattern {
    param([string[]]$Domains)
    $escaped = $Domains | ForEach-Object { [regex]::Escape($_) }
    $pattern = "@($($escaped -join '|'))\s*$"
    return [regex]::new(
        $pattern,
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
        [System.Text.RegularExpressions.RegexOptions]::Compiled
    )
}

function Get-PlannedValue {
    param([string]$CurrentValue, [regex]$DomainPattern, [string]$TargetDomain)
    if ($CurrentValue -match $DomainPattern) {
        $atIdx = $CurrentValue.LastIndexOf('@')
        if ($atIdx -ge 0) {
            return $CurrentValue.Substring(0, $atIdx + 1) + $TargetDomain
        }
    }
    return $null
}

function Get-RollbackCsvPath {
    param([string]$PlanPath)
    $dir  = [System.IO.Path]::GetDirectoryName($PlanPath)
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($PlanPath)
    $ext  = [System.IO.Path]::GetExtension($PlanPath)
    if ([string]::IsNullOrWhiteSpace($dir)) { $dir = '.' }
    return [System.IO.Path]::Combine($dir, "${stem}_Rollback${ext}")
}

function Get-ADCommonParams {
    $p = @{}
    if (-not [string]::IsNullOrWhiteSpace($DomainController)) {
        $p['Server'] = $DomainController
    }
    return $p
}

# ============================================================
# ANALYZE MODE
# ============================================================

function Invoke-AnalyzeMode {
    Write-LogSection "ANALYZE MODE — Domain Release Planning"
    Write-Log "Source domains : $($SourceDomains -join ', ')"
    Write-Log "Target domain  : $TargetDomain"
    Write-Log "Search bases   : $($SearchBases.Count)"
    Write-Log "Output CSV     : $PlanCsvPath"

    $ldapFilter    = Build-LDAPFilter -Domains $SourceDomains
    $domainPattern = Get-DomainPattern -Domains $SourceDomains
    $adParams      = Get-ADCommonParams
    $adProperties  = $AttributesToCheck + @('DistinguishedName', 'ObjectClass', 'sAMAccountName')

    Write-Log "LDAP filter: $ldapFilter" -Level DEBUG

    $allRows             = [System.Collections.Generic.List[PSCustomObject]]::new()
    $totalObjectsScanned = 0
    $ouCounts            = [ordered]@{}

    foreach ($ou in $SearchBases) {
        Write-LogSection "Scanning OU: $ou"
        $ouStart = Get-Date
        $ouCount = 0

        try {
            $objects = Get-ADObject `
                -LDAPFilter $ldapFilter `
                -SearchBase $ou `
                -Properties $adProperties `
                @adParams `
                -ErrorAction Stop

            foreach ($obj in $objects) {
                $ouCount++

                foreach ($attr in $AttributesToCheck) {
                    $raw = $obj.$attr
                    if ($null -eq $raw) { continue }

                    # Normalize to enumerable
                    $values = if ($raw -is [System.Collections.IEnumerable] -and $raw -isnot [string]) {
                        $raw
                    } else {
                        @($raw)
                    }

                    foreach ($val in $values) {
                        if ([string]::IsNullOrWhiteSpace($val)) { continue }

                        # Secondary client-side filter — confirms domain match
                        if ($val -notmatch $domainPattern) { continue }

                        $planned = Get-PlannedValue -CurrentValue $val -DomainPattern $domainPattern -TargetDomain $TargetDomain

                        $allRows.Add([PSCustomObject]@{
                            DistinguishedName = $obj.DistinguishedName
                            ObjectClass       = $obj.ObjectClass
                            sAMAccountName    = $obj.sAMAccountName
                            Attribute         = $attr
                            CurrentValue      = $val
                            PlannedValue      = $planned
                            CollisionFlag     = 'FALSE'
                            Action            = if ($planned) { 'Replace' } else { 'Remove' }
                            ApplyChange       = 'TRUE'
                            Notes             = ''
                        })
                    }
                }
            }
        }
        catch [Microsoft.ActiveDirectory.Management.ADServerDownException] {
            Write-Log "AD server unreachable scanning '$ou': $($_.Exception.Message)" -Level ERROR
            Write-Log $_.ScriptStackTrace -Level DEBUG
        }
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            Write-Log "OU not found: '$ou': $($_.Exception.Message)" -Level WARN
        }
        catch {
            Write-Log "Error scanning '$ou': $($_.Exception.Message)" -Level ERROR
            Write-Log $_.ScriptStackTrace -Level DEBUG
        }

        $duration = ((Get-Date) - $ouStart).TotalSeconds
        Write-Log "OU '$ou': $ouCount objects in $([math]::Round($duration,1))s"
        $totalObjectsScanned += $ouCount
        $ouCounts[$ou]        = $ouCount
    }

    # ---- Collision detection via HashSet ----
    # First pass: find all planned values that appear more than once
    $seenValues      = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $collisionValues = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($row in $allRows) {
        if ([string]::IsNullOrWhiteSpace($row.PlannedValue)) { continue }
        if ($row.Action -eq 'Remove') { continue }
        if (-not $seenValues.Add($row.PlannedValue)) {
            $null = $collisionValues.Add($row.PlannedValue)
        }
    }

    # Second pass: mark collision rows
    foreach ($row in $allRows) {
        if (-not [string]::IsNullOrWhiteSpace($row.PlannedValue) -and
            $collisionValues.Contains($row.PlannedValue)) {
            $row.CollisionFlag = 'TRUE'
            $row.PlannedValue  = 'COLLISION — MANUAL REVIEW REQUIRED'
        }
    }

    $collisionCount = ($allRows | Where-Object { $_.CollisionFlag -eq 'TRUE' }).Count

    # Write CSV
    $allRows | Export-Csv -Path $PlanCsvPath -NoTypeInformation -Encoding UTF8

    # Summary
    Write-LogSection "ANALYZE SUMMARY"
    Write-Log "Total objects scanned : $totalObjectsScanned"
    foreach ($ou in $SearchBases) {
        Write-Log "  $ou : $($ouCounts[$ou] ?? 0) objects"
    }
    Write-Log "Attribute values flagged      : $($allRows.Count)"
    if ($collisionCount -gt 0) {
        Write-Log "Collisions requiring review   : $collisionCount" -Level WARN
        Write-Log "ACTION REQUIRED: Edit '$PlanCsvPath' to resolve collision rows before running Apply mode." -Level WARN
    } else {
        Write-Log "Collisions requiring review   : 0" -Level SUCCESS
    }
    Write-Log "Plan CSV written to: $PlanCsvPath" -Level SUCCESS
}

# ============================================================
# APPLY / ROLLBACK SHARED LOGIC
# ============================================================

function Invoke-ApplyMode {
    param([switch]$IsRollback)

    $modeLabel = if ($IsRollback) { 'ROLLBACK' } else { 'APPLY' }
    $csvPath   = if ($IsRollback) { Get-RollbackCsvPath -PlanPath $PlanCsvPath } else { $PlanCsvPath }

    Write-LogSection "$modeLabel MODE"
    Write-Log "Reading plan from: $csvPath"

    if (-not (Test-Path $csvPath)) {
        Write-Log "Plan CSV not found: $csvPath" -Level ERROR
        return
    }

    $rows = Import-Csv -Path $csvPath -Encoding UTF8

    if (-not $rows -or @($rows).Count -eq 0) {
        Write-Log "CSV is empty: $csvPath" -Level WARN
        return
    }

    # Validate columns
    $existingCols = $rows[0].PSObject.Properties.Name
    $missingCols  = $CsvColumns | Where-Object { $_ -notin $existingCols }
    if ($missingCols) {
        Write-Log "CSV is missing required columns: $($missingCols -join ', ')" -Level ERROR
        return
    }

    # Generate rollback CSV before making any changes (Apply mode only)
    if (-not $IsRollback) {
        $rollbackPath = Get-RollbackCsvPath -PlanPath $PlanCsvPath
        $rollbackRows = foreach ($row in $rows) {
            [PSCustomObject]@{
                DistinguishedName = $row.DistinguishedName
                ObjectClass       = $row.ObjectClass
                sAMAccountName    = $row.sAMAccountName
                Attribute         = $row.Attribute
                CurrentValue      = $row.PlannedValue   # swapped
                PlannedValue      = $row.CurrentValue   # swapped
                CollisionFlag     = $row.CollisionFlag
                Action            = $row.Action
                ApplyChange       = $row.ApplyChange
                Notes             = $row.Notes
            }
        }
        $rollbackRows | Export-Csv -Path $rollbackPath -NoTypeInformation -Encoding UTF8
        Write-Log "Rollback CSV written to: $rollbackPath"
    }

    $adParams  = Get-ADCommonParams
    $processed = 0; $succeeded = 0; $skipped = 0; $failed = 0

    foreach ($row in $rows) {
        $processed++

        if ($row.ApplyChange -ne 'TRUE') {
            Write-Log "SKIP (ApplyChange=FALSE): $($row.DistinguishedName) [$($row.Attribute)]" -Level DEBUG
            $skipped++
            continue
        }

        if ([string]::IsNullOrWhiteSpace($row.PlannedValue) -or $row.PlannedValue -like 'COLLISION*') {
            Write-Log "SKIP (collision/blank): $($row.DistinguishedName) [$($row.Attribute)] CurrentValue='$($row.CurrentValue)'" -Level WARN
            $skipped++
            continue
        }

        $dn     = $row.DistinguishedName
        $attr   = $row.Attribute
        $oldVal = $row.CurrentValue
        $newVal = $row.PlannedValue
        $action = $row.Action

        Write-Log "Applying: $dn [$attr] '$oldVal' -> '$newVal' (Action=$action)"

        # UPN remove is not valid — check before entering try block
        if ($attr -eq 'userPrincipalName' -and $action -eq 'Remove') {
            Write-Log "Cannot remove UPN without a replacement value — skipping row." -Level WARN
            $skipped++
            continue
        }

        try {
            switch ($attr) {
                'userPrincipalName' {
                    Set-ADUser -Identity $dn -UserPrincipalName $newVal @adParams -ErrorAction Stop
                }
                'mail' {
                    if ($action -eq 'Replace') {
                        Set-ADObject -Identity $dn -Replace @{ mail = $newVal } @adParams -ErrorAction Stop
                    } else {
                        Set-ADObject -Identity $dn -Clear mail @adParams -ErrorAction Stop
                    }
                }
                { $MultiValueAttributes.Contains($_) } {
                    Set-ADObject -Identity $dn -Remove @{ $attr = $oldVal } @adParams -ErrorAction Stop
                    if ($action -eq 'Replace') {
                        Set-ADObject -Identity $dn -Add @{ $attr = $newVal } @adParams -ErrorAction Stop
                    }
                }
                default {
                    # targetAddress, msExchTargetAddress, msExchArchiveAddress
                    if ($action -eq 'Replace') {
                        Set-ADObject -Identity $dn -Replace @{ $attr = $newVal } @adParams -ErrorAction Stop
                    } else {
                        Set-ADObject -Identity $dn -Clear $attr @adParams -ErrorAction Stop
                    }
                }
            }

            # Verify the change
            $verified   = Get-ADObject -Identity $dn -Properties $attr @adParams -ErrorAction Stop
            $verifyRaw  = $verified.$attr
            $verifyVals = if ($verifyRaw -is [System.Collections.IEnumerable] -and $verifyRaw -isnot [string]) {
                @($verifyRaw)
            } else {
                @($verifyRaw)
            }

            $oldGone     = $oldVal -notin $verifyVals
            $newPresent  = ($action -eq 'Remove') -or ($newVal -in $verifyVals)

            if ($oldGone -and $newPresent) {
                Write-Log "VERIFIED: $dn [$attr]" -Level SUCCESS
                $succeeded++
            } else {
                Write-Log "VERIFY FAILED: $dn [$attr] — oldGone=$oldGone newPresent=$newPresent" -Level WARN
                $failed++
            }
        }
        catch {
            Write-Log "ERROR: $dn [$attr] — $($_.Exception.Message)" -Level ERROR
            Write-Log $_.ScriptStackTrace -Level DEBUG
            $failed++
        }
    }

    Write-LogSection "$modeLabel SUMMARY"
    Write-Log "Rows processed : $processed"
    Write-Log "Succeeded      : $succeeded" -Level $(if ($succeeded -gt 0) { 'SUCCESS' } else { 'INFO' })
    Write-Log "Skipped        : $skipped"
    Write-Log "Failed         : $failed" -Level $(if ($failed -gt 0) { 'ERROR' } else { 'SUCCESS' })
}

function Invoke-RollbackMode {
    Invoke-ApplyMode -IsRollback
}

# ============================================================
# MAIN
# ============================================================

if (-not $SourceDomains -or $SourceDomains.Count -eq 0) {
    throw "SourceDomains must contain at least one domain."
}
if ([string]::IsNullOrWhiteSpace($TargetDomain)) {
    throw "TargetDomain must be set."
}
if (-not $SearchBases -or $SearchBases.Count -eq 0) {
    throw "SearchBases must contain at least one OU."
}

Initialize-Log -Path $LogPath -ModeLabel $Mode

Write-LogSection "AD DOMAIN RELEASE — $Mode"
Write-Log "Script started   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Log "Mode             : $Mode"
Write-Log "Source domains   : $($SourceDomains -join ', ')"
Write-Log "Target domain    : $TargetDomain"
Write-Log "Search bases     : $($SearchBases -join ' | ')"
Write-Log "Plan CSV path    : $PlanCsvPath"
Write-Log "Log file         : $script:LogFile"
Write-Log "Domain controller: $(if (-not [string]::IsNullOrWhiteSpace($DomainController)) { $DomainController } else { '(default)' })"

switch ($Mode) {
    'Analyze'  { Invoke-AnalyzeMode }
    'Apply'    { Invoke-ApplyMode }
    'Rollback' { Invoke-RollbackMode }
    default    { Write-Log "Unknown mode: '$Mode'. Valid values: Analyze, Apply, Rollback" -Level ERROR }
}

Write-Log "Script completed : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level SUCCESS
