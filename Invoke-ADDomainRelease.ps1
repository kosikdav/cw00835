# ============================================================
# CONFIGURATION — edit these before running
# ============================================================

# Domains being released from the source M365 tenant
$SourceDomains = @(
    "contoso.com",
    "fabrikam.com"
)

# Domain to replace source domain addresses with
$TargetDomain = "ujvgroup.com"

# OUs to search — script processes each individually to prevent timeout
$SearchBases = @(
    "OU=Users,DC=contoso,DC=com",
    "OU=Groups,DC=contoso,DC=com",
    "OU=Contacts,DC=contoso,DC=com",
    "OU=ServiceAccounts,DC=contoso,DC=com"
)

# Folder where the log file and the two generated scripts are written
$OutputFolder = "."

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

# One cmdlet template per attribute. {0} = Identity (DN), {1} = new value,
# {2} = old value (only used by the multi-value attributes). Exchange
# attributes use Set-Recipient, which works across mailboxes, mail users,
# mail contacts, and mail-enabled groups without needing to know which one
# the object is. Everything else is a plain AD attribute edit.
$CmdletTemplates = @{
    'userPrincipalName'          = 'Set-ADUser -Identity {0} -UserPrincipalName {1}'
    'mail'                       = 'Set-Recipient -Identity {0} -WindowsEmailAddress {1}'
    'proxyAddresses'             = 'Set-Recipient -Identity {0} -EmailAddresses @{{Remove={2}; Add={1}}}'
    'targetAddress'              = 'Set-ADObject -Identity {0} -Replace @{{targetAddress={1}}}'
    'msExchTargetAddress'        = 'Set-ADObject -Identity {0} -Replace @{{msExchTargetAddress={1}}}'
    'msExchArchiveAddress'       = 'Set-ADObject -Identity {0} -Replace @{{msExchArchiveAddress={1}}}'
    'msExchShadowProxyAddresses' = 'Set-ADObject -Identity {0} -Remove @{{msExchShadowProxyAddresses={2}}} -Add @{{msExchShadowProxyAddresses={1}}}'
}

# ============================================================
# HELPER FUNCTIONS
# ============================================================

function Initialize-Log {
    param([string]$Path)
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
    param([string]$CurrentValue, [string]$TargetDomain)
    $atIdx = $CurrentValue.LastIndexOf('@')
    return $CurrentValue.Substring(0, $atIdx + 1) + $TargetDomain
}

function Get-UniqueValue {
    param(
        [string]$Value,
        [System.Collections.Generic.HashSet[string]]$TakenValues
    )
    $atIdx = $Value.LastIndexOf('@')
    $left  = $Value.Substring(0, $atIdx)
    $right = $Value.Substring($atIdx)
    $n = 2
    do {
        $candidate = "$left$n$right"
        $n++
    } while ($TakenValues.Contains($candidate))
    return $candidate
}

function ConvertTo-PSLiteral {
    param([string]$Value)
    return "'" + ($Value -replace "'", "''") + "'"
}

function Get-CmdletLine {
    param([PSCustomObject]$Row)
    $dn  = ConvertTo-PSLiteral $Row.DistinguishedName
    $new = ConvertTo-PSLiteral $Row.PlannedValue
    $old = ConvertTo-PSLiteral $Row.CurrentValue
    return $CmdletTemplates[$Row.Attribute] -f $dn, $new, $old
}

# ============================================================
# DISCOVERY
# ============================================================

function Invoke-Discover {
    Write-LogSection "DISCOVERY — Domain Release Scan"
    Write-Log "Source domains : $($SourceDomains -join ', ')"
    Write-Log "Target domain  : $TargetDomain"
    Write-Log "Search bases   : $($SearchBases.Count)"

    $ldapFilter    = Build-LDAPFilter -Domains $SourceDomains
    $domainPattern = Get-DomainPattern -Domains $SourceDomains
    $adProperties  = $AttributesToCheck + @('DistinguishedName', 'sAMAccountName')

    Write-Log "LDAP filter: $ldapFilter" -Level DEBUG

    $allRows             = [System.Collections.Generic.List[PSCustomObject]]::new()
    $totalObjectsScanned = 0
    $ouCounts            = [ordered]@{}

    foreach ($ou in $SearchBases) {
        Write-LogSection "Scanning OU: $ou"
        $ouStart = Get-Date
        $ouCount = 0

        try {
            $objects = Get-ADObject -LDAPFilter $ldapFilter -SearchBase $ou -Properties $adProperties -ErrorAction Stop

            foreach ($obj in $objects) {
                $ouCount++

                foreach ($attr in $AttributesToCheck) {
                    $raw = $obj.$attr
                    if ($null -eq $raw) { continue }

                    $values = if ($raw -is [System.Collections.IEnumerable] -and $raw -isnot [string]) {
                        $raw
                    } else {
                        @($raw)
                    }

                    foreach ($val in $values) {
                        if ([string]::IsNullOrWhiteSpace($val)) { continue }
                        if ($val -notmatch $domainPattern) { continue }

                        $allRows.Add([PSCustomObject]@{
                            DistinguishedName = $obj.DistinguishedName
                            sAMAccountName    = $obj.sAMAccountName
                            Attribute         = $attr
                            CurrentValue      = $val
                            PlannedValue      = Get-PlannedValue -CurrentValue $val -TargetDomain $TargetDomain
                            NaturalValue      = $null
                            CollisionFlag     = 'FALSE'
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

    # ---- Collision detection ----
    # A collision is two DIFFERENT objects ending up with the same value.
    # Multiple attributes on the SAME object sharing a value (e.g. mail and
    # userPrincipalName) is normal and must not be flagged.
    Write-LogSection "COLLISION DETECTION"
    $valueGroups = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[PSCustomObject]]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $allRows) {
        if (-not $valueGroups.ContainsKey($row.PlannedValue)) {
            $valueGroups[$row.PlannedValue] = [System.Collections.Generic.List[PSCustomObject]]::new()
        }
        $valueGroups[$row.PlannedValue].Add($row)
    }

    $takenValues    = [System.Collections.Generic.HashSet[string]]::new([string[]]$valueGroups.Keys, [System.StringComparer]::OrdinalIgnoreCase)
    $collisionCount = 0

    foreach ($group in $valueGroups.GetEnumerator()) {
        $naturalValue = $group.Key
        $dnGroups     = @($group.Value | Group-Object DistinguishedName)
        if ($dnGroups.Count -le 1) { continue }

        Write-Log "Collision on '$naturalValue' across $($dnGroups.Count) object(s): $(($dnGroups | ForEach-Object { $_.Group[0].sAMAccountName }) -join ', ')" -Level WARN

        foreach ($dnGroup in $dnGroups) {
            $unique = Get-UniqueValue -Value $naturalValue -TakenValues $takenValues
            $null = $takenValues.Add($unique)
            foreach ($row in $dnGroup.Group) {
                $row.NaturalValue  = $naturalValue
                $row.CollisionFlag = 'TRUE'
                $row.PlannedValue  = $unique
                $collisionCount++
            }
            Write-Log "  -> $($dnGroup.Name) [$(($dnGroup.Group | ForEach-Object { $_.Attribute }) -join ', ')] disabled; generated alternative: $unique" -Level WARN
        }
    }

    Write-LogSection "DISCOVERY SUMMARY"
    Write-Log "Total objects scanned : $totalObjectsScanned"
    foreach ($ou in $SearchBases) {
        Write-Log "  $ou : $($ouCounts[$ou] ?? 0) objects"
    }
    Write-Log "Attribute values flagged    : $($allRows.Count)"
    if ($collisionCount -gt 0) {
        Write-Log "Collisions requiring review : $collisionCount" -Level WARN
    } else {
        Write-Log "Collisions requiring review : 0" -Level SUCCESS
    }

    return $allRows
}

# ============================================================
# SCRIPT GENERATION
# ============================================================

function New-CmdletScripts {
    param(
        [System.Collections.Generic.List[PSCustomObject]]$Rows,
        [string]$ApplyPath,
        [string]$RevertPath
    )

    $applyName  = Split-Path -Leaf $ApplyPath
    $revertName = Split-Path -Leaf $RevertPath

    $applyLines = [System.Collections.Generic.List[string]]::new()
    $applyLines.AddRange([string[]]@(
        ('#' * 70),
        "# AD Domain Release — APPLY SCRIPT",
        "# Generated      : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "# Source domains : $($SourceDomains -join ', ')",
        "# Target domain  : $TargetDomain",
        "# Counterpart    : $revertName (run this to undo these changes)",
        "# Lines prefixed with '#' are disabled — typically a detected address",
        "# collision. Review the log file before enabling them.",
        ('#' * 70),
        ''
    ))

    $revertLines = [System.Collections.Generic.List[string]]::new()
    $revertLines.AddRange([string[]]@(
        ('#' * 70),
        "# AD Domain Release — REVERT SCRIPT",
        "# Generated      : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "# Source domains : $($SourceDomains -join ', ')",
        "# Target domain  : $TargetDomain",
        "# Counterpart    : $applyName (the script this undoes)",
        "# Lines prefixed with '#' are disabled — typically a detected address",
        "# collision. Review the log file before enabling them.",
        ('#' * 70),
        ''
    ))

    foreach ($attr in $AttributesToCheck) {
        $attrRows = @($Rows | Where-Object { $_.Attribute -eq $attr } | Sort-Object sAMAccountName, DistinguishedName)
        if ($attrRows.Count -eq 0) { continue }

        $applyLines.Add("# ===== $attr =====")
        $revertLines.Add("# ===== $attr =====")

        foreach ($row in $attrRows) {
            $forwardLine = Get-CmdletLine -Row $row

            $revertRow = [PSCustomObject]@{
                DistinguishedName = $row.DistinguishedName
                Attribute         = $row.Attribute
                CurrentValue      = $row.PlannedValue
                PlannedValue      = $row.CurrentValue
            }
            $revertLine = Get-CmdletLine -Row $revertRow

            if ($row.CollisionFlag -eq 'TRUE') {
                $applyLines.Add("# COLLISION: natural value '$($row.NaturalValue)' conflicts with another object — using generated alternative, review before enabling.")
                $applyLines.Add("# $forwardLine")
                $revertLines.Add("# COLLISION: counterpart of disabled $attr change for $($row.sAMAccountName) — review before enabling.")
                $revertLines.Add("# $revertLine")
            } else {
                $applyLines.Add($forwardLine)
                $revertLines.Add($revertLine)
            }
        }

        $applyLines.Add('')
        $revertLines.Add('')
    }

    $applyLines  | Set-Content -Path $ApplyPath  -Encoding UTF8
    $revertLines | Set-Content -Path $RevertPath -Encoding UTF8
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
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

$runTimestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath      = Join-Path $OutputFolder "DomainRelease_${runTimestamp}.log"
$applyPath    = Join-Path $OutputFolder "DomainRelease_Apply_${runTimestamp}.ps1"
$revertPath   = Join-Path $OutputFolder "DomainRelease_Revert_${runTimestamp}.ps1"

Initialize-Log -Path $logPath

Write-LogSection "AD DOMAIN RELEASE — DISCOVERY & SCRIPT GENERATION"
Write-Log "Script started   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Log "Source domains   : $($SourceDomains -join ', ')"
Write-Log "Target domain    : $TargetDomain"
Write-Log "Search bases     : $($SearchBases -join ' | ')"
Write-Log "Output folder    : $OutputFolder"
Write-Log "Log file         : $logPath"
Write-Log "Apply script     : $applyPath"
Write-Log "Revert script    : $revertPath"

$rows = Invoke-Discover
New-CmdletScripts -Rows $rows -ApplyPath $applyPath -RevertPath $revertPath

Write-Log "Apply script written : $applyPath" -Level SUCCESS
Write-Log "Revert script written: $revertPath" -Level SUCCESS
Write-Log "Review both scripts — lines prefixed with '#' need manual review before enabling." -Level WARN
Write-Log "Script completed : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level SUCCESS
