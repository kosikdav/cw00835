#Requires -Modules ActiveDirectory

# ============================================================
# CONFIGURATION — edit these before running
# ============================================================

# Domains being released from the source M365 tenant
$SourceDomains = @(
    "cez.cz",
    "cezdistribuce.cz"
)

# Domain to replace source domain addresses with
$TargetDomain = "ujvgroup.com"

# OUs to search — script processes each individually to prevent timeout
$SearchBases = @(
    "OU=aplikacni,OU=uzivatele,DC=cezdata,DC=corp",
    "OU=M365,OU=AAD,OU=Cloud,OU=skupiny,DC=cezdata,DC=corp"
)

# Folder where the log file and the two generated scripts are written
$OutputFolder = "."

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

# msExchRecipientTypeDetails bitmask -> recipient category. Values taken from
# the standard Exchange/Hybrid bit assignments. Anything not listed here
# (including objects with no value at all) is classified 'Unknown' and falls
# back to direct AD attribute edits instead of an Exchange cmdlet.
$RecipientTypeMap = @{
    [int64]1            = 'Mailbox'           # UserMailbox
    [int64]2            = 'Mailbox'           # LegacyMailbox
    [int64]4            = 'Mailbox'           # SharedMailbox
    [int64]16           = 'Mailbox'           # RoomMailbox
    [int64]32           = 'Mailbox'           # EquipmentMailbox
    [int64]64           = 'MailContact'       # MailContact
    [int64]128          = 'MailUser'          # MailUser
    [int64]256          = 'DistributionGroup' # MailUniversalDistributionGroup
    [int64]512          = 'DistributionGroup' # MailNonUniversalGroup
    [int64]1024         = 'DistributionGroup' # MailUniversalSecurityGroup
    [int64]8589934592   = 'RemoteMailbox'     # RemoteUserMailbox (hybrid)
    [int64]17179869184  = 'RemoteMailbox'     # RemoteRoomMailbox (hybrid)
    [int64]34359738368  = 'RemoteMailbox'     # RemoteEquipmentMailbox (hybrid)
    [int64]68719476736  = 'RemoteMailbox'     # RemoteSharedMailbox (hybrid)
}

# Recipient category -> Exchange Management Shell cmdlet
$ExchangeCmdletMap = @{
    'Mailbox'           = 'Set-Mailbox'
    'RemoteMailbox'     = 'Set-RemoteMailbox'
    'MailUser'          = 'Set-MailUser'
    'MailContact'       = 'Set-MailContact'
    'DistributionGroup' = 'Set-DistributionGroup'
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
    param([string]$CurrentValue, [regex]$DomainPattern, [string]$TargetDomain)
    if ($CurrentValue -match $DomainPattern) {
        $atIdx = $CurrentValue.LastIndexOf('@')
        if ($atIdx -ge 0) {
            return $CurrentValue.Substring(0, $atIdx + 1) + $TargetDomain
        }
    }
    return $null
}

function Get-ADCommonParams {
    $p = @{}
    if (-not [string]::IsNullOrWhiteSpace($DomainController)) {
        $p['Server'] = $DomainController
    }
    return $p
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

function Get-RecipientCategory {
    param($RecipientTypeDetails)
    if ($null -eq $RecipientTypeDetails) { return 'Unknown' }
    $key = [int64]$RecipientTypeDetails
    if ($RecipientTypeMap.ContainsKey($key)) { return $RecipientTypeMap[$key] }
    return 'Unknown'
}

# Returns the Exchange cmdlet parameter name for a given attribute and
# recipient category, or $null if there is no public EMS parameter for it
# (msExchTargetAddress, msExchArchiveAddress, msExchShadowProxyAddresses, and
# userPrincipalName all fall back to direct AD attribute edits).
function Get-ExchangeParameterName {
    param([string]$Category, [string]$Attribute)
    switch ($Attribute) {
        'mail'           { return 'WindowsEmailAddress' }
        'proxyAddresses' {
            return 'EmailAddresses'
        }
        'targetAddress' {
            switch ($Category) {
                'RemoteMailbox' { return 'RemoteRoutingAddress' }
                'MailContact'   { return 'ExternalEmailAddress' }
                'MailUser'      { return 'ExternalEmailAddress' }
                default         { return $null }
            }
        }
        default { return $null }
    }
}

function Get-CmdletLine {
    param([PSCustomObject]$Row)

    $dn       = ConvertTo-PSLiteral $Row.DistinguishedName
    $attr     = $Row.Attribute
    $newVal   = ConvertTo-PSLiteral $Row.PlannedValue
    $oldVal   = ConvertTo-PSLiteral $Row.CurrentValue
    $category = $Row.RecipientCategory

    # userPrincipalName is an AD/Entra auth attribute, never Exchange-managed.
    if ($attr -eq 'userPrincipalName') {
        $serverArg = if (-not [string]::IsNullOrWhiteSpace($DomainController)) { " -Server $(ConvertTo-PSLiteral $DomainController)" } else { '' }
        return "Set-ADUser -Identity $dn -UserPrincipalName $newVal$serverArg"
    }

    $exchParam = if ($category -ne 'Unknown') { Get-ExchangeParameterName -Category $category -Attribute $attr } else { $null }

    if ($exchParam) {
        $cmdlet = $ExchangeCmdletMap[$category]
        $dcArg  = if (-not [string]::IsNullOrWhiteSpace($DomainController)) { " -DomainController $(ConvertTo-PSLiteral $DomainController)" } else { '' }

        if ($exchParam -eq 'EmailAddresses') {
            if ($Row.Action -eq 'Replace') {
                return "$cmdlet -Identity $dn -EmailAddresses @{Remove=$oldVal; Add=$newVal}$dcArg"
            } else {
                return "$cmdlet -Identity $dn -EmailAddresses @{Remove=$oldVal}$dcArg"
            }
        } else {
            return "$cmdlet -Identity $dn -$exchParam $newVal$dcArg"
        }
    }

    # ---- Fallback: no EMS parameter for this attribute/category — edit AD directly ----
    $serverArg = if (-not [string]::IsNullOrWhiteSpace($DomainController)) { " -Server $(ConvertTo-PSLiteral $DomainController)" } else { '' }

    $line = switch ($attr) {
        'mail' {
            if ($Row.Action -eq 'Replace') {
                "Set-ADObject -Identity $dn -Replace @{mail=$newVal}"
            } else {
                "Set-ADObject -Identity $dn -Clear mail"
            }
        }
        { $MultiValueAttributes.Contains($_) } {
            if ($Row.Action -eq 'Replace') {
                "Set-ADObject -Identity $dn -Remove @{$attr=$oldVal} -Add @{$attr=$newVal}"
            } else {
                "Set-ADObject -Identity $dn -Remove @{$attr=$oldVal}"
            }
        }
        default {
            if ($Row.Action -eq 'Replace') {
                "Set-ADObject -Identity $dn -Replace @{$attr=$newVal}"
            } else {
                "Set-ADObject -Identity $dn -Clear $attr"
            }
        }
    }

    return "$line$serverArg"
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
    $adParams      = Get-ADCommonParams
    $adProperties  = $AttributesToCheck + @('DistinguishedName', 'ObjectClass', 'sAMAccountName', 'msExchRecipientTypeDetails')

    Write-Log "LDAP filter: $ldapFilter" -Level DEBUG

    $allRows             = [System.Collections.Generic.List[PSCustomObject]]::new()
    $totalObjectsScanned = 0
    $ouCounts            = [ordered]@{}
    $categoryCounts      = [ordered]@{}

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
                $category = Get-RecipientCategory -RecipientTypeDetails $obj.msExchRecipientTypeDetails
                $categoryCounts[$category] = ($categoryCounts[$category] ?? 0) + 1

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

                        $planned = Get-PlannedValue -CurrentValue $val -DomainPattern $domainPattern -TargetDomain $TargetDomain

                        $allRows.Add([PSCustomObject]@{
                            DistinguishedName = $obj.DistinguishedName
                            ObjectClass       = $obj.ObjectClass
                            sAMAccountName    = $obj.sAMAccountName
                            RecipientCategory = $category
                            Attribute         = $attr
                            CurrentValue      = $val
                            PlannedValue      = $planned
                            NaturalValue      = $null
                            CollisionFlag     = 'FALSE'
                            Action            = if ($planned) { 'Replace' } else { 'Remove' }
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
    Write-LogSection "COLLISION DETECTION"
    $valueGroups = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[PSCustomObject]]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $allRows) {
        if ($row.Action -ne 'Replace') { continue }
        if (-not $valueGroups.ContainsKey($row.PlannedValue)) {
            $valueGroups[$row.PlannedValue] = [System.Collections.Generic.List[PSCustomObject]]::new()
        }
        $valueGroups[$row.PlannedValue].Add($row)
    }

    $takenValues     = [System.Collections.Generic.HashSet[string]]::new([string[]]$valueGroups.Keys, [System.StringComparer]::OrdinalIgnoreCase)
    $collisionGroups = $valueGroups.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 }
    $collisionCount  = 0

    foreach ($group in $collisionGroups) {
        $naturalValue = $group.Key
        $members      = $group.Value
        Write-Log "Collision on '$naturalValue' across $($members.Count) object(s): $(($members | ForEach-Object { $_.sAMAccountName }) -join ', ')" -Level WARN

        foreach ($row in $members) {
            $unique = Get-UniqueValue -Value $naturalValue -TakenValues $takenValues
            $null = $takenValues.Add($unique)
            $row.NaturalValue  = $naturalValue
            $row.CollisionFlag = 'TRUE'
            $row.PlannedValue  = $unique
            $collisionCount++
            Write-Log "  -> $($row.DistinguishedName) [$($row.Attribute)] disabled; generated alternative: $unique" -Level WARN
        }
    }

    # ---- Drop rows that cannot be expressed as a cmdlet ----
    $validRows = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($row in $allRows) {
        if ($row.Attribute -eq 'userPrincipalName' -and $row.Action -eq 'Remove') {
            Write-Log "Cannot remove UPN without a replacement value — excluding from generated scripts: $($row.DistinguishedName)" -Level WARN
            continue
        }
        $validRows.Add($row)
    }

    Write-LogSection "DISCOVERY SUMMARY"
    Write-Log "Total objects scanned : $totalObjectsScanned"
    foreach ($ou in $SearchBases) {
        Write-Log "  $ou : $($ouCounts[$ou]) objects"
    }
    Write-Log "Recipient categories:"
    foreach ($cat in $categoryCounts.Keys) {
        Write-Log "  $cat : $($categoryCounts[$cat])"
    }
    Write-Log "Attribute values flagged    : $($validRows.Count)"
    if ($collisionCount -gt 0) {
        Write-Log "Collisions requiring review : $collisionCount" -Level WARN
    } else {
        Write-Log "Collisions requiring review : 0" -Level SUCCESS
    }

    return $validRows
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
        "# Requires: ActiveDirectory module AND an Exchange Management Shell",
        "# session (on-prem) for the Set-Mailbox / Set-RemoteMailbox / Set-MailUser /",
        "# Set-MailContact / Set-DistributionGroup lines below.",
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
        "# Requires: ActiveDirectory module AND an Exchange Management Shell",
        "# session (on-prem) for the Set-Mailbox / Set-RemoteMailbox / Set-MailUser /",
        "# Set-MailContact / Set-DistributionGroup lines below.",
        ('#' * 70),
        ''
    ))

    # msExchTargetAddress is auto-maintained by Exchange whenever targetAddress
    # is changed through an Exchange cmdlet — suppress the redundant line for
    # any DN where that's the case.
    $targetAddressEmsDNs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $Rows) {
        if ($row.Attribute -ne 'targetAddress') { continue }
        if ($row.RecipientCategory -eq 'Unknown') { continue }
        if (Get-ExchangeParameterName -Category $row.RecipientCategory -Attribute 'targetAddress') {
            $null = $targetAddressEmsDNs.Add($row.DistinguishedName)
        }
    }

    foreach ($attr in $AttributesToCheck) {
        $attrRows = @($Rows | Where-Object { $_.Attribute -eq $attr } | Sort-Object sAMAccountName, DistinguishedName)
        if ($attrRows.Count -eq 0) { continue }

        $applyLines.Add("# ===== $attr =====")
        $revertLines.Add("# ===== $attr =====")

        foreach ($row in $attrRows) {
            if ($attr -eq 'msExchTargetAddress' -and $targetAddressEmsDNs.Contains($row.DistinguishedName)) {
                $note = "# (msExchTargetAddress) auto-maintained by Exchange via the targetAddress change for $($row.sAMAccountName) — no separate cmdlet needed."
                $applyLines.Add($note)
                $revertLines.Add($note)
                continue
            }

            $forwardLine = Get-CmdletLine -Row $row

            $revertRow = [PSCustomObject]@{
                DistinguishedName = $row.DistinguishedName
                Attribute         = $row.Attribute
                CurrentValue      = $row.PlannedValue
                PlannedValue      = $row.CurrentValue
                Action            = $row.Action
                RecipientCategory = $row.RecipientCategory
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
Write-Log "Domain controller: $(if (-not [string]::IsNullOrWhiteSpace($DomainController)) { $DomainController } else { '(default)' })"

$rows = Invoke-Discover
New-CmdletScripts -Rows $rows -ApplyPath $applyPath -RevertPath $revertPath

Write-Log "Apply script written : $applyPath" -Level SUCCESS
Write-Log "Revert script written: $revertPath" -Level SUCCESS
Write-Log "Review both scripts — lines prefixed with '#' need manual review before enabling." -Level WARN
Write-Log "Script completed : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level SUCCESS
