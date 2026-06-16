<#
.SYNOPSIS
    Audits Active Directory objects for attributes referencing specified domains
    that would prevent domain release in Microsoft 365.

.DESCRIPTION
    Searches user, group, and contact objects across specified OUs for mail,
    proxyAddresses, userPrincipalName, and targetAddress attributes containing
    any of the specified domains. Processes each OU individually to prevent
    timeouts. Results are exported to CSV.

.NOTES
    Requires ActiveDirectory PowerShell module.
    Run from a machine with AD connectivity and sufficient read permissions.
#>

# ============================================================
# CONFIGURATION
# ============================================================

$domains = @(
    "ujv.cz",
    "egp.cz",
    "cvrez.cz",
    "icvr.cz",
    "radiomedic.cz",
    "engineeringpraha.cz",
    "skodapraha.cz",
    "vzuplzen.cz",
    "nqsafe.cz"
)

$searchBases = @(
    "DC=intra,DC=testhere,DC=cz"
)

$objectClasses = @("user", "group", "contact")

$exportPath = ".\DomainAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

$domainController = $null   # Set to specific DC hostname to target one DC, or leave $null for auto

# ============================================================
# SCRIPT
# ============================================================

#region Build LDAP filter

$domainFilters = foreach ($domain in $domains) {
    $escaped = [regex]::Escape($domain)
    "(proxyAddresses=*$escaped*)(mail=*$escaped*)(targetAddress=*$escaped*)(userPrincipalName=*$escaped*)"
}

$classFilters  = ($objectClasses | ForEach-Object { "(objectClass=$_)" }) -join ""
$domainFilters = $domainFilters -join ""
$ldapFilter    = "(&(|$classFilters)(|$domainFilters))"

#endregion

#region Common params

$adParams = @{
    LDAPFilter  = $ldapFilter
    Properties  = @("objectClass", "userPrincipalName", "mail", "proxyAddresses", "targetAddress")
    ErrorAction = "Stop"
}
if ($domainController) {
    $adParams.Server = $domainController
}

#endregion

#region Process OUs

$domainPattern = ($domains | ForEach-Object { [regex]::Escape($_) }) -join "|"
$results       = @()
$totalFound    = 0

foreach ($ou in $searchBases) {
    Write-Host "Scanning: $ou" -ForegroundColor Cyan

    try {
        $objects = Get-ADObject @adParams -SearchBase $ou

        $matched = $objects | Where-Object {
            ($_.proxyAddresses    -match $domainPattern) -or
            ($_.mail              -match $domainPattern) -or
            ($_.userPrincipalName -match $domainPattern) -or
            ($_.targetAddress     -match $domainPattern)
        }

        $count = ($matched | Measure-Object).Count
        $totalFound += $count
        Write-Host "  Found $count object(s)" -ForegroundColor $(if ($count -gt 0) { "Yellow" } else { "Green" })

        foreach ($obj in $matched) {

            # Identify which attributes are problematic and which domains they reference
            $flaggedAttributes = @()
            if ($obj.userPrincipalName -match $domainPattern) { $flaggedAttributes += "userPrincipalName" }
            if ($obj.mail              -match $domainPattern) { $flaggedAttributes += "mail" }
            if ($obj.targetAddress     -match $domainPattern) { $flaggedAttributes += "targetAddress" }
            if ($obj.proxyAddresses    -match $domainPattern) { $flaggedAttributes += "proxyAddresses" }

            foreach ($attribute in $flaggedAttributes) {
                $results += [PSCustomObject]@{
                    DistinguishedName       = $obj.DistinguishedName
                    ObjectClass             = $obj.ObjectClass
                    SearchBase              = $ou
                    UserPrincipalName       = $obj.userPrincipalName
                    Mail                    = $obj.mail
                    TargetAddress           = $obj.targetAddress
                    ProxyAddresses          = ($obj.proxyAddresses | Sort-Object) -join "; "
                    FlaggedAttribute        = $attribute
                }
            }
        }
    }
    catch [Microsoft.ActiveDirectory.Management.ADServerDownException] {
        Write-Warning "  DC unreachable for OU: $ou — skipping"
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        Write-Warning "  OU not found: $ou — skipping"
    }
    catch {
        Write-Warning "  Unexpected error on OU $ou`: $($_.Exception.Message)"
    }
}

#endregion

#region Export

if ($results.Count -gt 0) {
    $results | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nTotal objects found: $totalFound" -ForegroundColor Yellow
    Write-Host "Exported to: $exportPath" -ForegroundColor Green
}
else {
    Write-Host "`nNo objects found referencing the specified domains." -ForegroundColor Green
}

#endregion

