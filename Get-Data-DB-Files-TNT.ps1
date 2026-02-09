#######################################################################################################################
# Get-Data-DB-Files-TNT
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    [string]$workloads
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "db"
$LogFilePrefix		= "get-data-db-files-tnt"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

[array]$DBReportTenantDomains   = @()
[array]$DBReportEntraExtAttributes = @()

$TTL = 30

#######################################################################################################################

. $IncFile_StdLogStartBlock

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL

$UriResource = "domains"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
[array]$TenantDomains = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
ForEach ($Domain in $TenantDomains) {
    $Email = $Intune = $OrgIdAuthentication = $OfficeCommunicationsOnline = $false

    if ($Domain.supportedServices -contains "Email") {
        $Email = $true
    }
    if ($Domain.supportedServices -contains "Intune") {
        $Intune = $true
    }
    if ($Domain.supportedServices -contains "OrgIdAuthentication") {
        $OrgIdAuthentication = $true
    }
    if ($Domain.supportedServices -contains "OfficeCommunicationsOnline") {
        $OfficeCommunicationsOnline = $true
    }

    $DomainObject = [pscustomobject]@{
        domain              = $Domain.id
        isVerified          = $Domain.isVerified
        isRoot              = $Domain.isRoot
        isInitial           = $Domain.isInitial
        isAdminManaged      = $Domain.isAdminManaged
        isDefault           = $Domain.isDefault
        authenticationType  = $Domain.authenticationType
        email              = $Email
        intune             = $Intune
        officeCommunicationsOnline = $OfficeCommunicationsOnline
        orgIdAuthentication = $OrgIdAuthentication
        supportedServices   = $Domain.supportedServices -join ";"
        passwordNotificationWindowInDays = $Domain.passwordNotificationWindowInDays
        passwordValidityPeriodInDays       = $Domain.passwordValidityPeriodInDays
    }
    $DBReportTenantDomains += $DomainObject
}

Export-Report "DBReportTenantDomains" -Report $DBReportTenantDomains -Path $DBFileTenantDomains -SortProperty "domain"

##############################################################################
# extensions
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "directoryObjects/getAvailableExtensionProperties"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
$body = @{
	isSyncedFromOnPremises = $true
} | ConvertTo-Json
$Result = Invoke-RestMethod -Uri $Uri -Method "POST" -Body $body -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -ContentType "application/json"
foreach ($ext in $Result.value.name) {	
    $DBReportEntraExtAttributes += [pscustomobject]@{
        ShortName = $ext -replace '^(?:[^_]*_){2}', ''
        Name = $ext
    }
}

Export-Report "DBReportEntraExtAttributes" -Report $DBReportEntraExtAttributes -Path $DBFileEntraExtAttributes -SortProperty "ShortName"

. $IncFile_StdLogEndBlock
