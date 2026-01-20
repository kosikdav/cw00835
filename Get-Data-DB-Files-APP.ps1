#######################################################################################################################
# Get-Data-DB-Files-APP
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
$LogFilePrefix		= "get-data-db-files-app"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

[array]$DBReportAADPermissions  = $null
[array]$DBReportAADAppList      = $null
[array]$DBReportAADSPList       = $null

[int]$TTL = 30

#######################################################################################################################
. $IncFile_StdLogStartBlock


$AADPermCustom = Import-CSV -Path $DBFileAADPermCustom 
[array]$AADPermTemp = @()

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
$UriResource = "servicePrincipals(appId='00000003-0000-0000-c000-000000000000')"
$UriSelect = "id,appId,displayName,appRoles,oauth2PermissionScopes,resourceSpecificApplicationPermissions"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
[array]$GraphPermissions = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
[array]$srcObjArray = @($GraphPermissions.oauth2PermissionScopes,$GraphPermissions.appRoles,$GraphPermissions.resourceSpecificApplicationPermissions)   
foreach ($obj in $srcObjArray) {
    foreach ($permission in $obj) {
        $AADPermTemp += [PSCustomObject]@{
            id = $permission.id;
            value = $permission.value
        }
    }
}
$DBReportAADPermissions = $AADPermTemp + $AADPermCustom

#############################################

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
$UriResource = "applications"
$Uri = New-GraphUri -Version "beta" -Resource $UriResource
[array]$AADApplications = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
foreach ($Application in $AADApplications) {
    $DBReportAADAppList += [pscustomobject]@{
        Id				= $Application.id
        AppId			= $Application.appId
        DisplayName		= $Application.displayName
        CreatedDateTime = $Application.createdDateTime
        DeletedDateTime = $Application.deletedDateTime
        notes           = $Application.notes
    }        
}

#############################################

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
$UriResource = "servicePrincipals"
$Uri = New-GraphUri -Version "beta" -Resource $UriResource
[array]$AADServicePrincipals = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
ForEach ($ServicePrincipal in $AADServicePrincipals) {
    $DBReportAADSPList += [pscustomobject]@{
        Id										= $ServicePrincipal.id
        Enabled									= $ServicePrincipal.accountEnabled
        AppId									= $ServicePrincipal.appId
        DisplayName								= $ServicePrincipal.displayName
        AppDisplayName							= $ServicePrincipal.appDisplayName
        CreatedDateTime							= $ServicePrincipal.createdDateTime
        Desription								= $ServicePrincipal.appDescription
        PublisherName							= $ServicePrincipal.publisherName
    }
}

Export-Report "DBReportAADPermissions" -Report $DBReportAADPermissions -Path $DBFileAADPermissions
Export-Report "DBReportAADAppList" -Report $DBReportAADAppList -Path $DBFileAADAppList
Export-Report "DBReportAADSPList" -Report $DBReportAADSPList -Path $DBFileAADSPList

. $IncFile_StdLogEndBlock
