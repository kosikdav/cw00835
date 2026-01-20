#######################################################################################################################
# Get-User-License-Assignments.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    [string]$workloads
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "lic-assignment"
$LogFilePrefix		= "lic-assignment"

$OutputFolder		= "lic-assignment"
$OutputFilePrefix	= "user-lic-assignment"
$OutputFileSuffixErr = "err"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputFileAll = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Ext "csv"
$OutputFileErr = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixErr -Ext "csv"

[array]$LicenseAssignmentReport = @()
[hashtable]$AADGroups_DB = @{}

#######################################################################################################################
. $IncFile_StdLogStartBlock

$SKU_DB = Import-CSVToHashDB -Path $DBFileLicensingInfoSKUs -KeyName "skuId"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "groups"
$UriSelect = "id,displayName"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Select $UriSelect
[array]$AADGroups = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "Groups" -ProgressDots
foreach ($Group in $AADGroups) {
    $AADGroups_DB.Add($Group.id, $Group.displayName)
}

$UriResource = "users"
$UriSelect = "id,userPrincipalName,userType,displayName,companyName,department,jobTitle,licenseAssignmentStates"
$UriFilter = "userType+eq+'member'"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Select $UriSelect -Filter $UriFilter
[array]$AADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "AAD users" -ProgressDots
foreach ($User in $AADUsers) {
    foreach ($state in $user.licenseAssignmentStates) {
        if ($state.assignedByGroup -and $AADGroups_DB.ContainsKey($state.assignedByGroup)) {
            $AssignedByGroup = $AADGroups_DB[$state.assignedByGroup]
        } 
        else {
            $AssignedByGroup = $state.assignedByGroup
        }
        if ($state.skuId -and $SKU_DB.ContainsKey($state.skuId)) {
            $skuName = $SKU_DB[$state.skuId].skuDisplayName
        } 
        else {
            $skuName = $null
        }
        $ReportObject = [PSCustomObject]@{
            id = $user.id
            userPrincipalName = $user.userPrincipalName
            displayName = $user.displayName
            companyName = $user.companyName
            department = $user.department
            #jobTitle = $user.jobTitle
            skuId = $state.skuId
            skuName = $skuName
            assignedByGroup = $AssignedByGroup
            error = $state.error
            state = $state.state
        }
        $LicenseAssignmentReport += $ReportObject
    }
}
$LicenseAssignmentReportErr = $LicenseAssignmentReport | Where-Object { $_.error -ne "None" }

Export-Report "LicenseAssignmentReport" -Report $LicenseAssignmentReport -Path $OutputFileAll -SortProperty UserPrincipalName
Export-Report "LicenseAssignmentReportErr" -Report $LicenseAssignmentReportErr -Path $OutputFileErr -SortProperty UserPrincipalName

#######################################################################################################################

. $IncFile_StdLogEndBlock

