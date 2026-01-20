#######################################################################################################################
# Get-AAD-Groups-Reports-OU
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			    = "exports"
$LogFilePrefix		    = "aad-groups-reports-OU"
$OutputFileSuffixList	= "list"
$OutputFileSuffixMem	= "members"

$OutputFolderAzSub   	= "azure\reports"
$OutputFilePrefixAzSub	= "az-sub-groups"

$OutputFolderADO   	    = "aad-groups\ou-devops"
$OutputFilePrefixADO	= "ado-groups"

$OutputFolderEXO   	    = "aad-groups\ou-exo"
$OutputFilePrefixEXO	= "exo-groups"

$OutputFolderLic   	    = "aad-groups\ou-lic"
$OutputFilePrefixLic	= "lic-groups"

$OutputFolderIntn   	= "aad-groups\ou-intune"
$OutputFilePrefixIntn	= "intune-groups"

$OutputFolderMIP   	    = "aad-groups\ou-mip"
$OutputFilePrefixMIP	= "mip-groups"

$OutputFolderPwr   	    = "aad-groups\ou-pwr"
$OutputFilePrefixPwr    = "pwr-groups"

$OutputFolderSPO   	        = "aad-groups\ou-spo"
$OutputFilePrefixSPO	    = "spo-groups"

$OutputFolderAzOwners	= "azure\reports"
$OutputFilePrefixAzOwners 	= "az-owners-groups"

$ADCredentialPath = $aad_grp_mgmt_cred

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputFileAzSub = New-OutputFile -RootFolder $ROF -Folder $OutputFolderAzSub -Prefix $OutputFilePrefixAzSub -Ext "csv"
$OutputFileADO  = New-OutputFile -RootFolder $ROF -Folder $OutputFolderADO -Prefix $OutputFilePrefixADO -Ext "csv"
$OutputFileEXO  = New-OutputFile -RootFolder $ROF -Folder $OutputFolderEXO -Prefix $OutputFilePrefixEXO -Ext "csv"
$OutputFileLicList  = New-OutputFile -RootFolder $ROF -Folder $OutputFolderLic -Prefix $OutputFilePrefixLic -Suffix $OutputFileSuffixList -Ext "csv"
$OutputFileLicMem   = New-OutputFile -RootFolder $ROF -Folder $OutputFolderLic -Prefix $OutputFilePrefixLic -Suffix $OutputFileSuffixMem -Ext "csv"
$OutputFileIntn = New-OutputFile -RootFolder $ROF -Folder $OutputFolderIntn -Prefix $OutputFilePrefixIntn -Ext "csv"
$OutputFileMIP  = New-OutputFile -RootFolder $ROF -Folder $OutputFolderMIP -Prefix $OutputFilePrefixMIP -Ext "csv"
$OutputFilePwrList  = New-OutputFile -RootFolder $ROF -Folder $OutputFolderPwr -Prefix $OutputFilePrefixPwr -Suffix $OutputFileSuffixList -Ext "csv"
$OutputFilePwrMem   = New-OutputFile -RootFolder $ROF -Folder $OutputFolderPwr -Prefix $OutputFilePrefixPwr -Suffix $OutputFileSuffixMem -Ext "csv"
$OutputFileSPOList  = New-OutputFile -RootFolder $ROF -Folder $OutputFolderSPO -Prefix $OutputFilePrefixSPO -Suffix $OutputFileSuffixList -Ext "csv"
$OutputFileSPOMem   = New-OutputFile -RootFolder $ROF -Folder $OutputFolderSPO -Prefix $OutputFilePrefixSPO -Suffix $OutputFileSuffixMem -Ext "csv"
$OutputFileGrpMemAzOwners = New-OutputFile -RootFolder $ROF -Folder $OutputFolderAzOwners -Prefix $OutputFilePrefixAzOwners -Suffix $OutputFileSuffixMem -Ext "csv"

#######################################################################################################################

. $IncFile_StdLogStartBlock

[System.Collections.ArrayList]$GroupReportPwrMem = @()
[System.Collections.ArrayList]$GroupReportSPOMem = @()
[System.Collections.ArrayList]$GroupReportLicMem = @()
[System.Collections.ArrayList]$ReportGrpAzOwners = @()

Write-Log "AD credential file:  $($ADCredentialPath)"
Write-Log "AzSub OU group list:  $($OutputFileAzSub)"
Write-Log "ADO OU group list:    $($OutputFileADO)"
write-Log "EXO OU group list:    $($OutputFileEXO)"
write-Log "Lic OU group list:    $($OutputFileLicList)"
write-Log "Lic OU group members: $($OutputFileLicMem)"
write-Log "Intune OU group list: $($OutputFileIntn)"
write-Log "MIP OU group list:    $($OutputFileMIP)"
write-Log "Pwr OU group list:    $($OutputFilePwrList)"
write-Log "Pwr OU group members: $($OutputFilePwrMem)"
write-Log "SPO OU group list:    $($OutputFileSPOList)"
Write-Log "SPO OU group members: $($OutputFileSPOMem)"

if (-not $interactiveRun) {
    Write-Log "AD credential file: $($ADCredentialPath)"
    $ADCredential = Import-Clixml -Path $ADCredentialPath
}

$AADUsers_DB = Import-CSVtoHashDB -Path $DBFileUsersMemMin -KeyName "id"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

###################################################################
# Get Azure subscripiton owners
$UriResource = "groups"
$Uriselect = "id,displayName,mailEnabled,securityEnabled,mail,onPremisesSyncEnabled,groupTypes,resourceProvisioningOptions"
$UriFilter = "startswith(displayName,'CEZ_AZURE_')"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect -Filter $UriFilter -Count
[array]$GroupsAzOwners = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ConsistencyLevel "eventual"
foreach ($group in $GroupsAzOwners) {
    if ($Group.displayName -like "*_Owner") {
        Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
        $UriResource = "groups/$($Group.id)/members"
        $UriSelect = "id,displayName,userPrincipalName"
        $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
        [array]$Members = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON 
        if ($Members.Count -gt 0) {
            foreach ($member in $members) {
                $CurrentUser = $null
                if ($AADUsers_DB.ContainsKey($member.id)) {
                    $CurrentUser = $AADUsers_DB.Item($member.id)
                }
                $ReportGrpAzOwners += [pscustomobject]@{
                    GroupId				= $Group.id; 
                    GroupName			= $Group.displayName;
                    MailEnabled			= $Group.mailEnabled;
                    SecurityEnabled		= $Group.securityEnabled;
                    SyncedFromAD		= $Group.onPremisesSyncEnabled;
                    UserId				= $Member.id;
                    UserPrincipalName	= $Member.userPrincipalName;
                    UserDisplayName		= $Member.displayName;
                    UserMail			= $CurrentUser.mail;
                    CompanyName			= $CurrentUser.companyName;
                    Department			= $CurrentUser.department	
                }
            }
        }
    }
}
Export-Report "AAD groups Azure owners report" -Report $ReportGrpAzOwners -Path $OutputFileGrpMemAzOwners

###################################################################

$GroupReportADO		= Get-AADGroupListByOnpremOU -Credential $ADCredential -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -OU $OU_ADO
$GroupReportAzSub 	= Get-AADGroupListByOnpremOU -Credential $ADCredential -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -OU $OU_AzSub

###################################################################
# Get Pwr groups
$GroupReportPwrList	= Get-AADGroupListByOnpremOU -Credential $ADCredential -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -OU $OU_Pwr
foreach ($group in $GroupReportPwrList) {
    $members = Get-GroupMembersFromGraphById -id $group.AAD_id -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
    write-host "Group: $($group.AAD_DisplayName) $($group.AAD_id) has $($members.Count) members"
    foreach ($member in $members) {
        $memberObject = [pscustomobject]@{
            GroupId = $group.AAD_id;
            GroupName = $group.AAD_DisplayName;
            UserId = $member.id;
            DisplayName = $member.displayName;
            UserPrincipalName = $member.userPrincipalName;
            Mail = $member.mail;
            JobTitle = $member.jobTitle
        }
        $GroupReportPwrMem += $memberObject
    }
}

###################################################################
# Get SPO groups
$GroupReportSPOList = Get-AADGroupListByOnpremOU -Credential $ADCredential -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -OU $OU_SPO
foreach ($group in $GroupReportSPOList) {
    $members = Get-GroupMembersFromGraphById -id $group.AAD_id -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
    write-host "Group: $($group.AAD_DisplayName) $($group.AAD_id) has $($members.Count) members"
    foreach ($member in $members) {
        $memberObject = [pscustomobject]@{
            GroupId = $group.AAD_id;
            GroupName = $group.AAD_DisplayName;
            UserId = $member.id;
            DisplayName = $member.displayName;
            UserPrincipalName = $member.userPrincipalName;
            Mail = $member.mail;
            JobTitle = $member.jobTitle
        }
        $GroupReportSPOMem += $memberObject
    }
}

###################################################################
# Get Lic groups
$GroupReportLICList = Get-AADGroupListByOnpremOU -Credential $ADCredential -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -OU $OU_Lic
foreach ($group in $GroupReportLICList) {
    $members = Get-GroupMembersFromGraphById -id $group.AAD_id -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
    write-host "Group: $($group.AAD_DisplayName) $($group.AAD_id) has $($members.Count) members"
    foreach ($member in $members) {
        $memberObject = [pscustomobject]@{
            GroupId = $group.AAD_id;
            GroupName = $group.AAD_DisplayName;
            UserId = $member.id;
            DisplayName = $member.displayName;
            UserPrincipalName = $member.userPrincipalName;
            Mail = $member.mail;
            JobTitle = $member.jobTitle
        }
        $GroupReportLicMem += $memberObject
    }
}

Export-Report "ADO AAD groups list report" -Report $GroupReportADO -Path $OutputFileADO
Export-Report "AzSub AAD groups list report" -Report $GroupReportAzSub -Path $OutputFileAzSub
Export-Report "Pwr AAD groups list report" -Report $GroupReportPwrList -Path $OutputFilePwrList
Export-Report "Pwr AAD groups members report" -Report $GroupReportPwrMem -Path $OutputFilePwrMem
Export-Report "SPO AAD groups list report" -Report $GroupReportSPOList -Path $OutputFileSPOList
Export-Report "SPO AAD groups members report" -Report $GroupReportSPOMem -Path $OutputFileSPOMem
#Export-Report "EXO AAD groups list report" -Report $GroupReportEXO -Path $OutputFileEXO
Export-Report "Lic AAD groups list report" -Report $GroupReportLICList -Path $OutputFileLicList
Export-Report "Lic AAD groups members report" -Report $GroupReportLicMem -Path $OutputFileLicMem
#Export-Report "Intune AAD groups list report" -Report $GroupReportIntn -Path $OutputFileIntn
#Export-Report "MIP AAD groups list report" -Report $GroupReportMIP -Path $OutputFileMIP



#######################################################################################################################

. $IncFile_StdLogEndBlock
