######################################################################################################################
# Get-AAD-Groups-Reports
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "exports"
$LogFilePrefix		= "aad-groups-reports"
$OutputFolder		= "aad-groups\reports"
$OutputFilePrefix	= "aad-groups"

$OutputFileSuffixGrpLst	= "grp-lst"
$OutputFileSuffixGrpMem	= "grp-mem"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputFileGrpLst = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixGrpLst -Ext "csv"
$OutputFileGrpMem = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixGrpMem -Ext "csv"

[System.Collections.ArrayList]$ReportGrpLst = @()
[System.Collections.ArrayList]$ReportGrpMem = @()

#######################################################################################################################

. $IncFile_StdLogStartBlock

$AADUsers_DB = Import-CSVtoHashDB -Path $DBFileUsersAllStd -KeyName "id"

Write-Log "Getting AAD groups report as of: $(Get-Date -Format "yyyy-MM-dd HH:MM:ss")"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "groups"
$Uriselect = "id,displayName,createdDateTime,mailEnabled,securityEnabled,mail,onPremisesSyncEnabled,onPremisesSamAccountName,onPremisesSecurityIdentifier,groupTypes,resourceProvisioningOptions,isAssignableToRole"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Select $Uriselect
[array]$AADGroups = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "AAD groups" -ProgressDots

ForEach ($Group in $AADGroups) {
	Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
	$MemberCount = $OwnerCount = "n/a"
	$DoNotEnumerate = $GroupIsDynamic = $GroupIsUnified = $GroupIsTeam = $false

	if ($NoEnumerationGroups.Contains($Group.id)) {
		$DoNotEnumerate = $true
	}
	if ($Group.onPremisesSyncEnabled -and ($AADGroupsReportResolveOnprem -eq $False)) {
		$DoNotEnumerate = $true
	}
	if (($Group | Select-Object -ExpandProperty GroupTypes) -Contains "DynamicMembership") {
		$GroupIsDynamic = $true
	}
	if (($Group | Select-Object -ExpandProperty GroupTypes) -Contains "Unified") {
		$GroupIsUnified = $true
	}
	if (($Group | Select-Object -ExpandProperty ResourceProvisioningOptions) -Contains "Team") {
		$GroupIsTeam = $true
	}

	$RecordObjectGroup = [pscustomobject]@{
		GroupId				= $Group.id
		GroupName			= $Group.displayName
		Mail 				= $Group.mail
		MailEnabled			= $Group.mailEnabled
		SecurityEnabled		= $Group.securityEnabled
		Dynamic				= $GroupIsDynamic
		Unified				= $GroupIsUnified
		Team 				= $GroupIsTeam
		SyncedFromAD		= $Group.onPremisesSyncEnabled

		MemberId			= "n/a"
		MemberUPN			= "n/a"
		MemberDisplayName	= "n/a"
		MemberType			= "n/a"
		MemberMail			= "n/a"
		MailDomain			= "n/a"
		CompanyName			= "n/a"
		Department			= "n/a"
		Role				= "n/a"
	}

	if ($DoNotEnumerate -or $GroupIsDynamic -or $GroupIsUnified -or $GroupIsTeam) {
		$UriResource = "groups/$($Group.id)/members/`$count"
		$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
		$MemberCount = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ConsistencyLevel "eventual"
		$ReportGrpMem += $RecordObjectGroup
	}
	else {
		$UriResource = "groups/$($Group.id)/members"
		$UriSelect = "id,displayName,userPrincipalName"
		$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Select $UriSelect
		[array]$GroupMembers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
		$MemberCount = $GroupMembers.Count

		$UriResource = "groups/$($Group.id)/owners"
		$UriSelect = "id,displayName,userPrincipalName"
		$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Select $UriSelect
		[array]$GroupOwners = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
		$OwnerCount = $GroupOwners.Count

		if ($GroupMembers.Count -gt 0) {
			ForEach ($Member in $GroupMembers) {
				if (-not $Member.id) {
					continue
				}
				$MemberObject = [pscustomobject]@{
					GroupId				= $Group.id
					GroupName			= $Group.displayName
					Mail 				= $Group.mail
					MailEnabled			= $Group.mailEnabled
					SecurityEnabled		= $Group.securityEnabled
					Dynamic				= $GroupIsDynamic
					Unified				= $GroupIsUnified
					Team 				= $GroupIsTeam
					SyncedFromAD		= $Group.onPremisesSyncEnabled
					MemberId			= $Member.id
					MemberUPN			= "n/a"
					MemberDisplayName	= $Member.displayName
					MemberType			= "n/a"
					MemberMail			= "n/a"
					MailDomain			= "n/a"
					CompanyName			= "n/a"
					Department			= "n/a"
					Role				= "member"
				}

				if ($AADUsers_DB.ContainsKey($Member.id)) {
					$CurrentUser = $AADUsers_DB.Item($Member.id)
					If ($CurrentUser.mail) {
						$MemberObject.MemberMail = $CurrentUser.mail
						$MemberObject.MailDomain = $CurrentUser.mail.Split("@")[1]
						} 
					If ($CurrentUser.userPrincipalName) {
						$MemberObject.MemberUPN	= $CurrentUser.userPrincipalName
						$MemberObject.CompanyName = $CurrentUser.companyName
						$MemberObject.Department = $CurrentUser.department
					}
				}

				$MemberObject.MemberType = $Member."@odata.type".Replace("#microsoft.graph.","")
		
				if ($Member.id -in $GroupOwners.id) {
					$MemberObject.Role = "owner"
				}
				$ReportGrpMem += $MemberObject
			}
		}
	}

	$ReportGrpLst += [pscustomobject]@{
		GroupId				= $Group.id
		GroupName			= $Group.displayName
		CreatedDateTime		= $Group.createdDateTime
		MailEnabled			= $Group.mailEnabled
		SecurityEnabled		= $Group.securityEnabled
		AssignableToRole	= $Group.isAssignableToRole
		Mail				= $Group.Mail
		MailNickname		= $Group.MailNickname
		Dynamic				= $GroupIsDynamic
		Unified				= $GroupIsUnified
		Team 				= $GroupIsTeam
		SyncedFromAD		= $Group.onPremisesSyncEnabled
		SamAccountName 		= $Group.onPremisesSamAccountName
		SID 				= $Group.onPremisesSecurityIdentifier
		Members				= $MemberCount
		Owners				= $OwnerCount
	}
}

Export-Report "AAD groups list report" -Report $ReportGrpLst -Path $OutputFileGrpLst
Export-Report "AAD groups membership report" -Report $ReportGrpMem -Path $OutputFileGrpMem
Export-Report "AAD groups membership report (DB folder)" -Report $ReportGrpMem -Path $DBFileGroupsMembers


#######################################################################################################################

. $IncFile_StdLogEndBlock