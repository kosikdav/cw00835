#######################################################################################################################
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


function Initialize-TempVars {
	$script:UPN = [string]::Empty
	$script:Department = [string]::Empty
	$script:CompanyName = [string]::Empty
	$script:Mail = [string]::Empty
	$script:MailDomain = [string]::Empty
}

function Set-TempVars {
	param (
		[parameter(Mandatory = $true)]$UserId
	)
	# main function body ##################################
	$script:CurrentUser = $script:AADUsers_DB.Item($UserId)
	If ($script:CurrentUser.mail) {
		$script:Mail = $script:CurrentUser.mail
		$script:MailDomain = $script:Mail.Split("@")[1]
	} 
	If ($script:CurrentUser.userPrincipalName) {
		$script:UPN = $script:CurrentUser.userPrincipalName
		$script:Department = $script:CurrentUser.department
		$script:CompanyName = $script:CurrentUser.companyName
	}
}


#######################################################################################################################

. $IncFile_StdLogStartBlock

$AADUsers_DB = Import-CSVtoHashDB -Path $DBFileUsersAllStd -KeyName "id"

Write-Log "Getting AAD groups report as of: $(Get-Date -Format "yyyy-MM-dd HH:MM:ss")"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "groups"
$Uriselect = "id,displayName,mailEnabled,securityEnabled,mail,onPremisesSyncEnabled,groupTypes,resourceProvisioningOptions,isAssignableToRole"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Select $Uriselect
[array]$AADGroups = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "Getting AAD groups" -ProgressDots

ForEach ($Group in $AADGroups) {
	Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
	$MemberCount, $OwnerCount = "n/a"
	$GroupIsDynamic, $GroupIsUnified, $GroupIsTeam = $false

	if (($Group | Select-Object -ExpandProperty GroupTypes) -Contains "DynamicMembership") {
		$GroupIsDynamic = $true
	}
	if (($Group | Select-Object -ExpandProperty GroupTypes) -Contains "Unified") {
		$GroupIsUnified = $true
	}
	if (($Group | Select-Object -ExpandProperty ResourceProvisioningOptions) -Contains "Team") {
		$GroupIsTeam = $true
	}

	if ($NoEnumerationGroups.Contains($Group.id) -or $Group.onPremisesSyncEnabled -or $GroupIsTeam -or $GroupIsDynamic) {
		$UriResource = "groups/$($Group.id)/members/`$count"
		$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
		$MemberCount = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ConsistencyLevel "eventual"
		$ReportGrpMem += [pscustomobject]@{
			GroupId				= $Group.id
			GroupName			= $Group.displayName
			Mail 				= $Group.mail
			MailEnabled			= $Group.mailEnabled
			SecurityEnabled		= $Group.securityEnabled
			Dynamic				= $GroupIsDynamic
			Unified				= $GroupIsUnified
			Team 				= $GroupIsTeam
			SyncedFromAD		= $Group.onPremisesSyncEnabled

			UserId				= "n/a"
			UserPrincipalName	= "n/a"
			UserDisplayName		= "n/a"
			UserMail			= "n/a"
			MailDomain			= "n/a"
			CompanyName			= "n/a"
			Department			= "n/a"
			Role				= "n/a"
		}
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
		$OwnersUPN = $GroupOwners.userPrincipalName
		
		if ($GroupMembers.Count -gt 0) {
			ForEach ($Member in $GroupMembers) {
				Initialize-TempVars
				Set-TempVars -UserId $Member.Id
				if ($Member.UserPrincipalName -in $OwnersUPN) {
					$Role = "owner"
				}
				else {
					$Role = "member"
				}
				$RecordObject = [pscustomobject]@{
					GroupId				= $Group.id
					GroupName			= $Group.displayName
					Mail 				= $Group.mail
					MailEnabled			= $Group.mailEnabled
					SecurityEnabled		= $Group.securityEnabled
					Dynamic				= $GroupIsDynamic
					Unified				= $GroupIsUnified
					Team 				= $GroupIsTeam
					SyncedFromAD		= $Group.onPremisesSyncEnabled

					UserId				= $Member.id
					UserPrincipalName	= $Member.userPrincipalName
					UserDisplayName		= $Member.displayName
					UserMail			= $Mail
					MailDomain			= $MailDomain
					CompanyName			= $CompanyName
					Department			= $Department
					Role				= $Role
				}
				$ReportGrpMem += $RecordObject
			}
		}
	}

	$ReportGrpLst += [pscustomobject]@{
		GroupId				= $Group.id
		GroupName			= $Group.displayName
		CreatedDateTime		= $Group.createdDateTime
		CreatedByAppId		= $Group.createdbyAppId
		MailEnabled			= $Group.mailEnabled
		SecurityEnabled		= $Group.securityEnabled
		AssignableToRole	= $Group.isAssignableToRole
		Mail				= $Group.Mail
		MailNickname		= $Group.MailNickname
		Dynamic				= $GroupIsDynamic
		Unified				= $GroupIsUnified
		Team 				= $GroupIsTeam
		SyncedFromAD		= $Group.onPremisesSyncEnabled
		Members				= $MemberCount
		Owners				= $OwnerCount
	}
}

Export-Report "AAD groups list report" -Report $ReportGrpLst -Path $OutputFileGrpLst
Export-Report "AAD groups membership report" -Report $ReportGrpMem -Path $OutputFileGrpMem
Export-Report "AAD groups membership report (DB folder)" -Report $ReportGrpMem -Path $DBFileGroupsMembers


#######################################################################################################################

. $IncFile_StdLogEndBlock