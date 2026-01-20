#######################################################################################################################
# Get-AAD-Guests-Audit
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "exports"
$LogFilePrefix		= "aad-guests-audit"

$OutputFolder		= "aad-guests\audit"
$OutputFilePrefix	= "aad-guests-audit"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -FileDateYesterday -Ext "csv"

[array]$AuditLogEventReport = $null
$start = $strYesterdayUTCStart
$end = $strYesterdayUTCEnd

#######################################################################################################################

. $IncFile_StdLogStartBlock

$AADUsers_DB = Import-CSVtoHashDB -Path $DBFileUsersAllName -KeyName "userPrincipalName"

Write-Log "Getting AAD guests audit events for: $($strYesterday)"
Write-Log "Time interval: $($strYesterdayUTCStart) - $($strYesterdayUTCEnd)"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "auditLogs/directoryAudits"
$UriFilter = "loggedByService+eq+'Invited Users'+and+activityDateTime+ge+$($start)+and+activityDateTime+le+$($end)"
$Uri = New-GraphUri -Version "beta" -Resource $UriResource -Filter $UriFilter
$AuditLogEventsInvite = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
Write-Log "Log events (Invited Users) returned: $($AuditLogEventsInvite.Count)"

foreach ($AuditLogEvent in $AuditLogEventsInvite) {
	$Upn = $AuditLogEvent.initiatedBy.user.userPrincipalName
	$InvitedUser = $null
	$InitiatedByDisplayname = $null
	$InviteRedeemSource = $null
	$Query = $null
	
	Write-Host -NoNewline "."
	if ($AuditLogEvent.additionalDetails.Count -ge 1) {
		$InvitedUser = $($AuditLogEvent.additionalDetails | Where-Object { $_.key -eq "invitedUserEmailAddress"})[0].value;
	}
	else {
		$InvitedUser = "n/a"
	}
	if ($Upn) {
		if ($AADUsers_DB.Contains($Upn)) {
			$InitiatedByDisplayname = $AADUsers_DB[$Upn].displayName
		}
	}
	
	$TargetDN = $AuditLogEvent.TargetResources.DisplayName
	if ($TargetDN -clike '*Source:*') {
		$InviteRedeemSource = $TargetDN.Substring($TargetDN.IndexOf( 'Source:', 0 ) +8 )
	}
	else {
		$InviteRedeemSource = "n/a"
	}
	
	$AuditLogEventReport += [pscustomobject]@{
		ActivityDateTime 		= $AuditLogEvent.activityDateTime;
		LoggedByService			= $AuditLogEvent.loggedByService;
		Category				= $AuditLogEvent.category;
		Result					= $AuditLogEvent.result;
		ActivityDisplayName		= $AuditLogEvent.activityDisplayName;
		InitiatedByUsr_Id 		= $AuditLogEvent.initiatedBy.user.id;
		InitiatedByUsr_UPN 		= $AuditLogEvent.initiatedBy.user.userPrincipalName;
		InitiatedByUsr_Name		= $InitiatedByDisplayname;
		InitiatedByUsr_IP		= $AuditLogEvent.initiatedBy.user.ipAddress;
		InitiatedByApp 			= $AuditLogEvent.initiatedBy.app.displayName;
		InvitedUser 			= $InvitedUser;
		InviteRedeemSource		= $InviteRedeemSource;
		TargetUser				= "";
		TargetGroupId			= "";
		TargetGroupName			= ""
	}
} 

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "auditLogs/directoryAudits"
$UriFilter = "loggedByService+eq+'Core Directory'+and+category+eq+'GroupManagement'+and+activityDateTime+ge+$($start)+and+activityDateTime+le+$($end)"
$Uri = New-GraphUri -Version "beta" -Resource $UriResource -Filter $UriFilter
$AuditLogEventsGrp = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
Write-Log "Log events (Group Management) returned: $($AuditLogEventsGrp.Count)"

foreach ($AuditLogEvent in $AuditLogEventsGrp) {
	$TargetUser = (Out-String -InputObject $AuditLogEvent.TargetResources.UserPrincipalName).Trim()
	$UPN = $AuditLogEvent.initiatedBy.user.userPrincipalName
	if (($AuditLogEvent.Category -eq 'GroupManagement') -and ($TargetUser -like "*$($GuestUPNSuffix)")) {
		if ($Upn) {
			if ($AADUsers_DB.Contains($Upn)) {
				$InitiatedByDisplayname = $AADUsers_DB[$Upn].displayName
			}
		}
		if ($AuditLogEvent.ActivityDisplayName -eq 'Add member to group') {
			$TargetGroupId 		= ($AuditLogEvent.TargetResources.ModifiedProperties[0].NewValue).Replace('"',"")
			$TargetGroupName 	= ($AuditLogEvent.TargetResources.ModifiedProperties[1].NewValue).Replace('"',"")
			$AuditLogEventReport += [pscustomobject]@{
				ActivityDateTime 		= $AuditLogEvent.activityDateTime;
				LoggedByService			= $AuditLogEvent.loggedByService;
				Category				= $AuditLogEvent.category;
				Result					= $AuditLogEvent.result;
				ActivityDisplayName		= $AuditLogEvent.activityDisplayName;
				InitiatedByUsr_Id 		= $AuditLogEvent.initiatedBy.user.id;
				InitiatedByUsr_UPN 		= $AuditLogEvent.initiatedBy.user.userPrincipalName;
				InitiatedByUsr_Name		= $InitiatedByDisplayname;
				InitiatedByUsr_IP		= $AuditLogEvent.initiatedBy.user.ipAddress;
				InitiatedByApp 			= $AuditLogEvent.initiatedBy.app.displayName;
				InvitedUser 			= "";
				InviteRedeemSource		= "";
				TargetUser				= $TargetUser;
				TargetGroupId			= $TargetGroupId;
				TargetGroupName			= $TargetGroupName
			}
		} #if 'Add member'
		if ($AuditLogEvent.ActivityDisplayName -eq 'Remove member from group') {
			$TargetGroupId 		= ($AuditLogEvent.TargetResources.ModifiedProperties[0].OldValue).Replace('"',"")
			$TargetGroupName 	= ($AuditLogEvent.TargetResources.ModifiedProperties[1].OldValue).Replace('"',"")
			$AuditLogEventReport += [pscustomobject]@{
				ActivityDateTime 		= $AuditLogEvent.activityDateTime;
				LoggedByService			= $AuditLogEvent.loggedByService;
				Category				= $AuditLogEvent.category;
				Result					= $AuditLogEvent.result;
				ActivityDisplayName		= $AuditLogEvent.activityDisplayName;
				InitiatedByUsr_Id 		= $AuditLogEvent.initiatedBy.user.id;
				InitiatedByUsr_UPN 		= $AuditLogEvent.initiatedBy.user.userPrincipalName;
				InitiatedByUsr_Name		= $InitiatedByDisplayname;
				InitiatedByUsr_IP		= $AuditLogEvent.initiatedBy.user.ipAddress;
				InitiatedByApp 			= $AuditLogEvent.initiatedBy.app.displayName;
				InvitedUser 			= "";
				InviteRedeemSource		= "";
				TargetUser				= $TargetUser;
				TargetGroupId			= $TargetGroupId;
				TargetGroupName			= $TargetGroupName
			}
		} #if 'Remove member'
	} # if *#EXT#@xyz.onmicrosoft.com
} #foreach

Export-Report "AAD guests audit report" -Report $AuditLogEventReport -Path $OutputFile -SortProperty "ActivityDateTime"

#######################################################################################################################

. $IncFile_StdLogEndBlock
