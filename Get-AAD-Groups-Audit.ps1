#######################################################################################################################
# Get-AAD-Groups-Audit
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "exports"
$LogFilePrefix		= "aad-groups-audit"

$OutputFolder		= "aad-groups\audit"
$OutputFilePrefix	= "aad-groups-audit"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -FileDateYesterday -Ext "csv"

[array]$AuditLogEventReport = @()
$start = $strYesterdayUTCStart
$end = $strYesterdayUTCEnd

#######################################################################################################################

. $IncFile_StdLogStartBlock

$AADGroupsDB = Import-CSVtoHashDB -Path $DBFileGroupsAllMin -KeyName "id"
$AADUsersDB = Import-CSVtoHashDB -Path $DBFileUsersAllMin -KeyName "id"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "auditLogs/directoryAudits"
$UriFilter = "loggedByService+eq+'Core Directory'+and+category+eq+'GroupManagement'+and+activityDateTime+ge+$($start)+and+activityDateTime+le+$($end)"
$Uri = New-GraphUri -Version "beta" -Resource $UriResource -Filter $UriFilter
$AuditLogEventsGrp = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
Write-Log "Audit log events (Group Management) found: $($AuditLogEventsGrp.Count)"

$AuditLogEventsGrpCounter = 0
foreach ($AuditLogEvent in $AuditLogEventsGrp) {
	if ($AuditLogEvent.initiatedBy.app.displayName -eq "Microsoft Approval Management") {
		Continue
	}
	if ($AuditLogEvent.activityDisplayName -eq "Update group") {
		Continue
	}
	foreach ($targetResource in $AuditLogEvent.targetResources) {
		if ($targetResource.type -eq "User") {
			$PSObj_TargetUserUPN 	= $targetResource.userPrincipalName
			$PSObj_TargetUserId 	= $targetResource.id
		}
		if ($targetResource.type -eq "Group") {
			$PSObj_TargetGroupId 	= $targetResource.id
		}
	}
	$PSObj_InitiatedByUsr_Name = $null
	$PSObj_TargetUserName = $null
	$PSObj_TargetUserMail = $null
	$PSObj_TargetUserType = $null

	$PSObj_ActivityDateTime 	= $AuditLogEvent.activityDateTime;
	$PSObj_LoggedByService		= $AuditLogEvent.loggedByService;
	$PSObj_Category				= $AuditLogEvent.category;
	$PSObj_Result				= $AuditLogEvent.result;
	$PSObj_ActivityDisplayName	= $AuditLogEvent.activityDisplayName;
	$PSObj_InitiatedByUsr_Id 	= $AuditLogEvent.initiatedBy.user.id;
	$PSObj_InitiatedByUsr_UPN 	= $AuditLogEvent.initiatedBy.user.userPrincipalName;
	$PSObj_InitiatedByUsr_IP	= $AuditLogEvent.initiatedBy.user.ipAddress;
	$PSObj_InitiatedByApp 		= $AuditLogEvent.initiatedBy.app.displayName;
	$PSObj_TargetGroupName 		= $AADGroupsDB[$PSObj_TargetGroupId]

	$UID = $AuditLogEvent.initiatedBy.user.userPrincipalName
	if ($UID) {
		if ($AADUsersDB.Contains($UID)) {
			$PSObj_InitiatedByUsr_Name = $AADUsersDB[$UID].displayName
		}
	}
	
	$UID = $PSObj_TargetUserId
	if ($UID) {
		if ($AADUsersDB.Contains($UID)) {
			$PSObj_TargetUserName = $AADUsersDB[$UID].displayName
			$PSObj_TargetUserMail = $AADUsersDB[$UID].mail
			$PSObj_TargetUserType = $AADUsersDB[$UID].userType
		}
	}

	$AuditLogEventReport += [pscustomobject]@{
		ActivityDateTime 		= $PSObj_ActivityDateTime;
		LoggedByService			= $PSObj_LoggedByService;
		Category				= $PSObj_Category;
		Result					= $PSObj_Result;
		ActivityDisplayName		= $PSObj_ActivityDisplayName;
		InitiatedByUsr_UPN 		= $PSObj_InitiatedByUsr_UPN;
		InitiatedByUsr_Name		= $PSObj_InitiatedByUsr_Name;
		InitiatedByUsr_Id 		= $PSObj_InitiatedByUsr_Id;
		InitiatedByUsr_IP		= $PSObj_InitiatedByUsr_IP;
		InitiatedByApp 			= $PSObj_InitiatedByApp;
		TargetUserUPN			= $PSObj_TargetUserUPN;
		TargetUserName			= $PSObj_TargetUserName;
		TargetUserMail			= $PSObj_TargetUserMail;
		TargetUserMailDomain	= $PSObj_TargetUserMail.Split("@")[1];
		TargetUserType			= $PSObj_TargetUserType;
		TargetUserId			= $PSObj_TargetUserId;
		TargetGroupName			= $PSObj_TargetGroupName;
		TargetGroupId			= $PSObj_TargetGroupId
	}
	$AuditLogEventsGrpCounter++
}

Write-Log "Audit log events for AAD groups parsed: $($AuditLogEventsGrpCounter)"
Export-Report "AAD groups audit report" -Report $AuditLogEventReport -Path $OutputFile -SortProperty "ActivityDateTime"

#######################################################################################################################

. $IncFile_StdLogEndBlock
