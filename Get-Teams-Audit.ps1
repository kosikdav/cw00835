#######################################################################################################################
# Get-Teams-Audit
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder 				= "exports"
$LogFilePrefix			= "teams-audit"

$OutputFolder		    = "teams\audit"
$OutputFilePrefix		= "teams-audit"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -FileDateYesterday -Ext "csv"

[array]$AuditLogEventReport = @()
$AuditLogEventsGrpCounter = 0

$start = $strYesterdayUTCStart
$end = $strYesterdayUTCEnd

#######################################################################################################################

. $IncFile_StdLogStartBlock

$AADUsers_DB = Import-CSVtoHashDB -Path $DBFileUsersAllMin -KeyName "id"
$AADGroups_DB = Import-CSVtoHashDB -Path $DBFileGroupsAllMin -KeyName "id"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "auditLogs/directoryAudits"
$UriFilter = "loggedByService+eq+'Core Directory'+and+category+eq+'GroupManagement'+and+activityDateTime+ge+$($start)+and+activityDateTime+le+$($end)"
$Uri = New-GraphUri -Version "beta" -Resource $UriResource -Filter $UriFilter
$AuditLogEventsGrp = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON

foreach ($AuditLogEvent in $AuditLogEventsGrp) {
	$PSObj_InitiatedByUsr = $null
	$PSObj_InitiatedByUsr_Id = $null
	$PSObj_TargetUser_Id = $null
	$PSObj_TargetGroup_Id = $null

	if ($AuditLogEvent.activityDisplayName -eq "Update group") {
		Continue 
	}

	foreach ($targetResource in $AuditLogEvent.targetResources) {
		if ($targetResource.type -eq "User") {
			$PSObj_TargetUser_Id 	= $targetResource.id
		}
		if ($targetResource.type -eq "Group") {
			$PSObj_TargetGroup_Id 	= $targetResource.id
		}
	}

	if (-not($AADGroups_DB.ContainsKey($PSObj_TargetGroup_Id))) {
		Continue 
	} 

	$PSObj_ActivityDateTime 	= $AuditLogEvent.activityDateTime;
	$PSObj_LoggedByService		= $AuditLogEvent.loggedByService;
	$PSObj_Category				= $AuditLogEvent.category;
	$PSObj_Result				= $AuditLogEvent.result;
	$PSObj_ActivityDisplayName	= $AuditLogEvent.activityDisplayName;
	$PSObj_InitiatedByUsr_Id 	= $AuditLogEvent.initiatedBy.user.id;
	$PSObj_InitiatedByUsr_IP	= $AuditLogEvent.initiatedBy.user.ipAddress;
	$PSObj_InitiatedByApp 		= $AuditLogEvent.initiatedBy.app.displayName;

	if ($PSObj_InitiatedByUsr_Id) {
		if ($AADUsers_DB.Contains($PSObj_InitiatedByUsr_Id)) {
			$PSObj_InitiatedByUsr = $AADUsers_DB[$PSObj_InitiatedByUsr_Id]
		}
	}
	
	if ($PSObj_TargetUser_Id) {
		if ($AADUsers_DB.Contains($PSObj_TargetUser_Id)) {
			$PSObj_TargetUser = $AADUsers_DB[$PSObj_TargetUser_Id]
		}
	}

	if ($PSObj_TargetGroup_Id) {
		if ($AADGroups_DB.Contains($PSObj_TargetGroup_Id)) {
			$PSObj_TargetGroup = $AADGroups_DB[$PSObj_TargetGroup_Id]
		}
	}

	$AuditLogEventReport += [pscustomobject]@{
		ActivityDateTime 		= $PSObj_ActivityDateTime;
		LoggedByService			= $PSObj_LoggedByService;
		Category				= $PSObj_Category;
		Result					= $PSObj_Result;
		ActivityDisplayName		= $PSObj_ActivityDisplayName;

		InitiatedByUsr_Id 		= $PSObj_InitiatedByUsr.userId;
		InitiatedByUsr_UPN 		= $PSObj_InitiatedByUsr.userPrincipalName;
		InitiatedByUsr_DispName	= $PSObj_InitiatedByUsr.displayName;
		InitiatedByUsr_Type		= $PSObj_InitiatedByUsr.userType;

		InitiatedByUsr_IP		= $PSObj_InitiatedByUsr_IP;
		InitiatedByApp 			= $PSObj_InitiatedByApp;

		TargetUser_Id			= $PSObj_TargetUser.userId;
		TargetUser_UPN			= $PSObj_TargetUser.userPrincipalName;
		TargetUser_DispName		= $PSObj_TargetUser.displayName;
		TargetUser_Mail			= $PSObj_TargetUser.mail;
		TargetUser_Type			= $PSObj_TargetUser.userType;
		
		TargetGroup_Id			= $PSObj_TargetGroup.Id;
		TargetGroup_DispName	= $PSObj_TargetGroup.displayName;
	}
	$AuditLogEventsGrpCounter++
}

Write-Log "Total audit log events for Teams parsed: $($AuditLogEventsGrpCounter)"
Export-Report -Text "audit log events for Teams" -Report $AuditLogEventReport -SortProperty "ActivityDateTime" -Path $OutputFile

#######################################################################################################################

. $IncFile_StdLogEndBlock
