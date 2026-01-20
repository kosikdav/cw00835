#######################################################################################################################
# Get-Admin-Roles-Audit
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder 			= "exports"
$LogFilePrefix		= "admin-roles-audit"

$OutputFolder		= "admin-roles\audit"
$OutputFilePrefix	= "admin-roles-audit"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -FileDateYesterday -Ext "csv"

$start = $strYesterdayUTCStart
$end = $strYesterdayUTCEnd

[array]$RoleManagementAuditEventsReport = @()

#######################################################################################################################

. $IncFile_StdLogStartBlock

Write-Log "Output file: $($OutputFile)"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "auditLogs/directoryAudits"
$UriFilter = "category+eq+'RoleManagement'+and+activityDateTime+ge+$($start)+and+activityDateTime+le+$($end)&expand=children(=initiatedBy)"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Filter $UriFilter
$RoleManagementAuditEvents = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON

foreach ($RoleManagementAuditEvent in $RoleManagementAuditEvents) {
	foreach ($targetResource in $RoleManagementAuditEvent.targetResources) {
		if ($targetResource.type -eq "User") {
			$targetUserName = $targetResource.displayName
			$targetUserId = $targetResource.id
			$targetUserUPN = $targetResource.userPrincipalName
		}
		if ($targetResource.type -eq "Role") {
			$targetRoleId = $targetResource.id
			$targetRoleName = $targetResource.displayName
		}
		if ($targetResource.type -eq "Request") {
			$targetRequestId = $targetResource.id
		}
	}

	try {
		$RoleDefinitionOriginType = $($RoleManagementAuditEvent.additionalDetails|?{ $_.key -eq "RoleDefinitionOriginType"})[0].value;
	}
	catch {
		$RoleDefinitionOriginType = ""
	}
	
	try {
		$ipAddr = $($RoleManagementAuditEvent.additionalDetails|?{ $_.key -eq "ipaddr"})[0].value;
	}	
	catch {
		$ipAddr = ""
	}
	
	try {
		$resultReason = ($RoleManagementAuditEvent.resultReason).Trim("`t`n`r")
	}
	catch {
		$resultReason = ""	
	}

	if (($targetResource.type -eq "User") -and ($null -eq $targetUserName)) {
		Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
		$targetUser = Get-UserFromGraphById -Id $targetUserId -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
		$targetUserName = $targetUser.displayName
		$targetUserUPN = $targetUser.userPrincipalName
	}
	
	$RoleManagementAuditEventsReport += [pscustomobject]@{
		ActivityDateTime					= $RoleManagementAuditEvent.activityDateTime;
		LoggedByService						= $RoleManagementAuditEvent.loggedByService; 
		Category							= $RoleManagementAuditEvent.category; 
		CorrelationId						= $RoleManagementAuditEvent.correlationId; 
		Result								= $RoleManagementAuditEvent.result; 
		ResultReason						= $resultReason;
		OperationType						= $RoleManagementAuditEvent.operationType; 
		ActivityDisplayName					= $RoleManagementAuditEvent.activityDisplayName; 
		InitiatedBy							= $RoleManagementAuditEvent.initiatedBy;
		TargetRequestId						= $targetRequestId;
		TargetUserName						= $targetUserName;
		TargetUserUPN						= $targetUserUPN;
		TargetUserId						= $targetUserid;
		TargetRoleName						= $targetRoleName;
		TargetRoleId						= $targetRoleId;
		RoleDefinitionOriginType			= $RoleDefinitionOriginType;
		IpAddr								= $ipAddr
	}
}

Export-Report "Role management audit events" -Report $RoleManagementAuditEventsReport -Path $OutputFile

#######################################################################################################################

. $IncFile_StdLogEndBlock
