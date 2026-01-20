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

$LogFolder 			= "exports"
$LogFilePrefix		= "aad-apps-audit"

$OutputFolder 		= "aad-apps\audit"
$OutputFilePrefix	= "aad-apps-audit"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$OutputFile = New-OutputFile -RootFolder $REF -Folder $OutputFolder -Prefix $OutputFilePrefix -FileDateYesterday -Ext "csv"

[array]$AppManagementAuditEventsReport = $null
$AppManagementAuditEvents = $null
$start = $strYesterdayUTCStart
$end = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"

##################################################################################################

. $IncFile_StdLogBeginBlock

$AADApp_DB = Import-CSVtoHashDB -Path $DBFileAADAppList -KeyName "id"
$AADSP_DB = Import-CSVtoHashDB -Path $DBFileAADSPList -KeyName "id"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "auditLogs/directoryAudits"
$UriFilter = "category+eq+'ApplicationManagement'+and+activityDateTime+ge+$($start)+and+activityDateTime+le+$($end)"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Filter $UriFilter
$AppManagementAuditEvents = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
Write-Log "Application Management audit events retrieved: $($AppManagementAuditEvents.Count)"

foreach ($AppManagementAuditEvent in $AppManagementAuditEvents) {
	$initiatedById = $initiatedByName = $ipAddr = $null

	write-host "============================================================================================" -ForegroundColor White
	write-host "Category:$($AppManagementAuditEvent.category) OpType:$($AppManagementAuditEvent.operationType) Activity: $($AppManagementAuditEvent.activityDisplayName) ($($AppManagementAuditEvent.result))" -ForegroundColor Green -NoNewline
	write-host "--targetResources: $($AppManagementAuditEvent.targetResources.Count)"
	if ($AppManagementAuditEvent.initiatedBy.User.id) {
		$initiatedById = $AppManagementAuditEvent.initiatedBy.User.id
		$initiatedByName = $AppManagementAuditEvent.initiatedBy.User.userPrincipalName
		$ipAddr = $AppManagementAuditEvent.initiatedBy.User.ipAddress
	}
	if ($AppManagementAuditEvent.initiatedBy.App.servicePrincipalId) {
		$initiatedById = $AppManagementAuditEvent.initiatedBy.App.servicePrincipalId
		$initiatedByName = $AppManagementAuditEvent.initiatedBy.App.displayName
	}
	
	write-host $AppManagementAuditEvent.initiatedBy.app -ForegroundColor Magenta
	if ($AppManagementAuditEvent.activityDisplayName -eq "PATCH UserAuthMethod.PatchSignInPreferencesAsync") {
		continue
	}
	foreach ($targetResource in $AppManagementAuditEvent.targetResources) {
		write-host "--------" -ForegroundColor White
		write-host "id:$($targetResource.id) name:$($targetResource.displayName) type:$($targetResource.type)" -ForegroundColor Yellow -NoNewline
		write-host "--modifiedProperties: $($targetResource.modifiedProperties.Count)"
		
		$NoUpdatedProperties = $OnlyTargetIdUpdated = $false
		$modifiedProperties = [string]::Empty
		$targetApp = $null
		if ($targetResource.type -eq "Application") {
			write-host "*Application" -ForegroundColor DarkRed
			Write-Host $targetResource
			if ($AADApp_DB.ContainsKey($targetResource.id)) {
				$targetApp = $AADApps_DB[$targetResource.id]
				Write-Host $targetApp.displayName -ForegroundColor DarkBlue
			}
		}
		elseif ($targetResource.type -eq "ServicePrincipal") {
			write-host "*ServicePrincipal" -ForegroundColor DarkGreen
			Write-Host $targetResource
			if ($AADSP_DB.ContainsKey($targetResource.id)) {
				$targetApp = $AADSP_DB[$targetResource.id]
				Write-Host $targetApp.displayName -ForegroundColor DarkGreen
			}
		}
		
		foreach ($modifiedProperty in $targetResource.modifiedProperties) {
			write-host $modifiedProperty -ForegroundColor Cyan
			if (($modifiedProperty.displayName -eq "Included Updated Properties") -and ($null -eq $modifiedProperty.oldValue) -and ($null -eq $modifiedProperty.newValue)) {
				$NoUpdatedProperties = $true
			}
			if (($modifiedProperty.displayName -eq "TargetId.ServicePrincipalNames") -and ($null -eq $modifiedProperty.oldValue) -and ($modifiedProperty.newValue -eq $targetApp.appId)) {
				$OnlyTargetIdUpdated = $true
			}
	
			$modifiedProperties = $modifiedProperties + $modifiedProperty.displayName + ":" + $modifiedProperty.oldValue + "=>" + $modifiedProperty.newValue + ";"
		}
		$modifiedProperties = $modifiedProperties.TrimEnd(";")

		if ($targetResource.type -eq "User") {
			$targetUserName = $targetResource.displayName
			$targetUserUPN = $targetResource.userPrincipalName
			$targetUserId = $targetResource.id
		}
		if ($targetResource.type -eq "ServicePrincipal") {
			foreach ($targetResource in $AppManagementAuditEvent.targetResources) {
				$targetSPName = $targetResource.displayName
				$targetSPid = $targetResource.id
			}
		}
		if ($targetResource.type -eq "Request") {
			$targetRequestId = $targetResource.id
		}

		try {
			$resultReason = ($AppManagementAuditEvent.resultReason).Trim("`t`n`r")
		}
		catch {
			$resultReason = $null 
		}
	
		if (($targetResource.type -eq "User") -and ($null -eq $targetUserName)) {
			Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
			$targetUser = Get-UserFromGraphById -Id $targetUserId -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
			$targetUserName = $targetUser.displayName
			$targetUserUPN = $targetUser.userPrincipalName	
		}

		$AppManagementAuditEventsReport += [pscustomobject]@{
			ActivityDateTime					= $AppManagementAuditEvent.activityDateTime;
			LoggedByService						= $AppManagementAuditEvent.loggedByService; 
			Category							= $AppManagementAuditEvent.category; 
			CorrelationId						= $AppManagementAuditEvent.correlationId; 
			Result								= $AppManagementAuditEvent.result; 
			ResultReason						= $resultReason;
			OperationType						= $AppManagementAuditEvent.operationType; 
			ActivityDisplayName					= $AppManagementAuditEvent.activityDisplayName; 
			InitiatedById						= $initiatedById;
			InitiatedByName						= $initiatedByName;
			IpAddr								= $ipAddr;
			TargetAppName						= $targetApp.displayName;
			TargetAppId							= $targetApp.AppId;
			TargetRequestId						= $targetRequestId;
			TargetUserName						= $targetUserName;
			TargetUserUPN						= $targetUserUPN;
			TargetUserId						= $targetUserId;
			ModifiedProperties					= $modifiedProperties
		}
	} #targetResources
}

#######################################################################################################################

Export-Report -Report $AppManagementAuditEventsReport -Path $OutputFile -SortProperty "ActivityDateTime"

. $IncFile_StdLogEndBlock
