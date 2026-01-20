#######################################################################################################################
# Get-Admin-Roles-Reports
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder		= "exports"
$LogFilePrefix	= "aad-roles"

$OutputFolder   		= "admin-roles\reports"
$OutputFilePrefix		= "aad-roles"
$OutputFileSuffixDEF	= "definitions"
$OutputFileSuffixRMG 	= "RMG"
$OutputFileSuffixRAS 	= "RAS"
$OutputFileSuffixRES	= "RES"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputFileDEF = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixDEF -Ext "csv"
$OutputFileRMG = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixRMG -Ext "csv"
$OutputFileRAS = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixRAS -Ext "csv"
$OutputFileRES = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixRES -Ext "csv"

[array]$AdminRolesReportDEF = $null
[array]$AdminRolesReportPRM = $null
[array]$AdminRolesReportRMG = $null
[array]$AdminRolesReportRAS = $null
[array]$AdminRolesReportRES = $null





function Get-ScheduleReportObject {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)]$ScheduleObject,
		[Parameter(Mandatory=$true)]$PrincipalObject,
		[Parameter(Mandatory=$true)][string]$PrincipalType,
		[Parameter(Mandatory=$true)][string]$MemberType,
		$UserObject,
		[Parameter(Mandatory=$true)][string]$InheritedFromGroup
	)
	if ($PrincipalType -eq $TypeGroup) {
		$PrincipalUPN = $null
		$PrincipalName = $PrincipalObject.DisplayName
	}
	else {
		$PrincipalUPN = $UserObject.UserPrincipalName
		$PrincipalName = $UserObject.DisplayName
	}
	$ReportObject = [pscustomobject]@{
		ScheduleId						= $ScheduleObject.Id
		PrincipalId 					= $ScheduleObject.PrincipalId
		RoleDefinitionId				= $ScheduleObject.RoleDefinition.TemplateId
		RoleDisplayName					= $ScheduleObject.RoleDefinition.DisplayName
		PrincipalType					= $PrincipalType
		PrincipalUPN					= $PrincipalUPN
		PrincipalName					= $PrincipalName
		CompanyName						= $UserObject.companyName
		Department						= $UserObject.department
		JobTitle						= $UserObject.jobTitle
		CreatedUsing					= $ScheduleObject.createdUsing
		CreatedDateTime					= $ScheduleObject.createdDateTime
		ModifiedDateTime				= $ScheduleObject.modifiedDateTime
		Status							= $ScheduleObject.status
		MemberType						= $MemberType
		InhertitedFromGroup				= $InheritedFromGroup
		ScheduleInfoStartTime			= $ScheduleObject.scheduleInfo.startDateTime
		ScheduleInfoRecurrence			= $ScheduleObject.scheduleInfo.recurrence
		ScheduleInfoExpirationType		= $ScheduleObject.scheduleInfo.expiration.type
		ScheduleInfoExpirationEnd		= $ScheduleObject.scheduleInfo.expiration.end
		ScheduleInfoExpirationDuration	= $ScheduleObject.scheduleInfo.expiration.duration
	}
	return $ReportObject
}

#######################################################################################################################

. $IncFile_StdLogStartBlock

Write-Log "DEF output file: $($OutputFileDEF)"
Write-Log "RMG output file: $($OutputFileRMG)"
Write-Log "RAS output file: $($OutputFileRAS)"
Write-Log "RES output file: $($OutputFileRES)"

##################################################################################################
Write-Log "Role definitions" -ForegroundColor Yellow
##################################################################################################
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource  = "roleManagement/directory/roleDefinitions"
$UriSelect = "id,description,displayName,isBuiltIn,isEnabled,resourceScopes,rolePermissions,templateId,version"
$Uri = New-GraphUri -Version "beta" -Resource $UriResource -Select $UriSelect
[array]$RoleDefinitions = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
foreach ($RoleDefinition in $RoleDefinitions) {
	$ExcludedResourceActions = $Condition = [string]::Empty
	if ($RoleDefinition.rolePermissions.excludedResourceActions) {
		foreach ($Item in $RoleDefinition.rolePermissions.excludedResourceActions) {
			if ($Item.Length -ge 1) {
				$ExcludedResourceActions += $Item + ","
			}
		}
		$ExcludedResourceActions = $ExcludedResourceActions.TrimEnd(",")
	}
	if ($RoleDefinition.rolePermissions.Condition.Count) {
		foreach ($Item in $RoleDefinition.rolePermissions.Condition) {
			if ($Item.Length -ge 1) {
				$Condition += $Item + ","
			}
		}
		$Condition = $Condition.TrimEnd(",")
	}

	if ($RoleDefinition.rolePermissions.allowedResourceActions) {
		foreach ($AllowedResourceAction in $RoleDefinition.rolePermissions.allowedResourceActions) {
			$Namespace = $Entity = $PropertySet = $Action = [string]::Empty
			$Namespace = $AllowedResourceAction.Split("/")[0]
			$Entity = $AllowedResourceAction.Split("/")[1]
			if ($AllowedResourceAction -match "(?<=^[^/]+/[^/]+/).*(?=/[^/]*$)") {
				$PropertySet = $Matches[0]
			}
			if ($AllowedResourceAction -match "[^/]+$") {
				$Action = $Matches[0]
			}
			$AdminRolesReportDEF += [pscustomobject]@{
				Id			= $RoleDefinition.Id
				DisplayName	= $RoleDefinition.DisplayName
				IsBuilIn 	= $RoleDefinition.IsBuiltIn
				IsEnabled	= $RoleDefinition.IsEnabled
				isPrivileged = $RoleDefinition.isPrivileged
				TemplateId 	= $RoleDefinition.templateId
				Version 	= $RoleDefinition.version
				Namespace	= $Namespace
				Entity		= $Entity
				PropertySet	= $PropertySet
				Action		= $Action
				Permission	= $AllowedResourceAction
				ExcludedResourceActions = $ExcludedResourceActions
				Condition	= $Condition
			}
		}
	}
	else {
		$AdminRolesReportDEF += [pscustomobject]@{
			Id			= $RoleDefinition.Id
			DisplayName	= $RoleDefinition.DisplayName
			IsBuilIn 	= $RoleDefinition.IsBuiltIn
			IsEnabled	= $RoleDefinition.IsEnabled
			isPrivileged = $RoleDefinition.isPrivileged
			TemplateId 	= $RoleDefinition.templateId
			Version 	= $RoleDefinition.version
			Namespace	= "n/a"
			Entity		= "n/a"
			PropertySet	= "n/a"
			Action		= "n/a"
			Permission	= "none"
			ExcludedResourceActions = $ExcludedResourceActions
			Condition	= $Condition
		}
	}
}
Export-Report -Text "role definitions" -Report $AdminRolesReportDEF -Path $OutputFileDEF -SortProperty "DisplayName"

##################################################################################################
Write-Log "Role assignments (legacy)" -ForegroundColor Yellow
##################################################################################################
foreach ($RoleDefinition in $RoleDefinitions) {
	Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
	$UriResource = "roleManagement/directory/roleAssignments"
	$UriSelect = "id,principalId,principal"
	$UriFilter = "roleDefinitionId+eq+'{$($RoleDefinition.id)}'"
	$UriExpand = "principal"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect -Filter $UriFilter -Expand $UriExpand
	[array]$RoleAssignments = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON 
	foreach ($RoleAssignment in $RoleAssignments) {
		$AdminRolesReportRMG += [pscustomobject]@{
			RoleDefinitionId	= $RoleDefinition.Id;
			RoleDisplayName		= $RoleDefinition.DisplayName; 
			IsBuilIn 			= $RoleDefinition.IsBuiltIn; 
			IsEnabled			= $RoleDefinition.IsEnabled;
			PrincipalId			= $RoleAssignment.principalId;
			PrincipalUPN	 	= $RoleAssignment.principal.UserPrincipalName
			PrincipalName		= $RoleAssignment.principal.DisplayName
		}
	}
}

##################################################################################################
Write-Log "Role eligibility schedules" -ForegroundColor Yellow
##################################################################################################
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "roleManagement/directory/roleEligibilitySchedules"
$UriExpand = "RoleDefinition"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Expand $UriExpand
$RoleEligibilitySchedules = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
foreach ($RoleEligibilitySchedule in $RoleEligibilitySchedules) {
	$Principal = $User = $null
	$UriResource = "directoryObjects/$($RoleEligibilitySchedule.PrincipalId)"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
	try {
		$Principal = Invoke-RestMethod -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON
	}
	catch {
		$Principal = $null
	}
	if ($Principal."@odata.type" -eq $TypeUser) {
		$UriResource = "users/$($Principal.Id)"
		$UriSelect = "id,userPrincipalName,displayName,companyName,department,jobTitle"
		$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
		try {
			$User = Invoke-RestMethod -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON
		}
		catch {
			$User = $null
		}
	}
	else {
		if ($Principal."@odata.type" -eq $TypeGroup) {
			$Members = Get-GroupMembersFromGraphById -id $Principal.Id -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -Transitive:$true
			foreach ($Member in $Members) {
				$UriResource = "users/$($Member.Id)"
				$UriSelect = "id,userPrincipalName,displayName,companyName,department,jobTitle"
				$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
				try {
					$MemberUser = Invoke-RestMethod -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON
				}
				catch {
					$MemberUser = $null
				}
				$ScheduleReportObject = Get-ScheduleReportObject -ScheduleObject $RoleEligibilitySchedule -PrincipalObject $Principal -PrincipalType "User" -MemberType "GroupMember" -UserObject $MemberUser -InheritedFromGroup $Principal.DisplayName
				$AdminRolesReportRES += $ScheduleReportObject
			}
		}
	}
	$ScheduleReportObject = Get-ScheduleReportObject -ScheduleObject $RoleEligibilitySchedule -PrincipalObject $Principal -PrincipalType $Principal."@odata.type" -MemberType $RoleEligibilitySchedule.memberType -UserObject $User -InheritedFromGroup "n/a"
	$AdminRolesReportRES += $ScheduleReportObject
}


##################################################################################################
Write-Log "Role assignment schedules" -ForegroundColor Yellow
##################################################################################################
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "roleManagement/directory/roleAssignmentSchedules"
$UriExpand = "RoleDefinition"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Expand $UriExpand
$RoleAssignmentSchedules = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
foreach ($RoleAssignmentSchedule in $RoleAssignmentSchedules) {
	$Principal = $User = $null
	$UriResource = "directoryObjects/$($RoleAssignmentSchedule.PrincipalId)"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
	try {
		$Principal = Invoke-RestMethod -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON
	}
	catch {
		$Principal = $null
	}
	if ($Principal."@odata.type" -eq $TypeUser) {
		$UriResource = "users/$($Principal.Id)"
		$UriSelect = "id,userPrincipalName,displayName,companyName,department,jobTitle"
		$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
		try {
			$User = Invoke-RestMethod -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON
		}
		catch {
			$User = $null
		}
	}
	else {
		if ($Principal."@odata.type" -eq $TypeGroup) {
			$Members = Get-GroupMembersFromGraphById -id $Principal.Id -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -Transitive:$true
			foreach ($Member in $Members) {
				$UriResource = "users/$($Member.Id)"
				$UriSelect = "id,userPrincipalName,displayName,companyName,department,jobTitle"
				$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
				try {
					$MemberUser = Invoke-RestMethod -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON
				}
				catch {
					$MemberUser = $null
				}
				$ScheduleReportObject = Get-ScheduleReportObject -ScheduleObject $RoleAssignmentSchedule -PrincipalObject $Principal -PrincipalType "User" -MemberType "GroupMember" -UserObject $MemberUser -InheritedFromGroup $Principal.DisplayName
				$AdminRolesReportRAS += $ScheduleReportObject
			}
		}
	}
	$ScheduleReportObject = Get-ScheduleReportObject -ScheduleObject $RoleAssignmentSchedule -PrincipalObject $Principal -PrincipalType $Principal."@odata.type" -MemberType $RoleAssignmentSchedule.memberType -UserObject $User -InheritedFromGroup "n/a"
	$AdminRolesReportRAS += $ScheduleReportObject
}

Export-Report -Text "RMG (role management) report" -Report $AdminRolesReportRMG -Path $OutputFileRMG
Export-Report -Text "RAS (role assignment schedule) report" -Report $AdminRolesReportRAS -Path $OutputFileRAS -SortProperty "RoleDisplayName"
Export-Report -Text "RES (role eligibility schedule) report" -Report $AdminRolesReportRES -Path $OutputFileRES -SortProperty "RoleDisplayName"

#######################################################################################################################

. $IncFile_StdLogEndBlock

