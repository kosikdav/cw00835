#######################################################################################################################
# Copy-RoleDefinition
#######################################################################################################################
param(
	[Parameter(Mandatory=$true)][string]$SourceRoleId,
	[Parameter(Mandatory=$true)][string]$NewRoleDisplayName,
	[string]$NewRoleDescription
)
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

. $ScriptPath\include-Var-Define.ps1
. $ScriptPath\include-Var-Init.ps1
. $ScriptPath\include-Functions-Common.ps1

$NamespaceMSDir = "microsoft.directory"

Request-MSALToken -AppRegName $AppReg_ROLE_MGMT -TTL 30
$UriResource  = "roleManagement/directory/roleDefinitions/$($SourceRoleId)"
$UriSelect = "id,description,displayName,isBuiltIn,isEnabled,resourceScopes,rolePermissions,templateId,version"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
$SourceRoleDefinition = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_ROLE_MGMT].AccessToken -ContentType $ContentTypeJSON

if ($SourceRoleDefinition) {
	write-host "Cloning role definition $($SourceRoleDefinition.displayName)" -ForegroundColor Cyan
	[array]$NewRolePermissions = @()
	[array]$allowedResourceActions = @()
	[array]$condition = @()

	if ($SourceRoleDefinition.rolePermissions.allowedResourceActions) {
		foreach ($Item in $SourceRoleDefinition.rolePermissions.allowedResourceActions) {
			if ($Item.Split("/")[0] -eq $NamespaceMSDir) {
				$allowedResourceActions += $Item
			}
			else {
				write-host "Skipping $($Item) - not in $NamespaceMSDir namespace" -ForegroundColor DarkGray
			}
		}
		$NewRolePermissions = [pscustomobject]@{
			allowedResourceActions = $allowedResourceActions
		}
	} else {
		write-host "No allowed resource actions found" -ForegroundColor DarkGray
		Exit
	}
	
	if ($SourceRoleDefinition.rolePermissions.condition) {
		foreach ($Item in $SourceRoleDefinition.rolePermissions.allowedResourceActions) {
			$condition += $Item
		}
		$NewRolePermissions = [pscustomobject]@{
			condition = $condition
		}
	}
	
	$NewRoleDefinition = [pscustomobject]@{
		displayName = $NewRoleDisplayName
		description = $NewRoleDescription
		rolePermissions = $NewRolePermissions
		isEnabled = $true
	}

	$NewRoleDefinition = $NewRoleDefinition | ConvertTo-JSON -Depth 3
	write-host "New role definition JSON: $($NewRoleDefinition)" -ForegroundColor DarkGray

	$UriResource  = "roleManagement/directory/roleDefinitions"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
	Try {
		$Result = Invoke-RestMethod -Uri $Uri -Headers $AuthDB[$AppReg_ROLE_MGMT].AuthHeaders -ContentType $ContentTypeJSON -Method "POST" -Body $NewRoleDefinition
		write-host "Role definition $($NewRoleDisplayName) created successfully" -ForegroundColor Green
	} Catch {
		$ErrorMessage = $_.Exception.Message
		write-host "Error: $ErrorMessage"
	}
}
