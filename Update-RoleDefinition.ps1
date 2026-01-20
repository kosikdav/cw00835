#######################################################################################################################
# Update-RoleDefinition
#######################################################################################################################
param(
	[string]$RoleId
)
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

. $ScriptPath\include-Var-Define.ps1
. $ScriptPath\include-Var-Init.ps1
. $ScriptPath\include-Functions-Common.ps1


function Get-YesNoKeyboardInput {
    param (
        [Parameter(Mandatory=$true)][string]$Prompt
    )
    Write-Host "$($Prompt) [Y/N]" -ForegroundColor Yellow
    :prompt 
    while ($true) {
        switch ([console]::ReadKey($true).Key) {        
            { $_ -eq [System.ConsoleKey]::Y } { Return $true }        
            { $_ -eq [System.ConsoleKey]::N } { Return $false }        
            default { Write-Host "Only 'Y' or 'N' allowed!" }    
        }
    }
}
function Get-UserSelection {
	param(
		[string]$Text,
		[int]$LastIndex 
	)
	$PromptText = "Enter $($Text) (0-$($LastIndex)) or Q to quit"
	$Regex = "^[a-zA-Z0-9\- ]+$"
	Do {
		$UserInput = $null
		$Prompt = Read-Host $PromptText
		If ($Prompt -eq "Q") {
			Exit
		}
		if ($Prompt -match $Regex) {
			$UserInput = $Prompt
		} else {
			Write-Host $PromptText
		}
	} Until ($UserInput)
	return $UserInput
}

function Get-UserStringInput {
	param(
		[string]$Text
	)
	$PromptText = "Enter $($Text) or Q to quit"
	Do {
		$UserInput = $null
		$Prompt = Read-Host $PromptText
		If ($Prompt -eq "Q") {
			Exit
		}
		else {
			$UserInput = $Prompt
		}
	} Until ($UserInput)
	return $UserInput
}

Request-MSALToken -AppRegName $AppReg_ROLE_MGMT -TTL 30

if ($RoleId) {
	$UriResource  = "roleManagement/directory/roleDefinitions/$($RoleId)"
	$UriSelect = "id,description,displayName,isBuiltIn,isEnabled,resourceScopes,rolePermissions,templateId,version"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
	$TargetRoleDefinition = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_ROLE_MGMT].AccessToken -ContentType $ContentTypeJSON
	if ($TargetRoleDefinition) {
		if ($TargetRoleDefinition.isBuiltIn -eq $true) {
			write-host "Cannot edit built-in role definition $($TargetRoleDefinition.displayName)" -ForegroundColor Red
			Exit
		}
	}
	else {
		write-host "Role definition $($RoleId) not found" -ForegroundColor Red
		Exit
	}
}
else {
	$UriResource  = "roleManagement/directory/roleDefinitions"
	$UriSelect = "id,description,displayName,isBuiltIn,isEnabled,resourceScopes,rolePermissions,templateId,version"
	$UriFilter = "isBuiltIn+eq+false"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect -Filter $UriFilter
	[array]$TargetRoleDefinitions = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_ROLE_MGMT].AccessToken -ContentType $ContentTypeJSON
	foreach ($TargetRoleDefinition in $TargetRoleDefinitions) {
		write-host "[$($TargetRoleDefinitions.IndexOf($TargetRoleDefinition).ToString().PadRight(2))] $($TargetRoleDefinition.displayName)" -ForegroundColor Cyan
	}
	$RoleToEdit = Get-UserSelection -Text "role number" -LastIndex ($TargetRoleDefinitions.Count-1)
	$TargetRoleDefinition = $TargetRoleDefinitions[$RoleToEdit]
}

Clear-Host
write-host "Editing role $($TargetRoleDefinition.displayName) ($($TargetRoleDefinition.id))" -ForegroundColor Cyan
write-host "Select action:" -ForegroundColor Yellow
write-host "A - Add permissions" -ForegroundColor Yellow
write-host "R - Remove permissions" -ForegroundColor Yellow
write-host "Q - Quit" -ForegroundColor Yellow

Do {
	$Action = $null
	$KeyPressed = [console]::ReadKey($true).Key
	switch ($KeyPressed) {
		"A" { $Action = "Add" }
		"R" { $Action = "Remove" }
		"Q" { Exit }
	}
} Until ($Action)


if ($Action -eq "Remove") {
	Clear-Host
	write-host "Removing permissions from role $($TargetRoleDefinition.displayName) ($($TargetRoleDefinition.id))" -ForegroundColor Cyan
	[array]$NewRolePermissions = @()
	[array]$AllowedResourceActions = @()
	[array]$NewAllowedResourceActions = @()

	if ($TargetRoleDefinition.rolePermissions.allowedResourceActions) {
		foreach ($Item in $TargetRoleDefinition.rolePermissions.allowedResourceActions) {
			$AllowedResourceActions += $Item
		}
	}
	else {
		write-host "No allowedResourceActions found, role empty" -ForegroundColor DarkGray
		Exit
	}
	write-host "Existing allowedResourceActions:"
	foreach ($Item in $allowedResourceActions) {
		write-host "[$($allowedResourceActions.IndexOf($Item).ToString().PadRight(2))] " -NoNewline
		write-host $Item -ForegroundColor Cyan
	}
	$ActionToRemove = Get-UserSelection -Text "permission to remove" -LastIndex ($allowedResourceActions.Count-1)
	
	Write-Host "Removing $($allowedResourceActions[$ActionToRemove]) from allowedResourceActions" -ForegroundColor Yellow
	$NewAllowedResourceActions = $allowedResourceActions | Where-Object { $_ -ne $allowedResourceActions[$ActionToRemove] }
	write-host "Updated allowedResourceActions:"
	foreach ($Item in $AllowedResourceActions) {
		if ($AllowedResourceActions.IndexOf($Item) -eq $ActionToRemove){
			write-host "[$($AllowedResourceActions.IndexOf($Item).ToString().PadRight(2))] " -NoNewline
			write-host $Item -ForegroundColor DarkGray
		}
		else {
			write-host "[$($AllowedResourceActions.IndexOf($Item).ToString().PadRight(2))] " -NoNewline
			write-host $Item -ForegroundColor Cyan
		}
	}
	$NewRolePermissions = [pscustomobject]@{
		allowedResourceActions = $NewAllowedResourceActions
	}	
	$NewRoleDefinition = [pscustomobject]@{
		rolePermissions = $NewRolePermissions
	}

	$NewRoleDefinition = $NewRoleDefinition | ConvertTo-JSON -Depth 3
	#write-host "New role definition JSON: $($NewRoleDefinition)" -ForegroundColor DarkGray
	
	Exit

	$UriResource  = "roleManagement/directory/roleDefinitions/$($TargetRoleDefinition.id)"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
	Try {
		#$Result = Invoke-RestMethod -Uri $Uri -Headers $AuthDB[$AppReg_ROLE_MGMT].AuthHeaders -ContentType $ContentTypeJSON -Method "PATCH" -Body $NewRoleDefinition
		write-host "Role definition $($NewRoleDisplayName) updated successfully" -ForegroundColor Green
	} Catch {
		$ErrorMessage = $_.Exception.Message
		write-host "Error: $ErrorMessage"
	}
}

if ($Action -eq "Add") {
	[array]$ResourceActionsList = Import-CSVToArray -Path $DBFileAADResourceActions
	Clear-Host
	write-host "Adding permissions to role $($TargetRoleDefinition.displayName) ($($TargetRoleDefinition.id))" -ForegroundColor Cyan
	[array]$NewRolePermissions = @()
	[array]$allowedResourceActions = @()

	if ($TargetRoleDefinition.rolePermissions.allowedResourceActions) {
		foreach ($Item in $TargetRoleDefinition.rolePermissions.allowedResourceActions) {
			$allowedResourceActions += $Item
		}
	}
	else {
		write-host "No allowedResourceActions found, role empty" -ForegroundColor DarkGray
	}
	write-host "Existing allowedResourceActions:"
	foreach ($Item in $allowedResourceActions) {
		write-host "[$($allowedResourceActions.IndexOf($Item).ToString().PadRight(2))] " -NoNewline
		write-host $Item -ForegroundColor Cyan
	}
	Do {
		$VerifiedActionToAdd = $null
		$ActionToAdd = Get-UserStringInput -Text "permission to add"
		if ($ResourceActionsList.Name -contains $ActionToAdd.Trim()) {
			$VerifiedActionToAdd = $ActionToAdd.Trim()
		}
		else {
			write-host "Permission $($ActionToAdd) not found in resource actions list" -ForegroundColor Red
		}
	} Until ($VerifiedActionToAdd)

	Write-Host "Adding $($VerifiedActionToAdd) to allowedResourceActions" -ForegroundColor Yellow
	$NewAllowedResourceActions = $AllowedResourceActions += $VerifiedActionToAdd
	write-host "Updated allowedResourceActions:"
	foreach ($Item in $AllowedResourceActions) {
		write-host "[$($AllowedResourceActions.IndexOf($Item).ToString().PadRight(2))] " -NoNewline
		write-host $Item -ForegroundColor Cyan
	}
	write-host "[$($AllowedResourceActions.Count.ToString().PadRight(2))] " -NoNewline
	write-host $VerifiedActionToAdd -ForegroundColor Green

	$NewRolePermissions = [pscustomobject]@{
		allowedResourceActions = $NewAllowedResourceActions
	}	
	$NewRoleDefinition = [pscustomobject]@{
		rolePermissions = $NewRolePermissions
	}

	$NewRoleDefinition = $NewRoleDefinition | ConvertTo-JSON -Depth 3
	#write-host "New role definition JSON: $($NewRoleDefinition)" -ForegroundColor DarkGray
	
	Exit

	$UriResource  = "roleManagement/directory/roleDefinitions/$($TargetRoleDefinition.id)"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
	Try {
		$Result = Invoke-RestMethod -Uri $Uri -Headers $AuthDB[$AppReg_ROLE_MGMT].AuthHeaders -ContentType $ContentTypeJSON -Method "PATCH" -Body $NewRoleDefinition
		write-host "Role definition $($NewRoleDisplayName) updated successfully" -ForegroundColor Green
	} Catch {
		$ErrorMessage = $_.Exception.Message
		write-host "Error: $ErrorMessage"
	}
}

