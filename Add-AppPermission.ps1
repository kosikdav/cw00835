#######################################################################################################################
# Add-AppPermission
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
	[Parameter(Mandatory)][string]$Id,
	[Parameter(Mandatory)][string]$Permission,
	[ValidateSet("Application","Delegated")]$Type
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

[hashtable]$GraphPermissions_App_DB = @{}
[hashtable]$GraphPermissions_Dlg_DB = @{}

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

##################################################################################################

[array]$GraphPermissions = Import-CSVtoArray -Path $DBFileAADPermissions

$GraphPermissions_App = $GraphPermissions | Where-Object { $_.type -eq "application" }
$GraphPermissions_App | forEach-Object { $GraphPermissions_App_DB.add($_.value, $_) }

$GraphPermissions_Dlg = $GraphPermissions | Where-Object { $_.type -eq "delegated" }
$GraphPermissions_Dlg | forEach-Object { $GraphPermissions_Dlg_DB.add($_.value, $_) }

Request-MSALToken -AppRegName $AppReg_APP_MGMT -TTL 30

$Application = $AppRoleId = $null

$UriResource = "servicePrincipals(appId='$($Id)')"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
$Application = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_APP_MGMT].AccessToken -ContentType $ContentTypeJSON -Silent

if ($null -eq $Application) {
	$UriResource = "servicePrincipals/$Id"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
	$Application = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_APP_MGMT].AccessToken -ContentType $ContentTypeJSON -Silent
}

if ($null -eq $Application) {
	Write-Host "Application with Id or AppId '$($Id)' not found!"
	exit
}

switch ($Type) {
	"Application" {
		if ($GraphPermissions_App_DB.ContainsKey($Permission)) {
			$AppRoleId = $GraphPermissions_App_DB[$Permission].id
		}
		else {
			Write-Host "Application permission '$($Permission)' not found in database!"
			exit
		}
	}
	"Delegated" {
		if ($GraphPermissions_Dlg_DB.ContainsKey($Permission)) {
			$AppRoleId = $GraphPermissions_Dlg_DB[$Permission].id
		}
		else {
			Write-Host "Delegated permission '$($Permission)' not found in database!"
			exit
		}
	}
}

Write-host ("Permission:").PadRight(15) -NoNewline
Write-Host "$($Permission) ($($Type))" -ForegroundColor Yellow -NoNewline
Write-Host " (AppRoleId: $($AppRoleId))"
Write-Host "Application: " -NoNewline 
Write-Host $Application.displayName -ForegroundColor Cyan -NoNewline
Write-Host " (AppId: $($Application.appId))"
If (Get-YesNoKeyboardInput -Prompt "Continue?") {
	$Headers = $AuthDB[$AppReg_APP_MGMT].AuthHeaders 
	$UriResource = "servicePrincipals/$($Application.id)/appRoleAssignments"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
	$Body = @{
		"principalId" = $Application.id
		"resourceId"  = $MSGraphResourceId
		"appRoleId"   = $AppRoleId
	} | ConvertTo-Json
	write-host $Body 
	Try {
		$Result = Invoke-RestMethod -Uri $Uri -Headers $Headers -ContentType $ContentTypeJSON -Body $Body -Method "POST"

		if ($null -ne $Result) {
			Write-Host "Permission '$($Permission)' added successfully!" -ForegroundColor Green
		}
		else {
			Write-Host "Failed to add permission '$($Permission)'!" -ForegroundColor Red
		}
	}
	Catch {
		Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
		exit
	}

}
else {
	Exit
}

