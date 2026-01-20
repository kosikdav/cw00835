#######################################################################################################################
# Add-AppPermission
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
	[Parameter(Mandatory)][string]$Id,
	[Parameter(Mandatory)][string]$Permission
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

[hashtable]$GraphPermissions_DB = @{}

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

$GraphPermissions_DB = Import-CSVtoHashDB -Path $DBFileAADPermissions -KeyName "value"

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

if ($GraphPermissions_DB.ContainsKey($Permission)) {
	$AppRoleId = $GraphPermissions_DB[$Permission].id
}
else {
	Write-Host "Permission '$($Permission)' not found in permissions database!"
	exit
}

Write-host ("Permission:").PadRight(15) -NoNewline
Write-Host $Permission -ForegroundColor Yellow -NoNewline
Write-Host " (AppRoleId: $($AppRoleId))"
Write-Host ("Application:").PadRight(15) -NoNewline 
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
	$Result = Invoke-RestMethod -Uri $Uri -Headers $Headers -ContentType $ContentTypeJSON -Body $Body -Method "POST"

	if ($null -ne $Result) {
		Write-Host "Permission '$($Permission)' added successfully!" -ForegroundColor Green
	}
	else {
		Write-Host "Failed to add permission '$($Permission)'!" -ForegroundColor Red
	}
}
else {
	Exit
}

