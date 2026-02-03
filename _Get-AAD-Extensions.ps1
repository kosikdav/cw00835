#######################################################################################################################
# Get-AAD-Users-Reports.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder					= "exports"
$LogFilePrefix				= "aad-users-extensions"


#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$ADCredentialPath = $aadauthmobmgmt_cred

[array]$UserListReport = @()
[array]$DeletedUserListReport = @()
[array]$CopilotLicenseReport = @()
[hashtable]$SIA_DB = @{}
[hashtable]$ADUser_DB = @{}


#######################################################################################################################

##############################################################################
# deleted users
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "directoryObjects/getAvailableExtensionProperties"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
$body = @{
	isSyncedFromOnPremises = $true
} | ConvertTo-Json

$Result = Invoke-RestMethod -Uri $Uri -Method "POST" -Body $body -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -ContentType "application/json"
$Extensions = $Result.value.Name

write-host "Found $($extensions.count) extension properties."
foreach ($ext in $extensions ) {
	if ($ext -match '(?<=^(?:[^_]*_){2}).*') {
    	$result = $Matches[0]
    	Write-Host $result
	}
}

$UriResource = "users/2d847bc0-ce2f-45c3-a6fd-86a98c5ea8dd"
$UriSelect = "id,displayName,userPrincipalName," + ($Extensions -join ",")
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
$user = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
write-host $user




#######################################################################################################################

