
# Set-AAD-Synced-Users-Type.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "aad-guest-mgmt"
$LogFilePrefix		= "aad-ext-usr-type"
$daysBackOffset     = 30

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

#######################################################################################################################

. $IncFile_StdLogStartBlock

$guestsFixed = 0
$XTSyncCounter = 0

##############################################################################################
# read ext member users from Graph 
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriFilter = "userType eq 'Member' and onPremisesSyncEnabled eq false and externalUserState eq 'Accepted'"
$UriSelect = "id,userPrincipalName,userType,displayName,externalUserState,onPremisesSyncEnabled,onPremisesExtensionAttributes"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Filter $UriFilter -Select $UriSelect
[array]$AllExtMedmbers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ProgressDots -Text "AAD guest users"
write-host "Total external members: $($AllExtMedmbers.Count)"
#######################################################################################################################

. $IncFile_StdLogEndBlock
