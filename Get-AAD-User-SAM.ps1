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
$LogFilePrefix				= "aad-users"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"


#######################################################################################################################


Request-MSALToken -AppRegName $AppReg_LOG_READER_MIN -TTL 30
$UriResource = "users"
#$UriSelect = "id,UserPrincipalName,DisplayName"
$UriSelect = "id,UserPrincipalName,DisplayName,onPremisesSamAccountName"
$UriFilter = "startsWith(userPrincipalName,'david.kosik')"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect -Filter $UriFilter
[array]$Users = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER_MIN].AccessToken -ContentType $ContentTypeJSON

write-host $Users