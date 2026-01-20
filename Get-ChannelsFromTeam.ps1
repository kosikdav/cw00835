#######################################################################################################################
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
	$GroupId
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "exports"
$LogFilePrefix		= "teams-xxxreports"


#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

#######################################################################################################################


Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

$UriResource = "teams/$($GroupId)"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource 
$Team = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON

if ($Team) {
	write-host "Team $($Team.displayName)" -ForegroundColor Green
	$UriResource = "teams/$($GroupId)/channels"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource 
	[array]$Channels = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON

	foreach ($Channel in $Channels) {
		write-host "$($Channel.displayName) : $($Channel.id) "
	}
}	
else {
	Write-Host "Team with GroupId $GroupId not found"
	exit

}
