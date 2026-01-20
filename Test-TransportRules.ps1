#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
	$Recipient,
	$MessageFile,
	$Sender,
	$sendReportTo
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "mbxmgmt"
$LogFilePrefix		= "mbxmgmt-full"
$LogFileFreq		= "Y"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Freq $LogFileFreq -Ext "log"

#######################################################################################################################
. $IncFile_StdLogStartBlock

Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 120
$data = [System.IO.File]::ReadAllBytes($MessageFile)
Test-Message -MessageFileData $data -Sender $sender -Recipients $Recipient -sendReportTo $sendReportTo -TransportRules

