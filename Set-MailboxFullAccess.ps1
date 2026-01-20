#######################################################################################################################
# Set-MaiboxProperties
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
	$Identity,
	$User,
	$AutoMapping = $false
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "mbxmgmt"
$LogFilePrefix		= "mbx-full-access"
$LogFileFreq		= "Y"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Freq $LogFileFreq -Ext "log"


Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 120

Try {
	Add-MailboxPermission -Identity $Identity -User $User -AccessRights "FullAccess" -InheritanceType "All" -AutoMapping $AutoMapping
	Write-Log "INFO: Successfully set Full Access permission on mailbox '$Identity' for user '$User'."
}
Catch {
	Write-Log "ERROR: Failed to set Full Access permission on mailbox '$Identity' for user '$User'. $_" -MessageType "ERR"
}

