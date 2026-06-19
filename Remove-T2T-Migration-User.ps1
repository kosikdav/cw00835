#######################################################################################################################
# Get-T2T-Migration-Progress.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
	$Identity
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1
. $ScriptPath\include\include-functions-T2T.ps1

#######################################################################################################################

$LogFolder			= "t2t"
$LogFilePrefix		= "t2t-migration-users-remove"
$LogFileFreq		= "Y"

#######################################################################################################################

. $ScriptPath\include-Script-Start-Include.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Freq $LogFileFreq -Ext "log"


#######################################################################################################################

$Dst_AppReg_EXO_MGMT = $AppReg_EXO_MGMT
Connect-EXOService -AppRegName $Dst_AppReg_EXO_MGMT  -TTL 30
Remove-MigrationUser -Identity $Identity -Confirm:$false
