#######################################################################################################################
# Sync-AAD-Groups
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder          = "aad-group-mgmt"
$LogFilePrefix      = "aad-group-mgmt"
$LogFileSuffix      = "mirror"
$LogFileFreq        = "YMD"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Suffix $LogFileSuffix -Ext "log"

#######################################################################################################################

. $IncFile_StdLogStartBlock

Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
$GroupsToProcess = Import-CSVtoArray -Path $ConfigFile_AADGroupMirror

foreach ($Group in $GroupsToProcess) {
    Sync-GraphGroups -SourceGroup $Group.source_id -TargetGroup $Group.target_id -Mirror:$true -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken
}

#######################################################################################################################

. $IncFile_StdLogEndBlock
