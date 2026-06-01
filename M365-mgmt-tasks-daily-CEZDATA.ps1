#######################################################################################################################
# M365-mgmt-tasks-daily-CEZDATA
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$Stopwatch =  [system.diagnostics.stopwatch]::StartNew()
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Include.ps1

$ScriptList = @(
    "Set-AAD-Synced-Users-Type.ps1",
    "Update-Usage-Location.ps1",
    "Remove-Stale-AAD-Guests.ps1",
    "Set-AAD-MFA-Phone-From-IDM.ps1",
    "Set-AAD-Guests-Attributes.ps1",
    "Update-ADO-Guest-Group-Owners.ps1"
)

$LogFile = New-OutputFile -RootFolder $RLF -Prefix "_M365-mgmt-tasks-daily" -Ext "log"

#######################################################################################################################

. $IncFile_StdLogStartBlock

ForEach ($Script in $ScriptList) {
    Write-Log "Starting $($script)"
    Start-Process -FilePath $psexe -ArgumentList "-File $($ScriptPath)\$($Script) -VariableDefinitionFile $($VariableDefinitionFile)" -Wait 
}

. $IncFile_StdLogEndBlock
