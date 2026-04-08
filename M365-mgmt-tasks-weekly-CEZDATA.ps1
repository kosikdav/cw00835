#######################################################################################################################
# M365-download-reports
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$Stopwatch =  [system.diagnostics.stopwatch]::StartNew()
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdIncBlock.ps1

$ScriptList = @(
)

$LogFile = New-OutputFile -RootFolder $RLF -Prefix "_M365-mgmt-tasks-weekly" -Ext "log"

#######################################################################################################################

. $IncFile_StdLogStartBlock

ForEach ($Script in $ScriptList) {
    Write-Log "Starting $($script)"
    Start-Process -FilePath $psexe -ArgumentList "-File $($ScriptPath)\$($Script) -VariableDefinitionFile $($VariableDefinitionFile)" -Wait 
}

. $IncFile_StdLogEndBlock
