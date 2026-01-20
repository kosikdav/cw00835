#######################################################################################################################
# M365-download-audit
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$Stopwatch =  [system.diagnostics.stopwatch]::StartNew()
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdIncBlock.ps1

$ScriptList = @(
    #"Get-PowerAut-Audit.ps1",
    #"Get-PowerApps-Audit.ps1",
    "Get-Copilot-Audit.ps1",
    "Get-AAD-Guests-Audit.ps1",
    "Get-Admin-Roles-Audit.ps1",
    "Get-AAD-Groups-Audit.ps1",
    "Get-Teams-Audit.ps1",
    "Get-AAD-Apps-Audit.ps1"
    "Get-AAD-SignIns.ps1"
)

$LogFile = New-OutputFile -RootFolder $RLF -Prefix "_M365-download-audit" -Ext "log"

#######################################################################################################################

. $IncFile_StdLogBeginBlock

Write-Log "PS EXE $($ps5exe)"

ForEach ($Script in $ScriptList) {
    Write-Log "Starting $($script)"
    Start-Process -FilePath $ps5exe -ArgumentList "-File $($ScriptPath)\$($Script) -VariableDefinitionFile $($VariableDefinitionFile)" -Wait
}

. $IncFile_StdLogEndBlock
