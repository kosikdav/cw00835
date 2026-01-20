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
    "Get-Admin-Roles-Reports.ps1",
    "Get-SKU-Report.ps1",
    "Get-AAD-Guests-Reports.ps1",
    "Get-AAD-Users-Reports.ps1",
    "Get-M365-User-License-Assignments.ps1",
    "Get-AAD-Devices-Reports.ps1",
    "Get-Intune-Devices-Reports.ps1",
    "Get-AAD-Groups-Reports.ps1",
    "Get-AAD-Groups-Reports-OU.ps1",
    "Get-Teams-Reports.ps1",
    "Get-Teams-Reports-Chats.ps1",
    "Get-AAD-Apps-Reports.ps1",
    "Get-EXO-Mailboxes-Reports.ps1",
    "Get-EXO-Role-Assignments.ps1",
    "Get-SPO-Reports.ps1",
    "Get-Reports-Tenaur.ps1"
)

$LogFile = New-OutputFile -RootFolder $RLF -Prefix "_M365-download-reports" -Ext "log"

#######################################################################################################################

. $IncFile_StdLogBeginBlock

Write-Log "PS EXE $($ps5exe)"

ForEach ($Script in $ScriptList) {
    Write-Log "Starting $($script)"
    Start-Process -FilePath $ps5exe -ArgumentList "-File $($ScriptPath)\$($Script) -VariableDefinitionFile $($VariableDefinitionFile)" -Wait    
}

. $IncFile_StdLogEndBlock
