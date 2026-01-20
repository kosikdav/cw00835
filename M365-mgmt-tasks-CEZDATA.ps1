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
    "Update-Usage-Location.ps1",
    "Get-M365-LicMgmt-Audit.ps1",
    "Update-AAD-Groups.ps1",    
    "Update-AD-Groups.ps1",
    "Sync-AAD-Groups.ps1",
    "Update-Teams-Channels-Membership.ps1",
    "Set-AAD-Guests-Attributes.ps1",
    "Update-AAD-Devices.ps1",
    "Remove-Stale-AAD-Guests.ps1"
)

$LogFile = New-OutputFile -RootFolder $RLF -Prefix "_M365-mgmt-tasks" -Ext "log"

#######################################################################################################################

. $IncFile_StdLogBeginBlock

ForEach ($Script in $ScriptList) {
    Write-Log "Starting $($script)"
    Start-Process -FilePath $psexe -ArgumentList "-File $($ScriptPath)\$($Script) -VariableDefinitionFile $($VariableDefinitionFile)" -Wait 
}

. $IncFile_StdLogEndBlock
