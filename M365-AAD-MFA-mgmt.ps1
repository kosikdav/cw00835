$EnableOnScreenLogging = $true
$ManualAuth = $false
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

. $ScriptPath\include-Functions-Init.ps1

$MainStopwatch  =  [system.diagnostics.stopwatch]::StartNew()
$LogPath        = "d:\logs\"
$LogFilePrefix  = "O365-AAD-MFA-mgmt-"
$LogFileName    = GetLogFileName("Y")

. $ScriptPath\include-functions-common.ps1

LogWrite -LogString "--------------------------------------------------------------------------------"
LogWrite -LogString "Script start: $($ScriptName)"

LogWrite -LogString "Starting Set-AADMFAPhoneFromIDM"
Start-Process powershell {d:\scripts\Set-AADMFAPhoneFromIDM.ps1} -Wait

LogWrite -LogString "Starting Set-AADMFAMethods"
Start-Process powershell {d:\scripts\Set-AADMFAMethods.ps1} -Wait

LogWrite -LogString "Run time: $($MainStopwatch.Elapsed)"
LogWrite -LogString "Script end"
