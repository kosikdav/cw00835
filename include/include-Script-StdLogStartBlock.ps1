Write-Log "."
Write-Log "/-----------------------------------------------------------------------------------------------------------------------"
Write-Log "Script file: $($ScriptPath)\$($ScriptName)"
Write-Log "PowerShell version: $($PSVersion)"
If ([Environment]::UserInteractive) {
    Write-Log "Running interactively" -ForegroundColor DarkBlue -BackgroundColor Green
}
Else {
    Write-Log "Running non-interactively"
}
If ($VariableDefinitionFile) {
    Write-Log "VariableDefinitionFile: $($VariableDefinitionFile)"
}
If ($LogFile) {
    Write-Log "Log file: $($LogFile)"
}
If ($OutputFolder) {
    Write-Log "Output folder: $($OutputFolder)"
}
If ($OutputFile) {
    Write-Log "Output file: $($OutputFile)"
}
