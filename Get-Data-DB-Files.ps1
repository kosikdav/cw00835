#######################################################################################################################
# Get-Data-DB-Files
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    [string]$workloads
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1
$ScriptNamePrexix = "Get-Data-DB-Files"

#######################################################################################################################

$LogFolder			= "db"
$LogFilePrefix		= "get-data-db-files"

if ($workloads) {
    $workloadArray = $workloads.ToUpper() -split ','
}
else {
    $workloadArray = @("TNT","APP","LIC","ROL","EXT","USR","TMS","GRP")
}

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

#######################################################################################################################
. $IncFile_StdLogStartBlock

Write-Log "Workloads: $($workloadArray -join ' ')"

ForEach ($Workload in $workloadArray) {
    $script = $ScriptNamePrexix + "-" + $Workload + ".ps1"
    Write-Log "Starting $($script)"
    Start-Process -FilePath $ps5exe -ArgumentList "-File $($ScriptPath)\$($Script) -VariableDefinitionFile $($VariableDefinitionFile)"
}

. $IncFile_StdLogEndBlock
