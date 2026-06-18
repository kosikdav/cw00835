#######################################################################################################################
# Get-T2T-Migration-Progress.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1
. $ScriptPath\include\include-functions-T2T.ps1

#######################################################################################################################

$LogFolder			= "t2t"
$LogFilePrefix		= "t2t-migration-users"
$LogFileFreq		= "Y"

$OutputFolder = "t2t-ujvrez\migration-users"
$OutputFilePrefix = "t2t-migration-users"

#######################################################################################################################

. $ScriptPath\include-Script-Start-Include.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Freq $LogFileFreq -Ext "log"
$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Freq "YMDHM" -Ext "csv"

[array]$Report = @()

$MigrationEndpoint = "UJVREZ_T2T_EXO_MIGRATION_ENDPOINT"
$BatchPrefix = "UJV"

$MigrationUsers = @()
$TotalMigrationUsers = 0
$TotalSyncedUsers = 0

#######################################################################################################################

$Dst_AppReg_EXO_MGMT = $AppReg_EXO_MGMT

Connect-EXOService -AppRegName $Dst_AppReg_EXO_MGMT  -TTL 120

$MigrationBatches = Get-MigrationBatch -Endpoint $MigrationEndpoint
Write-host "Found $($MigrationBatches.Count) migration batches"

foreach ($Batch in $MigrationBatches) {
	write-host ($Batch.Identity.ToString()).PadRight(25) -ForegroundColor Green -NoNewline
	[array]$Result = Get-MigrationUser -BatchId $Batch.Identity
	write-host ": $($Result.Count)"
	$TotalMigrationUsers += $Result.Count
	[array]$SyncedUsers = $Result | Where-Object { $_.Status -eq "Synced" }
	$TotalSyncedUsers += $SyncedUsers.Count
	$MigrationUsers += $Result
}

write-host "Total migration users: $TotalMigrationUsers"
write-host "Total synced users: $TotalSyncedUsers"
write-host "Progress: $([math]::Round(($TotalSyncedUsers / $TotalMigrationUsers) * 100, 2)) %"

#######################################################################################################################

. $IncFile_StdLogEndBlock
