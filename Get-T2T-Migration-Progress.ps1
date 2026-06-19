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
	write-host ($Batch.Identity.ToString()).PadRight(25) -NoNewline
	[array]$Result = Get-MigrationUser -BatchId $Batch.Identity
	[array]$FailedUsers = $Result | Where-Object { $_.Status -eq "Failed" }
	[array]$SyncedUsers = $Result | Where-Object { $_.Status -eq "Synced" }
	write-host ("total: {0,3} " -f $Result.Count) -NoNewline
	write-host ("synced: {0,3} " -f $SyncedUsers.Count) -NoNewline -ForegroundColor Green
	write-host ("pending: {0,3} " -f ($Result.Count - $SyncedUsers.Count)) -NoNewline -ForegroundColor Yellow
	write-host ("failed: {0,3}" -f $FailedUsers.Count) -ForegroundColor Red
	$TotalMigrationUsers += $Result.Count
	$TotalSyncedUsers += $SyncedUsers.Count
	$MigrationUsers += $Result
}
write-host
write-host ("Migration users: {0,4}" -f $TotalMigrationUsers)
write-host ("Synced users:    {0,4}" -f $TotalSyncedUsers)
write-host ("Pending users:   {0,4}" -f ($TotalMigrationUsers - $TotalSyncedUsers)) -ForegroundColor cyan
Write-Host
write-host ("Progress:         {0,6:N2} %" -f ($([math]::Round(($TotalSyncedUsers / $TotalMigrationUsers) * 100, 2)))) -ForegroundColor Yellow


