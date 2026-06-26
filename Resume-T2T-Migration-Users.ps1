#######################################################################################################################
# Resume-T2T-Migration-Users.ps1
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
$LogFilePrefix		= "t2t-migration-users-resume"
$LogFileFreq		= "Y"

#######################################################################################################################

. $ScriptPath\include-Script-Start-Include.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Freq $LogFileFreq -Ext "log"

[array]$Report = @()

$MigrationEndpoint = "EXO_CROSS_TENANT_ENDPOINT"
$HoldApplieErrString = "Cross tenant move is not supported when source mailbox has a hold"

$MigrationUsers = @()
$Report = @()

#######################################################################################################################

. $IncFile_StdLogStartBlock

$Dst_AppReg_EXO_MGMT = $AppReg_AIRPLUS_EXO_MGMT

Connect-EXOService -AppRegName $Dst_AppReg_EXO_MGMT  -TTL 120

$MigrationBatches = Get-MigrationBatch -Endpoint $MigrationEndpoint
Write-Log "Found $($MigrationBatches.Count) migration batches"

foreach ($Batch in $MigrationBatches) {
	write-host ($Batch.Identity.ToString()).PadRight(25) -ForegroundColor Green -NoNewline
	[array]$Result = Get-MigrationUser -BatchId $Batch.Identity
	write-host ": $($Result.Count)"
	$MigrationUsers += $Result
}

foreach ($User in $MigrationUsers) {
	if (($User.Status -eq "Failed") -and ($User.StatusSummary -eq "Failed")) {
		#write-host "User $($User.Identity) has failed migration with status summary: $($User.StatusSummary)" -ForegroundColor Yellow
		if ($User.ErrorSummary -like "$HoldApplieErrString*") {
			Start-MigrationUser -Identity $User.Identity -Confirm:$false
			Write-Log "Resuming migration for user $($User.Identity)" -ForegroundColor Green
		}
	}
}

#######################################################################################################################

. $IncFile_StdLogEndBlock
