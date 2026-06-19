#######################################################################################################################
# Update-Lic-Report-HTML
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    [string]$workloads
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
write-host "Script path: $ScriptPath"
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "db"
$LogFilePrefix		= "update-t2t-report-html"

#######################################################################################################################
write-host 3
. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

[array]$LicReport = @()
[hashtable]$MSLicData_DB = @{}
[int]$TTL = 30

$outFile = "c:\inetpub\sites\t2texo" + "\default.htm"

$MigrationEndpoint = "UJVREZ_T2T_EXO_MIGRATION_ENDPOINT"
$BatchPrefix = "UJV"
$Report = @()
$FinalizedUsersHTML = [string]::Empty
[array]$FinalizedUsersList = @()
$TotalMigrationUsers = $TotalSyncedUsers = $TotalPendingUsers = $TotalFinalizedUsers = $TotalFailedUsers = 0

#######################################################################################################################

. $IncFile_StdLogStartBlock

$Dst_AppReg_EXO_MGMT = $AppReg_EXO_MGMT

Connect-EXOService -AppRegName $Dst_AppReg_EXO_MGMT  -TTL 120

$MigrationBatches = Get-MigrationBatch -Endpoint $MigrationEndpoint
Write-host "Found $($MigrationBatches.Count) migration batches"

foreach ($Batch in $MigrationBatches) {
	write-host ($Batch.Identity.ToString()).PadRight(25) -NoNewline
	
	[array]$Result = Get-MigrationUser -BatchId $Batch.Identity
	
	[array]$SyncedUsers = $Result | Where-Object { $_.Status -eq "Synced" }
	[array]$SyncingUsers = $Result | Where-Object { $_.Status -eq "Syncing" }
	[array]$FinalizedUsers = $Result | Where-Object { $_.Status -eq "Completed" }
	[array]$FailedUsers = $Result | Where-Object { $_.Status -eq "Failed" }
	
	write-host ("total: {0,3} " -f $Result.Count) -NoNewline
	write-host ("synced: {0,3} " -f $SyncedUsers.Count) -NoNewline -ForegroundColor Yellow
	write-host ("syncing: {0,3} " -f $SyncingUsers.Count) -NoNewline -ForegroundColor Yellow
	write-host ("finalized: {0,3} " -f $FinalizedUsers.Count) -NoNewline -ForegroundColor Green
	write-host ("failed: {0,3}" -f $FailedUsers.Count) -ForegroundColor Red
	$TotalMigrationUsers += $Result.Count
	$TotalSyncedUsers += $SyncedUsers.Count
	$TotalPendingUsers += $PendingUsers.Count
	$TotalFinalizedUsers += $FinalizedUsers.Count
	$TotalFailedUsers += $FailedUsers.Count
    $Report += [PSCustomObject]@{
        BatchName = $Batch.Identity.ToString()
        TotalUsers = $Result.Count
        SyncedUsers = $SyncedUsers.Count
        SyncingUsers = $SyncingUsers.Count
        FinalizedUsers = $FinalizedUsers.Count
        FailedUsers = $FailedUsers.Count
    }
	$MigrationUsers += $Result
    foreach ($User in $FinalizedUsers) {
        $FinalizedUsersList += $User.Identity
    }
}

$TotalProgress = [math]::Round(($TotalFinalizedUsers / $TotalMigrationUsers) * 100, 2)

$DateGenerated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
foreach ($User in $FinalizedUsersList) {
    $FinalizedUsersHTML += "<p>$User</p>`n"
}

# Convert to HTML table with some styling
$html = @"
<html>
<head>
    <title>T2T EXO migration status</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        table { border-collapse: collapse; width: 60%; }
        th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
        t    /* Highlight styles */
        .low { background-color: #ffdddd; }    /* light red */
        .high { background-color: #ddffdd; }   /* light green */
        .notice { background-color: #fff7cc; } /* light yellow */
    </style>
</head>

<body>
    <h1>T2T EXO Migration Status Report</h1>
    <h3>Total progress: $TotalProgress %</h3>
    <p>Finalized users: $TotalFinalizedUsers</p>
    
    <table>
        <tr>
            <th>Batch name</th>
            <th>Total</th>
            <th>Synced</th>
            <th>Syncing</th>
            <th>Finalized</th>
            <th>% finalized</th>
            <th>Failed</th>
        </tr>
"@

foreach ($item in $Report) {
    $FinalizedPercent = [math]::Round(($item.FinalizedUsers / $item.TotalUsers) * 100, 0)
    $FinalizedPercentString = ("{0,2}" -f $FinalizedPercent)
    $rowClass = ""
    $fontColorFinalized = "black"
    $fontColorFailed = "black"
    $fontStyleFinalized = "normal"
    $fontStyleFailed = "normal"
    if ($item.FinalizedUsers -eq $item.TotalUsers) {
        $fontStyleFinalized = "bold"
        $fontColorFinalized = "green"
    }
    if ($item.FailedUsers -gt 0) {
        $fontStyleFailed = "bold"
        $fontColorFailed = "red"
    }
    $html += "        <tr class='$rowClass'>
    <td style='font-weight:$fontStyle;'>$($item.BatchName)</td>
    <td style='text-align:right'>$($item.TotalUsers)</td>
    <td style='text-align:right'>$($item.SyncedUsers)</td>
    <td style='text-align:right'>$($item.SyncingUsers)</td>
    <td style='text-align:right; color:$fontColorFinalized; font-weight:$fontStyleFinalized;'>$($item.FinalizedUsers)</td>
    <td style='text-align:right; color:$fontColorFinalized; font-weight:$fontStyleFinalized;'>$FinalizedPercentString</td>
    <td style='text-align:right; color:$fontColorFailed; font-weight:$fontStyleFailed;'>$($item.FailedUsers)</td>
    </tr>`n"
}

# Close the HTML
$html += @"
    </table>
    $FinalizedUsersHTML
    <p>
    <p>Report updated: $DateGenerated</p>
</body>
</html>
"@

# Output to file

$html | Out-File -FilePath $outFile -Encoding UTF8
    
. $IncFile_StdLogEndBlock
