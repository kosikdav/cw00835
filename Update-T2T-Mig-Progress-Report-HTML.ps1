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
#######################################################################################################################

. $IncFile_StdLogStartBlock

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
	write-host ("failed: {0,3}" -f $FailedUsers.Count) -ForegroundColor Red
	$TotalMigrationUsers += $Result.Count
	$TotalSyncedUsers += $SyncedUsers.Count
	$MigrationUsers += $Result
    $Report += [PSCustomObject]@{
        BatchName = $Batch.Identity.ToString()
        TotalUsers = $Result.Count
        SyncedUsers = $SyncedUsers.Count
        PendingUsers = $Result.Count - $SyncedUsers.Count
        FailedUsers = $FailedUsers.Count
        PercentageSynced = if ($Result.Count -gt 0) { "{0:N2}" -f (($SyncedUsers.Count / $Result.Count) * 100) } else { "0.00" }
    }
}

$TotalProgress = [math]::Round(($TotalSyncedUsers / $TotalMigrationUsers) * 100, 2)
$PendingUsers = $TotalMigrationUsers - $TotalSyncedUsers


$DateGenerated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
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
    <h3>Total Progress: $TotalProgress %</h3>
    <p>Pending Users: $PendingUsers</p>
    
    <table>
        <tr>
            <th>Batch name</th>
            <th>Total</th>
            <th>Synced</th>
            <th>Pending</th>
            <th>Failed</th>
            <th>Percentage synced</th>
        </tr>
"@

foreach ($item in $Report) {
    $rowClass = ""
    $fontColor = "black"
    $fontStyle = "normal"
    if ($HighlightedSKUs -contains $item.batchName) {
        $rowClass = "notice"
        $fontStyle = "bold"
    }
    $html += "        <tr class='$rowClass'>
    <td style='font-weight:$fontStyle;'>$($item.BatchName)</td>
    <td>$($item.TotalUsers)</td>
    <td>$($item.SyncedUsers)</td>
    <td>$($item.PendingUsers)</td>
    <td style='color:$fontColor; font-weight:$fontStyle;'>$($item.FailedUsers)</td>
    <td>$($item.PercentageSynced)</td></tr>`n"
}   

# Close the HTML
$html += @"
    </table>
    <p>This report provides an overview of the migration status.</p>
    <p>Report updated: $DateGenerated</p>
</body>
</html>
"@

# Output to file

$html | Out-File -FilePath $outFile -Encoding UTF8
    
. $IncFile_StdLogEndBlock
