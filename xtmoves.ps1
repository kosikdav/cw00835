# Cycles through the previous 24 hours in 30-minute windows
# and calls Get-SPOCrossTenantUserContentMoveState for each window

$PartnerCrossTenantHostUrl = "https://cezdata-my.sharepoint.com"
$Limit = 1000

# Anchor everything to "now" in UTC
$EndTime   = (Get-Date).ToUniversalTime()
$StartTime = (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(-1).ToUniversalTime()

# UTC format expected by SPO cmdlets (ISO 8601 with Z suffix)
$DateFormat = "yyyy-MM-ddTHH:mm:ssZ"

$WindowMinutes = 10
$OverLapMinutes = 2
$Results = @()

$WindowStart = $StartTime

write-host "Start time: $($StartTime.ToString($DateFormat))"
write-host "End time:   $($EndTime.ToString($DateFormat))"
write-host "Window size: $WindowMinutes minutes"
write-host "Window overlap: $OverLapMinutes minutes"
write-host

while ($WindowStart -lt $EndTime) {

    $WindowEnd = $WindowStart.AddMinutes($WindowMinutes)
    if ($WindowEnd -gt $EndTime) { $WindowEnd = $EndTime }

    $MoveStartTime = $WindowStart.AddMinutes(-$OverLapMinutes).ToString($DateFormat)
    $MoveEndTime   = $WindowEnd.AddMinutes($OverLapMinutes).ToString($DateFormat)

    Write-Host "Querying window: $MoveStartTime  ->  $MoveEndTime     " -ForegroundColor Cyan -NoNewline

    try {
        $WindowResult = Get-SPOCrossTenantUserContentMoveState -PartnerCrossTenantHostUrl $PartnerCrossTenantHostUrl -Limit $Limit -MoveStartTime $MoveStartTime -MoveEndTime $MoveEndTime
        if ($WindowResult) {
            $Results += $WindowResult
            write-host $WindowResult.Count
        }
        else {
            Write-Host "0"
        }
    }
    catch {
        Write-Warning "Failed for window $MoveStartTime to $MoveEndTime : $_"
    }

    $WindowStart = $WindowEnd
}
write-host "Finished querying. Total records before deduplication: $($Results.Count)" -ForegroundColor Green
$Results = $Results | Sort-Object -Property MoveJobId -Unique
write-host "Total records after deduplication: $($Results.Count)" -ForegroundColor Green

Write-Host "`nTotal records collected: $($Results.Count)" -ForegroundColor Green

# Optional: export to CSV
# $Results | Export-Csv -Path ".\SPOCrossTenantMoveState.csv" -NoTypeInformation
