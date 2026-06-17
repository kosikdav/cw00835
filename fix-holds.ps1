write-host "SRC mailboxes total: " -NoNewline
$outputFolder = "D:\exports"
[array]$EXOMailboxes = Get-Mailbox -ResultSize Unlimited
write-host $EXOMailboxes.count
$report = @()
$stopwatch = New-Object System.Diagnostics.Stopwatch
$stopwatch.Start()

write-host "Removing holds for mailboxes..."
foreach ($Mailbox in $EXOMailboxes) {
    #write-host "Processing mailbox: $($Mailbox.primarySMTPAddress)" -ForegroundColor Green
    if ($Mailbox.LitigationHoldEnabled) {
        write-host " Removing Litigation Hold for mailbox: $($Mailbox.primarySMTPAddress)"
        Set-Mailbox -Identity $Mailbox.primarySMTPAddress -LitigationHoldEnabled $false -Confirm:$false
        $report += $Mailbox.primarySMTPAddress
    }
    if ($Mailbox.DelayHoldApplied) {
        write-host " Removing Delay Hold for mailbox: $($Mailbox.primarySMTPAddress)"
        Set-Mailbox -Identity $Mailbox.primarySMTPAddress -RemoveDelayHoldApplied
        $report += $Mailbox.primarySMTPAddress
    }
    if ($Mailbox.DelayReleaseHoldApplied) {
        write-host " Removing Delay Release Hold for mailbox: $($Mailbox.primarySMTPAddress)"
        Set-Mailbox -Identity $Mailbox.primarySMTPAddress -RemoveDelayReleaseHoldApplied
        $report += $Mailbox.primarySMTPAddress
    }
    if ($Mailbox.RetentionHoldEnabled) {
        write-host " Removing Retention Hold for mailbox: $($Mailbox.primarySMTPAddress)"
        Set-Mailbox -Identity $Mailbox.primarySMTPAddress -RetentionHoldEnabled $false
        $report += $Mailbox.primarySMTPAddress
    }
    if ($Mailbox.ComplianceTagHoldApplied) {
        write-host " Removing Compliance Tag Hold for mailbox: $($Mailbox.primarySMTPAddress)"
        Set-Mailbox -Identity $Mailbox.primarySMTPAddress -RemoveComplianceTagHoldApplied -ProvideConsent
        $report += $Mailbox.primarySMTPAddress
    }
}
write-host "Removing holds for mailboxes completed."
$report = $report | Sort-Object -Unique
write-host "Total mailboxes processed: $($report.count)"
write-host "Time taken: $($stopwatch.Elapsed.ToString())"

$OutputFile = Join-Path -Path $outputFolder -ChildPath "RemovedHolds_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$report | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
