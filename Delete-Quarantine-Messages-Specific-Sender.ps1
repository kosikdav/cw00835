#######################################################################################################################
# Remove-Quarantine-Messages-Specific-Sender.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    [string]$Sender,
    [int]$DaysBack = 30,
    [int]$IntervalMinutes = 30
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Init.ps1

#######################################################################################################################

$LogFolder          = "exo-quarantine"
$LogFilePrefix      = "delete-quarantine-messages-specific-sender"
$LogFileFreq        = "YMD"

#######################################################################################################################

. $ScriptPath\include-Script-Start-Include.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Suffix $LogFileSuffix -Ext "log"

[array]$QuarantineMessages = @()

#######################################################################################################################

. $IncFile_StdLogStartBlock

if ($ProcessWholeBacklog) {
    write-log "ProcessWholeBacklog: $ProcessWholeBacklog" -ForegroundColor Green
    write-log "IntervalMinutes: $IntervalMinutes" -ForegroundColor Green
}

Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 60

#########################################################################################
# Process quarantined Messages

$EndTime = Get-Date
$MaxDate = $EndTime.AddDays(-$DaysBack)
$StartTime = $EndTime.AddMinutes(-$IntervalMinutes)
Do {
    $StartTime = $EndTime.AddMinutes(-$IntervalMinutes)
    write-interactive "Reading messages received between $($StartTime) and $($EndTime): " -NoNewLine
    [array]$CurrentQuarantineMessages  = Get-QuarantineMessage -ReleaseStatus "NOTRELEASED" -SenderAddress $Sender -StartReceivedDate $StartTime -EndReceivedDate $EndTime
    $QuarantineMessages += $CurrentQuarantineMessages
    write-interactive "$($CurrentQuarantineMessages.Count) ($($QuarantineMessages.Count))" -ForegroundColor Green
    $EndTime = $StartTime
} Until ($StartTime -lt $MaxDate)

write-log "Quarantine messages to process: $($QuarantineMessages.Count)" -ForegroundColor Green

$DeleteCounter = 0
foreach ($Message in $QuarantineMessages) {

    $EXOError = $False
    Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 30
    Try {
        Delete-QuarantineMessage -Identity $Message.Identity -Confirm:$false -ErrorAction Stop
        Write-Log "deleting id:$($Message.MessageId) received:$($Message.ReceivedTime) sender:$($Message.SenderAddress) recipient:$($Message.RecipientAddress -join ',')" -ForegroundColor Yellow
        $DeleteCounter++
    }
    Catch {
        Write-Log "Failed to delete message from quarantine: id:$($Message.MessageId) sender:$($Message.SenderAddress) recipient:$($Message.RecipientAddress -join ',')" -ForegroundColor Magenta
        Write-Interactive $_.Exception.Message -ForegroundColor Red
        $EXOError = $True            
    }
}
Write-Log "Total messages deleted from quarantine: $DeleteCounter" -ForegroundColor Green

#######################################################################################################################

. $IncFile_StdLogEndBlock
