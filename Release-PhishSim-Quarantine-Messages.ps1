#######################################################################################################################
# Release-PhishSim-Quarantine-Messages.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    [switch]$ProcessWholeBacklog
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Init.ps1

#######################################################################################################################

$LogFolder          = "cybeready"
$LogFilePrefix      = "release-phishsim-quarantine-messages"
$LogFileFreq        = "YMD"

#######################################################################################################################

. $ScriptPath\include-Script-Start-Include.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Suffix $LogFileSuffix -Ext "log"

$folder = "d:\scripts-m365\cezdata\"
$IntervalMinutes = 30
$PhishSim_SenderDomains_File    = $folder + "phishsim-sender-domains.txt"
$PhishSim_SenderIPs_File        = $folder + "phishsim-sender-IPs.txt"
[array]$QuarantineMessages = @()
[array]$ToBeDeletedRecords = @()
$DB_changed = $false
Function Get-PolicyConfigFileEntries {
    param(
        [string]$FilePath
    )
    if (Test-Path -Path $FilePath) {
        $entries = Get-Content -Path $FilePath
        $entries = $entries | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        $entries = $entries | ForEach-Object { $_.Trim() }
        return $entries | Sort-Object -Unique
    } else {
        write-log "File not found: $FilePath" -ForegroundColor Red
        return @()
    }
}

#######################################################################################################################

. $IncFile_StdLogStartBlock

if ($ProcessWholeBacklog) {
    write-log "ProcessWholeBacklog: $ProcessWholeBacklog" -ForegroundColor Green
    write-log "IntervalMinutes: $IntervalMinutes" -ForegroundColor Green
}

$PhishSimSenderDomains  = Get-PolicyConfigFileEntries -FilePath $PhishSim_SenderDomains_File
$PhishSimSenderIPs      = Get-PolicyConfigFileEntries -FilePath $PhishSim_SenderIPs_File

# load DB mailbox-mgmt-db from file or initialize empty
if (test-path $DBFileEXOQuarantineMgmt) {
    Try {
        $EXOQuarantineMgmt_DB = Import-Clixml -Path $DBFileEXOQuarantineMgmt
        Write-Log "DB file $($DBFileEXOQuarantineMgmt) imported successfully, $($EXOQuarantineMgmt_DB.count) records found"
    } 
    Catch {
        Write-Log "Error importing $($DBFileEXOQuarantineMgmt), creating empty DB" -MessageType "Error"
        [hashtable]$EXOQuarantineMgmt_DB = @{}
        $DB_changed = $true
    }
}
else {
    Write-Log "DB file $($DBFileEXOQuarantineMgmt) not found, creating empty DB" -MessageType "Error"
    [hashtable]$EXOQuarantineMgmt_DB = @{}
    $DB_changed = $true
}

Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 60

#########################################################################################
# Process quarantined Messages

If ($ProcessWholeBacklog) {
    write-log "Processing whole backlog of quarantined messages" -ForegroundColor Green
    $EndTime = Get-Date
    $MaxDate = $EndTime.AddDays(-30)
    $StartTime = $EndTime.AddMinutes(-$IntervalMinutes)
    Do {
        $StartTime = $EndTime.AddMinutes(-$IntervalMinutes)
        write-interactive "Reading messages received between $($StartTime) and $($EndTime): " -NoNewLine
        [array]$CurrentQuarantineMessages  = Get-QuarantineMessage -ReleaseStatus "NOTRELEASED" -PolicyType "HostedContentFilterPolicy" -StartReceivedDate $StartTime -EndReceivedDate $EndTime
        $QuarantineMessages += $CurrentQuarantineMessages
        write-interactive "$($CurrentQuarantineMessages.Count) ($($QuarantineMessages.Count))" -ForegroundColor Green
        $EndTime = $StartTime
    } Until ($StartTime -lt $MaxDate)
} else {
    write-log "Processing only top 100 quarantined messages" -ForegroundColor Green
    [array]$QuarantineMessages  = Get-QuarantineMessage -ReleaseStatus "NOTRELEASED" -PolicyType "HostedContentFilterPolicy"
}

write-log "Quarantine messages to process: $($QuarantineMessages.Count)" -ForegroundColor Green
$ReleaseCounter = 0
$IgnoredMessagesCount++

foreach ($Message in $QuarantineMessages) {
    if ($EXOQuarantineMgmt_DB.ContainsKey($message.Identity)) {
        Write-Interactive "skipping $($message.Identity)" -ForegroundColor Yellow -BackgroundColor DarkGray
        $IgnoredMessagesCount++
        continue
    }

    $EXOError = $False
    $HeaderIPs = $HeaderDomains = $null

    $MessageRecord = [PSCustomObject]@{
        Identity = $message.Identity
        Id = $message.MessageId
        ReceivedTime = $message.ReceivedTime
        processedDate = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
        ReleaseStatus = "IGNORED"
    }
    $DB_changed = $true
    
    Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 30

    Try {
        $Header = Get-QuarantineMessageHeader -Identity $Message.Identity -ErrorAction Stop
    }
    Catch {
        Write-Log "Failed to get message header: id:$($Message.MessageId) sender:$($Message.SenderAddress) recipient:$($Message.RecipientAddress)" -ForegroundColor Magenta
        Write-Interactive $_.Exception.Message -ForegroundColor Red
        $EXOError = $True
        Continue
    }
    
    $HeaderIPs = $PhishSimSenderIPs | Where-Object { $Header.Header -like "*$_*" }
    $HeaderDomains = $PhishSimSenderDomains | Where-Object { $Header.Header -like "*$_*" }
    
    if ($HeaderIPs -and $HeaderDomains) {
        Try {
            Release-QuarantineMessage -Identity $Message.Identity -ReleaseToAll -Force -Confirm:$false -ErrorAction Stop
            Write-Log "releasing id:$($Message.MessageId) received:$($Message.ReceivedTime) recipient:$($Message.RecipientAddress -join ',') matched IPs:$($HeaderIPs -join ',') matched domains:$($HeaderDomains -join ',')" -ForegroundColor Yellow
            $ReleaseCounter++
            $MessageRecord.ReleaseStatus = "RELEASED"
        }
        Catch {
            Write-Log "Failed to release message from quarantine: id:$($Message.MessageId) sender:$($Message.SenderAddress) recipient:$($Message.RecipientAddress)" -ForegroundColor Magenta
            Write-Interactive $_.Exception.Message -ForegroundColor Red
            $EXOError = $True            
        }
    }
    if (-not $EXOError) {
   		$EXOQuarantineMgmt_DB.Add($message.Identity, $MessageRecord)
    }
}
Write-Log "Total messages released from quarantine: $ReleaseCounter" -ForegroundColor Green

#find expired blobs in DB
[datetime]$date = (get-date).AddDays(-40)
foreach ($Identity in $EXOQuarantineMgmt_DB.Keys) {
    if ($EXOQuarantineMgmt_DB[$Identity].receivedTime -lt $date) {
        $ToBeDeletedRecords += $Identity
        Write-Interactive "Expired record: $($EXOQuarantineMgmt_DB[$Identity].identity) received: $($EXOQuarantineMgmt_DB[$Identity].receivedTime) processed: $($EXOQuarantineMgmt_DB[$Identity].processedDate) " -ForegroundColor Red
    }
}
Write-Log "Expired records in DB: $($ToBeDeletedRecords.Count)"

#delete expired blobs from DB
if ($ToBeDeletedRecords.Count -gt 0) {
	Write-Log "Deleting expired records from DB..."
	foreach ($Identity in $ToBeDeletedRecords) {
		$EXOQuarantineMgmt_DB.Remove($Identity)
		$DB_changed = $true
	}
}

#saving DB XML if needed
if (($EXOQuarantineMgmt_DB.count -gt 0) -and ($DB_changed)){
    Try {
        $EXOQuarantineMgmt_DB | Export-Clixml -Path $DBFileEXOQuarantineMgmt
        Write-Log "DB file $($DBFileEXOQuarantineMgmt) exported successfully, $($EXOQuarantineMgmt_DB.count) records saved"
    }
    Catch {
        Write-Log "Error exporting $($DBFileEXOQuarantineMgmt)" -MessageType "Error"
    }
}

#######################################################################################################################

. $IncFile_StdLogEndBlock
