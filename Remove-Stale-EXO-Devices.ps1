#######################################################################################################################
# Get-EXO-Devices-Reports
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder		= "exo-mobile-device-mgmt"
$LogFilePrefix  = "remove-stale-exo-devices"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$DB_changed = $false
$ToBeDeletedRecords = @()
$CutoffDays = 350
$StaleDeviceCutoffDate = (Get-Date).AddDays(-$CutoffDays)

#######################################################################################################################

. $ScriptPath\include-Script-StartLog-Generic.ps1
Write-Log "Cutoff days: $($CutoffDays)"
Write-Log "Stale device cutoff date: $($StaleDeviceCutoffDate.ToString("yyyy-MM-dd"))"

# load DB mailbox-mgmt-db from file or initialize empty
if (Test-Path $DBFileEXOMobileDeviceMgmt) {
    Try {
        $EXODeviceMgmt_DB = Import-Clixml -Path $DBFileEXOMobileDeviceMgmt
        Write-Log "DB file $($DBFileEXOMobileDeviceMgmt) imported successfully, $($EXODeviceMgmt_DB.count) records read"
    } 
    Catch {
        Write-Log "Error importing $($DBFileEXOMobileDeviceMgmt), creating empty DB" -MessageType "Error"
        [hashtable]$EXODeviceMgmt_DB = @{}
        $DB_changed = $true
    }
}
else {
    Write-Log "DB file $($DBFileEXOMobileDeviceMgmt) not found, creating empty DB" -MessageType "Error"
    [hashtable]$EXODeviceMgmt_DB = @{}
    $DB_changed = $true
}

# retrieve devices ##################################### 
Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 120
write-host "Get-MobileDevice..." -NoNewline
[array]$Devices = Get-MobileDevice -ResultSize Unlimited
write-host "done"
Write-Log "Total devices read: $($Devices.Count)"

$CounterDeleted = 0
ForEach ($Device in $Devices) {
    $CurrentDBRecord = $Stat = $null
    $FirstSyncTime = $LastAttemptSync = $LastSuccessSync = $DaysSinceFirstSync = $DaysSinceLastSuccessSync = [DateTime]::MinValue
    $DBLastSuccessSync = [DateTime]::MinValue

    if ($EXODeviceMgmt_DB.ContainsKey($Device.DeviceId)) {
        $CurrentDBRecord = $EXODeviceMgmt_DB[$Device.DeviceId]
        Try {
            $DBLastSuccessSync = [DateTime]::Parse($CurrentDBRecord.LastSuccessSync)
        }
        Catch {
            $DBLastSuccessSync = [DateTime]::MinValue
        }
        if ($DBLastSuccessSync -gt $StaleDeviceCutoffDate) {
            #Write-Host "Skipping device $($Device.Identity), already processed on $($CurrentDBRecord.processedDate)"
            continue
        }
    }

    Try {
        $stat = Get-MobileDeviceStatistics -Identity $Device.ExchangeObjectId
        $DeviceRecord = [PSCustomObject]@{
            deviceId = $Device.DeviceId
            deviceUserAgent = $Device.DeviceUserAgent
            identity = $Device.Identity
            whenCreated = $Device.WhenCreated
            firstSyncTime = $Device.FirstSyncTime
            lastPolicyUpdateTime = $stat.LastPolicyUpdateTime
            lastSyncAttemptTime = $stat.LastSyncAttemptTime
            lastSuccessSync = $stat.LastSuccessSync
            lastPingHeartbeat = $stat.LastPingHeartbeat
            exchangeObjectId = $Device.ExchangeObjectId
            processedDate = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
        }

        if ($EXODeviceMgmt_DB.ContainsKey($Device.DeviceId)) {
            $EXODeviceMgmt_DB[$Device.DeviceId] = $DeviceRecord
        }
        else {
            $EXODeviceMgmt_DB.Add($Device.DeviceId, $DeviceRecord)
        }
        $DB_changed = $true
    }
    Catch{
        Write-Log "ERROR retrieving statistics for device: $($Device.deviceId)" -MessageType "ERR"
        write-Log $_.Exception.Message -MessageType "ERR"
        continue
    }

    Try {
        $FirstSyncTime = [DateTime]::Parse($Device.FirstSyncTime)
    }
    Catch {
        $FirstSyncTime = [DateTime]::MinValue
    }

    Try {
        $LastAttemptSync = [DateTime]::Parse($stat.LastSyncAttemptTime)
    }
    Catch {
        $LastAttemptSync = [DateTime]::MinValue
    }

    Try {
        $LastSuccessSync = [DateTime]::Parse($stat.LastSuccessSync)
    }
    Catch {
        $LastSuccessSync = [DateTime]::MinValue
    }

    $DaysSinceFirstSync = ((Get-Date) - $FirstSyncTime).Days
    $DaysSinceLastSuccessSync = ((Get-Date) - $LastSuccessSync).Days

    if (($Device.WhenCreated -gt $StaleDeviceCutoffDate) -or ($LastSuccessSync -gt $StaleDeviceCutoffDate) -or ($LastAttemptSync -gt $StaleDeviceCutoffDate) -or ($FirstSyncTime -gt $StaleDeviceCutoffDate)) {
        write-host "Device $($Device.deviceId): FirstSyncTime=$FirstSyncTime ($DaysSinceFirstSync days), LastSuccessSync=$LastSuccessSync ($DaysSinceLastSuccessSync days)" -ForegroundColor "Green"
    }
    else{
        #stale device, proceed to remove
        Try{
            Remove-MobileDevice -Identity $Device.ExchangeObjectId -Confirm:$false
            $ToBeDeletedRecords += $Device.DeviceId
            $CounterDeleted++
            Write-Log "Removed device: $($Device.deviceId) $($Device.Identity) - FirstSync $($FirstSyncTime) LastSuccessSync $($LastSuccessSync) LastAttemptSync $($LastAttemptSync)" -ForegroundColor Yellow
        }
        Catch{
            Write-Log "ERROR removing device: $($Device.deviceId)" -MessageType "ERR"
        }
    }
}

if ($ToBeDeletedRecords.Count -gt 0) {
	Write-Log "Deleting $($ToBeDeletedRecords.Count) records from DB"
    Write-Log "Before: $($EXODeviceMgmt_DB.Count)"
	foreach ($Key in $ToBeDeletedRecords) {
		$EXODeviceMgmt_DB.Remove($Key)
	}
    $DB_changed = $true
    Write-Log "After: $($EXODeviceMgmt_DB.Count)"
}

<#
if ($EXODeviceMgmt_DB.Count -ne $Devices.Count)  {
    write-host "DB records count ($($EXODeviceMgmt_DB.Count)) different from devices count ($($Devices.Count))"
    Write-Host "Cleaning up expired records from DB..."
    #delete expired records from DB
    foreach ($Key in $EXODeviceMgmt_DB.Keys) {   
        if (-not ($Devices.deviceId -contains $Key)) {
            $ToBeDeletedRecords += $Key
        }
    }
}

#>

#saving DB XML if needed
if (($EXODeviceMgmt_DB.count -gt 0) -and ($DB_changed)){
    Try {
        $EXODeviceMgmt_DB | Export-Clixml -Path $DBFileEXOMobileDeviceMgmt
        Write-Log "DB file $($DBFileEXOMobileDeviceMgmt) exported successfully, $($EXODeviceMgmt_DB.count) records saved"
    }
    Catch {
        Write-Log "Error exporting $($DBFileEXOMobileDeviceMgmt)" -MessageType "Error"
    }
}

#######################################################################################################################
write-log "Total devices removed: $CounterDeleted"
. $ScriptPath\include-Script-EndLog-generic.ps1
