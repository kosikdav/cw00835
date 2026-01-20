#######################################################################################################################
# Get-PowerApps-Audit
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1

#######################################################################################################################

$LogFolder			    = "exports"
$LogFilePrefix		    = "powerapps-audit"

$OutputFolder           = "power-platform\audit"
$OutputFilePrefix       = "powerplat-audit"
$OutputFileSuffix       = "powerapps"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1
. $IncFile_AIP_labels
. $IncFile_Functions_Audit

$daysBackOffset = 0

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffix -FileDateDayOffset (-$daysBackOffset-1) -Ext "csv"

#setting unified log query parameters
$recordType = "PowerAppsApp"
$operations = ""
$resultSize = 1000
$intervalMinutes = 60
$errorSleep = 30
$stdSleep = 1

#setting date variables

[DateTime]$Start = ((Get-Date -Hour 0 -Minute 0 -Second 0)).AddDays(-1 -$daysBackOffset)
[DateTime]$End = $Start.AddDays(1)

$DoW = $Start.DayOfWeek

[int]$totalCount = 0
[int]$indexErrorCountTotal = 0
$currentStart = $start
[array]$PowerAppsAuditLog = @()

#######################################################################################################################

. $IncFile_StdLogStartBlock

Write-Log "Manual auth: $($ManualAuth)"
Write-Log "RecordType: $($recordType)"
Write-Log "Operations: $($operations)"
Write-Log "PageSize: $($resultSize) records"
Write-Log "Query interval: $($intervalMinutes) minutes"
Write-Log "Query start: $($start)"
Write-Log "Query end:   $($end)"
Write-Log "Output file: $($OutputFile)"

$PwrEnvironments_DB = Import-CSVtoHashDB -path $DBFilePwrEnvironments -keyname "EnvironmentId"

while ($true) {
    [int]$currentCount = 0
    [int]$pageCount = 0
    [int]$indexErrorCountCycle = 0

    $currentEnd = $currentStart.AddMinutes($intervalMinutes)
    if ($currentEnd -gt $end) {
        $currentEnd = $end
    }

    if ($currentStart -eq $currentEnd) {
        break
    }
    
    Write-Host
    Write-Host "Retrieving " -ForegroundColor Green -NoNewline
    Write-Host "$($recordType) " -ForegroundColor Cyan -NoNewline
    Write-Host "activities between " -ForegroundColor Green -NoNewline
    Write-Host "$($currentStart) " -ForegroundColor Yellow -NoNewline
    Write-Host "and " -ForegroundColor Green -NoNewline
    Write-Host "$($currentEnd) " -ForegroundColor Yellow -NoNewline
    Write-Host "($($DoW))" -ForegroundColor Green
    
    $sessionID = [Guid]::NewGuid().ToString()

    do {
        Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 60
        Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
        $resCount = 0
        $indexError = $false
        $queryError = $false
        Try {
            $results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -RecordType $recordType -SessionId $sessionID -SessionCommand ReturnLargeSet -ResultSize $resultSize -ErrorAction Stop -WarningAction Stop
            $resCount = $results.Count
            $pageCount++
        }
        Catch {
            $msg = $_.Exception.Message
            Write-Host "Exception: $($msg)" -ForegroundColor White -BackgroundColor DarkBlue
            Write-Host $_.Exception -ForegroundColor White -BackgroundColor DarkBlue
            if (($msg.Contains("errors during the search process")) -or ($msg.Contains("underlying connection was closed")) -or ($msg.Contains("operation has timed out"))) {
                $queryError = $true 
            }
            else {
                Exit
            }
        }
        if ($resCount -ne 0) {
            #storing resultIndex and ResultCount properties of first and last members of $results array
            $resIndFrst = $results[0].ResultIndex
            $resCntFrst = $results[0].ResultCount
            $resIndLast = $results[$resCount-1].ResultIndex
            $resCntLast = $results[$resCount-1].ResultCount
            if (($resIndFrst -eq -1) -or ($resIndLast -eq -1)) {
                $indexError = $true
                $indexErrorCountCycle++
                $indexErrorCountTotal++
                Write-Log "Search-UnifiedAuditLog index error - $($currentStart) - $($currentEnd) ($($DoW))" -MessageType "ERR"
                Write-Host "Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -RecordType $record -Operations $operations -SessionId $sessionID -SessionCommand ReturnLargeSet -ResultSize $resultSize -ErrorAction Stop -WarningAction Stop" -ForegroundColor Red
                If (($indexErrorCountTotal -gt $SearchUALIndexErrorMaxTotal) -or ($indexErrorCountCycle -gt $SearchUALIndexErrorMaxCycle)) {
                    Write-Log "Index error count exceeded maximum limit" -MessageType "ERR"
                    Exit
                }
            }
            
            if (-not ($indexError -or $queryError)) {
                foreach ($result in $results) {
                    $targetObjectName = [string]::Empty
                    $auditData = ConvertFrom-Json -InputObject $result.AuditData
                    $AdditionalInfo = ConvertFrom-Json -InputObject $auditData.AdditionalInfo 
                    if ($AdditionalInfo.targetObjectId) {
                        $targetObjectName = Get-GraphUserById -Id $AdditionalInfo.targetObjectId -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
                    }
                    #Write-Host $auditData.operation -ForegroundColor Green -NoNewline
                    #Write-Host $AdditionalInfo

                    $auditObject = [pscustomobject]@{
                        CreationTime            = $auditData.CreationTime;
                        Id                      = $auditData.Id;
                        OperationType           = $auditData.Operation;
                        UserId                  = $auditData.UserKey;
                        UserPrincipalName       = $auditData.UserId;
                        RecordType              = $auditData.RecordType;
                        Workload                = $auditData.Workload;
                        ClientIP                = $auditData.ClientIP;
                        AppName                 = $auditData.AppName;
                        environmentName         = $AdditionalInfo.environmentName;
                        environmentDisplayName  = $AdditionalInfo.environmentDisplayName;
                        resourceDisplayName     = $AdditionalInfo.resourceDisplayName;
                        targetObjectId          = $AdditionalInfo.targetObjectId;
                        targetObjectName        = $targetObjectName.userPrincipalName;
                        permissionType          = $AdditionalInfo.permissionType
                    }
                    $PowerAppsAuditLog += $auditObject
                }#foreach ($result in $results)
                
                $currentTotal = $resCntFrst
                $totalCount += $resCount
                $currentCount += $resCount
                
                #end reached successfully
                if ($currentTotal -eq $resIndLast) {
                    Start-Sleep -s $stdSleep
                    break
                }
            
            }#if (-not ($indexError -or $queryError))
            else {
                if ($indexError) {
                    Write-Host "INDEX ERROR OCCURED" -ForegroundColor Red
                    Write-host "$($resIndFrst)/$($resCntFrst) $($resIndLast)/$($resCntLast)" -ForegroundColor Red
                    Start-SleepDots "Waiting $($errorSleep) seconds and then retrying" -Seconds $errorSleep -ForegroundColor "Red"
                    break
                }
                if ($queryError) {
                    Write-Host "QUERY ERROR OCCURED" -ForegroundColor Red
                    Start-SleepDots "Waiting $($errorSleep) seconds and then retrying" -Seconds $errorSleep -ForegroundColor "Red"
                    Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 60 -ForceReconnect
                    break    
                }
            }
        }#if ($resCount -ne 0)
    }#do (inner loop)
    while ($resCount -ne 0)

    if (-not ($indexError -or $queryError)) {
        $currentStart = $currentEnd
    }
}#while ($true)

Write-Log "END. Total count: " -NoNewline
Write-Log "$($totalCount)" -ForegroundColor Yellow
Write-Log "Index errors: $($indexErrorCountTotal)"

Export-Report -Text "PowerAppsApp audit data" -Report $PowerAppsAuditLog -Path $OutputFile -SortProperty "CreationTime"

#######################################################################################################################

. $IncFile_StdLogEndBlock    
