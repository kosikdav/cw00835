#######################################################################################################################
# Get-Copilot-Audit
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    $SpecificDate,
    $MinusDays,
    [switch]$NoRollup
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1

if ($SpecificDate -and $MinusDays) {
    Write-Host "You can specify either -SpecificDate or -MinusDays parameter, not both."
    Exit
}

#######################################################################################################################

$LogFolder			    = "exports"
$LogFilePrefix		    = "copilot-audit"

$OutputFolder               = "copilot\audit"
$OutputFilePrefix           = "copilot-audit"
$OutputFileSuffix           = "interactions"
$OutputFileSuffixMonthly    = "interactions-monthly"

#setting unified log query parameters
$record = "CopilotInteraction"
$operations = "CopilotInteraction"
$resultSize = 1000
$intervalMinutes = 60
$errorSleep = 30
$stdSleep = 3

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1
. $IncFile_AIP_labels
. $IncFile_Functions_Audit


if ($MinusDays) {
    Try {
        $daysBackOffset = [int]$MinusDays
    }
    Catch {
        $daysBackOffset = 1
    }
} 
else {
    $daysBackOffset = 1
}

if ($SpecificDate) {
    Try {
        [DateTime]$Start = (Get-Date -Date $SpecificDate -Hour 0 -Minute 0 -Second 0)
        [DateTime]$End = $Start.AddDays(1)
    }
    Catch {
        [DateTime]$Start = ((Get-Date -Hour 0 -Minute 0 -Second 0)).AddDays(-1 -$daysBackOffset)
        [DateTime]$End = $Start.AddDays(1)
    }
}
else {
    [DateTime]$Start = ((Get-Date -Hour 0 -Minute 0 -Second 0)).AddDays(-1 -$daysBackOffset)
    [DateTime]$End = $Start.AddDays(1)
}

$FileDateCurrentMonth = Get-Date -Date $Start -UFormat "%Y-%m"
$DoW = $Start.DayOfWeek

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffix -FileDateDayOffset (-$daysBackOffset-1) -Ext "csv"
$OutputFileMonthly = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixMonthly -FileDateDayOffset (-$daysBackOffset-1) -Ext "csv" -Freq "YM"

#setting date variables

[int]$totalCount = 0
[int]$indexErrorCountTotal = 0
$currentStart = $start
[array]$CopilotAuditLog = @()

#######################################################################################################################

. $IncFile_StdLogStartBlock
Write-Log "Output file monthly: $($OutputFileMonthly)"

Write-Log "Manual auth: $($ManualAuth)"
Write-Log "RecordType: $($record)"
Write-Log "Operations: $($operations)"
Write-Log "PageSize: $($resultSize) records"
Write-Log "Query interval: $($intervalMinutes) minutes"
Write-Log "Query start: $($start)"
Write-Log "Query end:   $($end)"
Write-Log "Output file: $($OutputFile)"

$AADUSER_DB = @{}
$AADUSER_DB = Import-CSVtoHashDB -Path $DBFileUsersMemLic -KeyName "id"

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
    Write-Host "Retrieving activities between " -ForegroundColor Green -NoNewline
    Write-Host "$($currentStart) " -ForegroundColor Yellow -NoNewline
    Write-Host "and " -ForegroundColor Green -NoNewline
    Write-Host "$($currentEnd) " -ForegroundColor Yellow -NoNewline
    Write-Host "($($DoW))" -ForegroundColor Green
    #setting SessionID for Search-UnifiedAuditLog
    $sessionID = [Guid]::NewGuid().ToString()

    do {
        Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 60
        $resCount = 0
        $indexError = $false
        $queryError = $false
        Write-Host "Search-UnifiedAuditLog (SessionID: " -ForegroundColor Gray -NoNewline
        Write-Host "$($sessionID)" -ForegroundColor White -NoNewline
        Write-Host ") Run time:$($Stopwatch.Elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Gray
        Try {
            $results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -RecordType $record -SessionId $sessionID -SessionCommand ReturnLargeSet -ResultSize $resultSize -ErrorAction Stop -WarningAction Stop
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
            $indClr = "White"
            $cntClr = "White"
            $resIndFrst = $results[0].ResultIndex
            $resCntFrst = $results[0].ResultCount
            $resIndLast = $results[$resCount-1].ResultIndex
            $resCntLast = $results[$resCount-1].ResultCount
            if (($resIndFrst -eq -1) -or ($resIndLast -eq -1)) {
                $indexError = $true
                $indClr = "Red"
                $indexErrorCountCycle++
                $indexErrorCountTotal++
                Write-Log "Search-UnifiedAuditLog index error - $($currentStart) - $($currentEnd) ($($DoW))" -MessageType "ERR"
                Write-Host "Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -RecordType $record -Operations $operations -SessionId $sessionID -SessionCommand ReturnLargeSet -ResultSize $resultSize -ErrorAction Stop -WarningAction Stop" -ForegroundColor Red
                If (($indexErrorCountTotal -gt $SearchUALIndexErrorMaxTotal) -or ($indexErrorCountCycle -gt $SearchUALIndexErrorMaxCycle)) {
                    Write-Log "Index error count exceeded maximum limit" -MessageType "ERR"
                    Exit
                }            }
            if ($indexError -and (($resCntFrst -eq 0) -or ($resCntLast -eq 0))) {
                $cntClr = "Red"
            }
            if (($pageCount -eq 1) -and (-not($indexError))) {
                Write-Host "Records:" -ForegroundColor Green -NoNewline
                Write-Host "$($resCntFrst) " -ForegroundColor Yellow -NoNewline
                Write-Host "Pages:" -ForegroundColor Green -NoNewline
                Write-Host "$([Math]::Ceiling($resCntFrst/$resultSize)) " -ForegroundColor Yellow
            }
            
            if (-not ($indexError -or $queryError)) {
                Write-Host "Page:" -ForegroundColor Gray -NoNewline
                Write-Host "$($pageCount)/$([Math]::Ceiling($resCntFrst/$resultSize)) " -ForegroundColor  Yellow -NoNewline
                Write-Host "Records:$($resCount) (First:" -ForegroundColor Gray -NoNewline
                Write-Host "$($resIndFrst)" -ForegroundColor $indClr  -NoNewline
                Write-Host "/" -ForegroundColor Gray -NoNewline
                Write-Host "$($resCntFrst)" -ForegroundColor $cntClr -NoNewline
                Write-Host ") (Last:" -ForegroundColor Gray -NoNewline
                Write-Host "$($resIndLast)" -ForegroundColor $indClr -NoNewline
                Write-Host "/" -ForegroundColor Gray -NoNewline
                Write-Host "$($resCntLast)" -ForegroundColor $cntClr -NoNewline
                Write-Host ")" -ForegroundColor Gray

                foreach ($result in $results) {
                    $ContextArray = $ResourceArray = @()
                    $Contexts = $AccessedResources = $AISystemPlugin = [string]::Empty
                    $user = $null

                    $auditData = ConvertFrom-Json -InputObject $result.AuditData

                    if ($auditData.CopilotEventData.Contexts) {
                        foreach ($context in $auditData.CopilotEventData.Contexts) {
                            $ContextArray += "[Id=" + $Context.id + "," + "Type=" + $context.type + "]"
                        }
                        $Contexts = $ContextArray -join ";"
                    }
                    if ($auditData.CopilotEventData.AccessedResources) {
                        foreach ($Resource in $auditData.CopilotEventData.AccessedResources) {
                            $ResourceArray += "[" + $Resource.Name + "]"
                        }
                        $AccessedResources = $ResourceArray -join ";"
                    }
                    if ($auditData.UserKey) {
                        $user = $AADUSER_DB[$auditData.UserKey]
                    }
                    
                    $AISystemPlugin = $($auditData.CopilotEventData.AISystemPlugin.name + " (" + $auditData.CopilotEventData.AISystemPlugin.id + ")")

                    $auditObject = [pscustomobject]@{
                        CreationTime            = $auditData.CreationTime
                        Id                      = $auditData.Id
                        AISystemPlugin          = $AISystemPlugin
                        OperationType           = $auditData.Operation
                        UserId                  = $auditData.UserKey
                        UserPrincipalName       = $auditData.UserId
                        UserDisplayName         = $user.DisplayName
                        CompanyName             = $user.CompanyName
                        Department              = $user.Department
                        CopilotLicense          = $user.CopilotLicense
                        RecordType              = $auditData.RecordType
                        Workload                = $auditData.Workload
                        ClientIP                = $auditData.ClientIP
                        AppHostName             = $auditData.CopilotEventData.AppHost
                        Contexts                = $Contexts
                        AccessedResources       = $AccessedResources
                        ThreadId                = $auditData.CopilotEventData.ThreadId
                        ModelName               = $auditData.CopilotEventData.ModelTransparencyDetails.ModelName
                    }
                    $CopilotAuditLog += $auditObject
                }#foreach ($result in $results)
                
                $currentTotal = $resCntFrst
                $totalCount += $resCount
                $currentCount += $resCount
                
                #end reached successfully
                if ($currentTotal -eq $resIndLast) {
                    #exporting/appending B2B guest all audit events to output CSV file
                    Write-Host "Total processed records:" -ForegroundColor Gray -NoNewline
                    Write-Host "$($currentTotal) " -ForegroundColor Yellow -NoNewline
                    Write-Host "Pages:" -ForegroundColor Gray -NoNewline
                    Write-Host "$($pageCount) " -ForegroundColor Yellow -NoNewline
                    Write-Host "Records/page:$($resultSize)" -ForegroundColor Gray
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

Export-Report -Text "copilot audit data" -Report $CopilotAuditLog -Path $OutputFile -SortProperty "CreationTime"

#######################################################################################################################

if (-not ($NoRollup)) {
    Write-Log "Monthly rollup:" -ForegroundColor Yellow

    $Path = [System.IO.Path]::Combine($ROF,$OutputFolder,"*.csv")
    $Include = "*.csv"
    $Exclude = "*$($OutputFileSuffixMonthly)*"
    $ExistingCSVFiles = Get-ChildItem -Path $Path -Include $Include -Exclude $Exclude -Filter "*$($FileDateCurrentMonth)*"
    $getFirstLine = $true
    $totalLines = 0
    Remove-Item -Path $OutputFileMonthly -ErrorAction SilentlyContinue

    foreach ($file in $ExistingCSVFiles) {
        $lines = Get-Content $file.FullName      
        $linesToWrite = switch($getFirstLine) {
            $true  {$lines}
            $false {$lines | Select-Object -Skip 1}
        }
        $getFirstLine = $false    
        Write-Log "$($file.FullName) - $($linesToWrite.Count) lines" 
        $totalLines += $linesToWrite.Count
        Add-Content -Path $OutputFileMonthly -Value $linesToWrite -Encoding "UTF8"
    }
    Write-Log "Added $($totalLines) to $($OutputFileMonthly)"
}

#######################################################################################################################

. $IncFile_StdLogEndBlock    
