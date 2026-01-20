$EnableOnScreenLogging = $true
$ManualAuth = $false
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
$Stopwatch =  [system.diagnostics.stopwatch]::StartNew()
$daysBackOffset = 10

. $ScriptPath\include-Functions-Init.ps1
. $ScriptPath\include-Root-Vars.ps1

##################################################################################################

$LogFolder			= "exports"
$LogFilePrefix		= "guest-file-access-"
$LogFileFreq		= "Y"

$OutputFolder		= "spo\audit"
$OutputFilePrefix	= "guest-file-access-"
$OutputFileFreq   	= "YMD"

$OutputFileSuffixAll = ""
$OutputFileSuffixSNE = "-sensitive-noenc"

#setting unified log query parameters
$record = "SharePointFileOperation"
$operations = "FileAccessed,FileAccessedExtended,FileDownloaded,FileSyncDownloadedFull"
$resultSize = 1000
$intervalMinutes = 10
$errorSleep = 30
$stdSleep = 1

##################################################################################################

$LogPath            = Join-Path $root_log_folder $LogFolder
$LogFileName        = $LogFilePrefix + (GetTimestamp($LogFileFreq)) + ".log"

$OutputPath         = Join-Path $root_export_folder $OutputFolder

$OutputFileNameAll  = $OutputFilePrefix + $strYesterday + $OutputFileSuffixAll + ".csv"
$OutputFileNameSNE  = $OutputFilePrefix + $strYesterday + $OutputFileSuffixSNE + ".csv"

$OutputFileAll      = Join-Path $OutputPath $OutputFileNameAll
$OutputFileSNE      = Join-Path $OutputPath $OutputFileNameSNE

. $ScriptPath\include-Functions-Common.ps1
. $ScriptPath\include-Functions-Audit-Records.ps1

#setting date variables
[DateTime]$Start = ((Get-Date -Hour 0 -Minute 0 -Second 0)).AddDays(-1-$daysBackOffset)
[DateTime]$End = $Start.AddDays(1)

#setting guest suffix value
$guestUPNSuffix = "#ext#@"+$TenantName
$DoW = $Start.DayOfWeek

##################################################################################################
. $ScriptPath\include-appreg-CEZ_EXO_MBX_MGMT.ps1
. $ScriptPath\include-ConnectEXO.ps1

LogWrite -LogString "Manual auth: $($ManualAuth)"
LogWrite -LogString "RecordType: $($record)"
LogWrite -LogString "Operations: $($operations)"
LogWrite -LogString "PageSize: $($resultSize) records"
LogWrite -LogString "Query interval: $($intervalMinutes) minutes"
LogWrite -LogString "Query start: $($start)"
LogWrite -LogString "Query end:   $($end)"
LogWrite -LogString "Output file ALL: $($OutputFileAll)"
LogWrite -LogString "Output file SNE: $($OutputFileSNE)"
LogWrite -LogString "Guest UPN suffix: $($guestUPNSuffix)"

. $ScriptPath\include-appreg-CEZ_AAD_USR_REPORT.ps1
. $ScriptPath\include-GetMSALToken.ps1

$O365TeamGroup_DB = @{}
$O365TeamGroup_DB = ImportCSVtoHashDB $DBFileTeams "FilesFolderUrl"

$AADGuest_DB = @{}
$AADGuest_DB = ImportCSVtoHashDB $DBFileGuests "mail"

$currentStart = $start
$totalCount = 0
$indexError = $false
[array]$SPOAuditLogSNE = $null

while ($true) {
    Write-Host
    $currentEnd = $currentStart.AddMinutes($intervalMinutes)
    if ($currentEnd -gt $end) {
        $currentEnd = $end
    }

    if ($currentStart -eq $currentEnd) {
        break
    }

    Write-Host "Retrieving activities between " -ForegroundColor Green -NoNewline
    Write-Host "$($currentStart) " -ForegroundColor Yellow -NoNewline
    Write-Host "and " -ForegroundColor Green -NoNewline
    Write-Host "$($currentEnd) " -ForegroundColor Yellow -NoNewline
    Write-Host "($($DoW))" -ForegroundColor Green
    #setting SessionID for Search-UnifiedAuditLog
    $sessionID = "AuditLog_" + (Get-Date).ToString("yyyyMMddHHmmssfff")
    
    $currentCount = 0
    $pageCount = 0
    $currentLabeledAll = 0
    $currentLabeledSNE = 0        
    [array]$PartialSPOAuditLogAll = $null
    
    do {
        $resCount = 0
        $indexError = $false
        Write-Host "Search-UnifiedAuditLog (SessionID: " -ForegroundColor Gray -NoNewline
        Write-Host "$($sessionID)" -ForegroundColor White -NoNewline
        Write-Host ") Run time:$($Stopwatch.Elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Gray
        Try {
            $results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -RecordType $record -Operations $operations -SessionId $sessionID -SessionCommand ReturnLargeSet -ResultSize $resultSize
            $resCount = $results.Count
            $pageCount++
        }
        Catch {
	        Write-Host "Error: $($_.Exception.Message)"
            Exit
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
            }
            if ($indexError -and (($resCntFrst -eq 0) -or ($resCntLast -eq 0))) {
                $cntClr = "Red"
            }
            if (($pageCount -eq 1) -and (-not($indexError))) {
                Write-Host "Records:" -ForegroundColor Green -NoNewline
                Write-Host "$($resCntFrst) " -ForegroundColor Yellow -NoNewline
                Write-Host "Pages:" -ForegroundColor Green -NoNewline
                Write-Host "$([Math]::Ceiling($resCntFrst/$resultSize)) " -ForegroundColor Yellow
            }
           
            if (-not $indexError) {
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
                    #searching for activities by B2B guest accounts - UPN ends with #ext#@cezdata.onmicrosoft.com
                    #we are not interested in anything else
                    if ($result.UserIds.EndsWith($guestUPNSuffix)) {
                        $auditData = ConvertFrom-Json -InputObject $result.AuditData
                        if ($auditData.SensitivityLabelId) {
                            $sensitivityLabelName = $AIPLabelDB.Item($auditData.SensitivityLabelId)
                        }
                        if ($auditData.SiteUrl) {
                            $SiteURL = ($auditData.SiteUrl).Trim().ToLower()
                            if ($SiteURL.EndsWith("/")) {
                                $SiteURL = $SiteURL.Trim("/")
                            }
                            $currentTeam = $O365TeamGroup_DB.Item($SiteURL)
                            if ($currentTeam) {
                                switch ($currentTeam.Level) {
                                    "Team"      {$TeamName = $currentTeam.Level + ":" + $currentTeam.teamName}
                                    "Channel"   {$TeamName = $currentTeam.Level + ":" + $currentTeam.channelName}
                                    Default     {$TeamName = ""}
                                }
                                $SiteOwners = $currentTeam.Owners
                                $TeamId     = $currentTeam.TeamId
                            }
                        }
                        
                        $mail = GetMailFromGuestUPN $auditData.UserId
                        if ($AADGuest_DB.ContainsKey($mail)) {
                            [pscustomobject]$guestData = $AADGuest_DB[$mail]
                        }
                        
                        $auditObject = [pscustomobject]@{
                            CreationTime            = $auditData.CreationTime;
                            CorrelationId           = $auditData.CorrelationId;
                            Id                      = $auditData.Id;
                            EventSource             = $auditData.EventSource;
                            Workload                = $auditData.Workload;
                            OperationType           = $auditData.Operation;
                            UserId                  = $auditData.UserId;
                            Mail                    = $mail;
                            UserType                = Get-AuditUserTypeFromCode $auditData.UserType;
                            guestId                 = $guestData.userId;
                            displayName             = $guestData.displayName;
                            createdDateTime         = $guestData.createdDateTime;
                            createdBy               = $guestData.createdBy;
                            mailDomain              = $guestData.MailDomain;
                            extAADTenantId		    = $guestData.ExtAADTenantId;
                            extAADDisplayName	    = $guestData.ExtAADDisplayName;
                            extAADdefaultDomain	    = $guestData.ExtAADdefaultDomain
                            ClientIP                = $auditData.ClientIP;
                            RecordType              = Get-AuditRecordTypeFromCode $auditData.RecordType;
                            Version                 = $auditData.Version;
                            IsManagedDevice         = $auditData.IsManagedDevice;
                            ItemType                = $auditData.ItemType;
                            SensitivityLabelId      = $auditData.SensitivityLabelId;
                            SensitivityLabelName    = $sensitivityLabelName;
                            SensitivityLabelOwner   = $auditData.SensitivityLabelOwnerEmail;
                            SourceFileExtension     = $auditData.SourceFileExtension;
                            SiteURL                 = $SiteURL;
                            Team                    = $TeamName;
                            Site                    = $SiteName;
                            TeamId                  = $TeamId;
                            Owners                  = $SiteOwners;
                            SourceFileName          = $auditData.SourceFileName;
                            SourceRelativeUrl       = $auditData.SourceRelativeUrl;
                            ObjectId                = $auditData.ObjectId;
                            ListId                  = $auditData.ListId;
                            ListItemUniqueId        = $auditData.ListItemUniqueId;
                            WebId                   = $auditData.WebId;
                            ListBaseType            = $auditData.ListBaseType;
                            ListServerTemplate      = $auditData.ListServerTemplate;
                            SiteUserAgent           = $auditData.SiteUserAgent;
                            HighPrioMediaProc       = $auditData.HighPriorityMediaProcessing
                        }
                        $PartialSPOAuditLogAll += $auditObject
                        if ($null -ne $auditData.SensitivityLabelId) {
                            if ($AIPLabelDBSNE.Contains($auditData.SensitivityLabelId)) {
                                $SPOAuditLogSNE += $auditObject
                                $currentLabeledSNE++
                                Write-Host "$($result.CreationDate) - $(GetMailFromGuestUPN $auditData.UserId) - Label: $($sensitivityLabelName)" -ForegroundColor Red
                            }
                            else {
                                $currentLabeledAll++
                                Write-Host "$($result.CreationDate) - $(GetMailFromGuestUPN $auditData.UserId) - Label: $($sensitivityLabelName)" -ForegroundColor Cyan
                            }
                        }
                    
                    }
                }#foreach ($result in $results)
                
                $currentTotal = $resCntFrst
                $totalCount += $resCount
                $currentCount += $resCount
                
                #end reached successfully
                if ($currentTotal -eq $resIndLast) {
                    #exporting/appending B2B guest all audit events to output CSV file
                    Write-Host "Records:" -ForegroundColor Gray -NoNewline
                    Write-Host "$($currentTotal) " -ForegroundColor Yellow -NoNewline
                    Write-Host "Pages:" -ForegroundColor Gray -NoNewline
                    Write-Host "$($pageCount) " -ForegroundColor Yellow -NoNewline
                    Write-Host "Records/page:$($resultSize). " -ForegroundColor Gray -NoNewline
                    if ($PartialSPOAuditLogAll.Count -gt 0) {
                        Write-Host "Guest records:" -ForegroundColor Gray -NoNewline
                        Write-Host "$($PartialSPOAuditLogAll.Count) " -ForegroundColor Yellow -NoNewline
                        Write-Host "Labeled:" -ForegroundColor Gray -NoNewline
                        Write-Host "$($currentLabeledAll) " -ForegroundColor Yellow -NoNewline
                        Write-Host "Sensitive NoEnc:" -ForegroundColor Gray -NoNewline
                        Write-Host "$($currentLabeledSNE)" -ForegroundColor Red -NoNewline
                        Write-Host ". Next $($intervalMinutes) minute interval." -ForegroundColor Gray
                        $PartialSPOAuditLogAll | Export-Csv -Path $OutputFileAll -Append -NoTypeInformation
                    }
                    else {
                        Write-Host "Nothing to export. Next $($intervalMinutes) minute interval." -ForegroundColor Gray
                    }
                    Start-Sleep -s $stdSleep
                    break
                }
            
            }#if (-not $indexError)
            else {
                Write-Host "Index error occurred. Waiting $($errorSleep) seconds and then retrying." -ForegroundColor Red -NoNewline
                for ($sec = 1; $sec -le $errorSleep; $sec++) {
                    Start-Sleep -s 1
                    Write-Host "." -ForegroundColor Red -NoNewline
                }
                Write-Host
                break
            }
        }#if ($resCount -ne 0)
    }#do (inner loop)
    while ($resCount -ne 0)

    if (-not $indexError) {
        $currentStart = $currentEnd
    }
    #reauthentication to EXO if last auth older than 30 minutes
    if ($StopwatchEXO.Elapsed.TotalMinutes -ge 30) {
        . $ScriptPath\include-appreg-CEZ_EXO_MBX_MGMT.ps1
        . $ScriptPath\include-connectEXO.ps1
    }
}#while ($true)

#exporting/appending B2B guest SNE audit events to output CSV file
if ($SPOAuditLogSNE.Count -gt 0) {
    $SPOAuditLogSNE | Export-Csv -Path $OutputFileSNE -Append -NoTypeInformation
}


Write-Host "END. Total count: " -NoNewline
Write-Host "$($totalCount)" -ForegroundColor Yellow -NoNewline
Write-Host " Run time: $($Stopwatch.Elapsed.ToString('hh\:mm\:ss'))"

. $ScriptPath\include-Script-EndLog-generic.ps1