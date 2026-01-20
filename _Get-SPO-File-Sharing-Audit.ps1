#######################################################################################################################
# Get-SPO-File-Sharing-Audit
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1

#######################################################################################################################

$LogFolder			    = "exports"
$LogFilePrefix		    = "spo-file-sharing-audit"

$OutputFolder           = "spo\audit"
$OutputFilePrefix       = "file-access"
$OutputFileSuffix       = "sharing"

$daysBackOffset = 10
If (-not [Environment]::UserInteractive) {
    $daysBackOffset = 0
}
$fileDateDayOffset = -1 - $daysBackOffset

#setting date variables
[DateTime]$Start = ((Get-Date -Hour 0 -Minute 0 -Second 0)).AddDays(-1 -$daysBackOffset)
[DateTime]$End = $Start.AddDays(1)
$DoW = $Start.DayOfWeek

#setting unified log query parameters
$record = "SharePointSharingOperation"
$operations = "CompanyLinkCreated,SharingSet,SecureLinkCreated,AddedToSecureLink,SharingInvitationCreated"
$resultSize = 1000
$intervalMinutes = 60
$errorSleep = 30
$stdSleep = 1

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1
. $IncFile_AIP_labels
. $IncFile_Functions_Audit

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputFileAll = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffix -FileDateDayOffset $fileDateDayOffset -Ext "csv"

[array]$SPOSharingAuditLogAll = @()
[int]$totalCount = 0
[int]$totalGuestsALL = 0
[int]$indexErrorCountTotal = 0
$currentStart = $start

#######################################################################################################################

. $IncFile_StdLogStartBlock

Write-Log "Manual auth: $($ManualAuth)"
Write-Log "RecordType: $($record)"
Write-Log "Operations: $($operations)"
Write-Log "PageSize: $($resultSize) records"
Write-Log "Query interval: $($intervalMinutes) minutes"
Write-Log "DaysBackOffset: $($daysBackOffset)"
Write-Log "Query start: $($start)"
Write-Log "Query end:   $($end)"
Write-Log "Output file ALL: $($OutputFileAll)"
Write-Log "Guest UPN suffix: $($guestUPNSuffix)"

$O365TeamGroup_DB = Import-CSVtoHashDB -Path $DBFileTeamsChannelsOwners -Keyname "FilesFolderUrl"
$AADGuest_DB = Import-CSVtoHashDB -Path $DBFileGuests -KeyName "UserPrincipalName"

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
    #$sessionID = "AuditLog_" + (Get-Date).ToString("yyyyMMddHHmmssfff")
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
            $results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -RecordType $record -Operations $operations -SessionId $sessionID -SessionCommand ReturnLargeSet -ResultSize $resultSize -ErrorAction Stop -WarningAction Stop
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
                }
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
                    #if ($auditData.userid -eq "app@sharepoint") {
                        #continue
                    #}
                    #write-host $result.UserIds -ForegroundColor Yellow
                    $guestData = $null
                    $SiteURL = $TeamName = $ChannelName = $targetUserMail = $sensitivityLabelName = [string]::Empty
                    $auditData = ConvertFrom-Json -InputObject $result.AuditData
                    if ($auditData.SiteUrl) {
                        $SiteURL = ($auditData.SiteUrl).Trim().ToLower().TrimEnd("/")
                    }
                    if ($auditData.SensitivityLabelId) {
                        if ($AIPLabelDB.ContainsKey($auditData.SensitivityLabelId)) {
                            $sensitivityLabelName = $AIPLabelDB.Item($auditData.SensitivityLabelId)
                        }
                        else {
                            $sensitivityLabelName = "UNKNOWN"
                        }
                    }

                    #searching for activities by B2B guest accounts - UPN ends with #ext#@xyz.onmicrosoft.com

                        if ($SiteUrl) {
                            $currentTeam = $O365TeamGroup_DB.Item($SiteURL)
                            if ($currentTeam) {
                                $TeamName = $currentTeam.teamName
                                if ($currentTeam.Level -eq "Channel") {
                                    $ChannelName = $currentTeam.channelName
                                }
                            }
                        }
                        
                        if ($auditData.TargetUserOrGroupName) {
                            if ($auditData.TargetUserOrGroupName.Contains("@")) {
                                $TargetUserMail = $auditData.TargetUserOrGroupName.Trim().ToLower()
                            }
                            if ($AADGuest_DB.ContainsKey($auditData.TargetUserOrGroupName)) {
                                [pscustomobject]$guestData = $AADGuest_DB[$auditData.TargetUserOrGroupName]
                                $totalGuestsALL++
                                $targetUserMail = $guestData.mail.ToLower()
                            }   
                        }

                        $auditObject = [pscustomobject]@{
                            CreationTime            = $auditData.CreationTime;
                            CorrelationId           = $auditData.CorrelationId;
                            Id                      = $auditData.Id;
                            EventSource             = $auditData.EventSource;
                            Workload                = $auditData.Workload;
                            OperationType           = $auditData.Operation;
                            UserId                  = $auditData.UserId;
                            UserType                = Get-AuditUserTypeFromCode $auditData.UserType;

                            ClientIP                = $auditData.ClientIP;
                            AuthenticationType      = $auditData.AuthenticationType;
                            BrowserName             = $auditData.BrowserName;
                            BrowserVersion          = $auditData.BrowserVersion;
                            Platform                = $auditData.Platform;
                            RecordType              = Get-AuditRecordTypeFromCode $auditData.RecordType;
                            Version                 = $auditData.Version;
                            IsManagedDevice         = $auditData.IsManagedDevice;
                            ItemType                = $auditData.ItemType;
                            SensitivityLabelId      = $auditData.SensitivityLabelId;
                            SensitivityLabelName    = $sensitivityLabelName;
                            SourceFileExtension     = $auditData.SourceFileExtension;
                            SiteURL                 = $auditData.SiteUrl;
                            TeamName                = $TeamName;
                            ChannelName             = $ChannelName;
                            SourceFileName          = $auditData.SourceFileName;
                            SourceRelativeUrl       = $auditData.SourceRelativeUrl;
                            ObjectId                = $auditData.ObjectId;
                            ListId                  = $auditData.ListId;
                            ListItemUniqueId        = $auditData.ListItemUniqueId;
                            WebId                   = $auditData.WebId;
                            UserAgent               = $auditData.UserAgent;
                            TargetUserType          = $auditData.TargetUserOrGroupType;
                            TargetName              = $auditData.TargetUserOrGroupName;
                            TargetMail              = $TargetUserMail;
                            guestId                 = $guestData.Id;
                            guestDisplayName        = $guestData.displayName;
                            guestCreatedDateTime    = $guestData.createdDateTime;
                            guestCreatedBy          = $guestData.employeeType;
                            guestMailDomain         = $guestData.MailDomain;
                            guestExtAADTenantId		= $guestData.ExtAADTenantId;
                            guestExtAADDisplayName	= $guestData.ExtAADDisplayName
                        }
                        $SPOSharingAuditLogAll += $auditObject
                    
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
                    #Write-host "$([char]36)results[0].ResultIndex = $($resIndFrst)" -ForegroundColor Red
                    #Write-host "$([char]36)results[0].ResultCount = $($resCntFrst)" -ForegroundColor Red
                    #Write-host "$([char]36)results[$($resCount-1)].ResultIndex = $($resIndLast)" -ForegroundColor Red
                    #Write-host "$([char]36)results[$($resCount-1)].ResultCount = $($resCntLast)" -ForegroundColor Red
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

Write-Log "totalGuestsALL: $($totalGuestsALL)"
Write-Log "$($totalCount)" -ForegroundColor Yellow
Write-Log "Index errors: $($indexErrorCountTotal)"

Export-Report -Text "SPOSharingAuditLogAll report" -Report $SPOSharingAuditLogAll -Path $OutputFileAll


#######################################################################################################################

. $IncFile_StdLogEndBlock
