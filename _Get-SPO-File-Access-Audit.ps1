#######################################################################################################################
# Get-SPO-File-Access-Audit
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1

#######################################################################################################################

$LogFolder			    = "exports"
$LogFilePrefix		    = "spo-file-access-audit"

$OutputFolder           = "spo\audit"
$OutputFilePrefix       = "file-access"

$OutputFileSuffixAll    = "guests-all"
$OutputFileSuffixSNE    = "guests-sensitive-noenc"
$OutputFileSuffixDEL    = "users-deleted-odfb"

$daysBackOffset = 0
If (-not [Environment]::UserInteractive) {
    $daysBackOffset = 0
}
$fileDateDayOffset = -1 - $daysBackOffset

#setting date variables
[DateTime]$Start = ((Get-Date -Hour 0 -Minute 0 -Second 0)).AddDays(-1 -$daysBackOffset)
[DateTime]$End = $Start.AddDays(1)
$DoW = $Start.DayOfWeek

#setting unified log query parameters
$record = "SharePointFileOperation"
$operations = "FileAccessed,FileAccessedExtended,FileDownloaded,FileSyncDownloadedFull"
$resultSize = 1000
$intervalMinutes = 5
$errorSleep = 30
$stdSleep = 1

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1
. $IncFile_AIP_labels
. $IncFile_Functions_Audit

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputFileAll = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixAll -FileDateDayOffset $fileDateDayOffset -Ext "csv"
$OutputFileSNE = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixSNE -FileDateDayOffset $fileDateDayOffset -Ext "csv"
$OutputFileDEL = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixDEL -FileDateDayOffset $fileDateDayOffset -Ext "csv"

[hashtable]$deletedUsers_DB = @{}
[int]$totalCount = 0
[int]$totalGuestsALL = 0
[int]$totalLabeledAll = 0
[int]$totalLabeledSNE = 0
[int]$totalAccessedDEL = 0
[int]$indexErrorCountTotal = 0

[array]$ReportSPOAuditLogAll = @()
[array]$ReportSPOAuditLogSNE = @()
[array]$ReportSPOAuditLogDEL = @()

$now = Get-Date
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
Write-Log "Output file SNE: $($OutputFileSNE)"
Write-Log "Output file DEL: $($OutputFileDEL)"
Write-Log "Guest UPN suffix: $($guestUPNSuffix)"

$AADUsers_DB = Import-CSVtoHashDB -Path $DBFileUsersMemMin -KeyName "UserPrincipalName"

$twoDaysAgo = (Get-Date).AddDays(-2).ToString("yyyy-MM-ddTHH:mm:ssZ")
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "directory/deletedItems/microsoft.graph.user"
$UriSelect = "userPrincipalName,deletedDateTime,department,companyName,displayName,mail,onPremisesSamAccountName,onPremisesUserPrincipalName"
$UriFilter = "deletedDateTime+le+$twoDaysAgo"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect -Filter $UriFilter
$deletedUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
#write-host $deletedUsers.count

#write-host "deletedUsers:$($deletedUsers.count)"
foreach ($user in $deletedUsers) {
    if (-not($user.userPrincipalName.EndsWith($GuestUPNSuffix,$CCIgnoreCase))) {
        $OriginalUPN = $user.userPrincipalName.Substring($DelUPNPrefixLength)
        $ODfBUPN = $OriginalUPN -replace '\.', '_' -replace '@', '_'     
        $deletedUserObject = [pscustomobject]@{
            deletedDateTime = $user.deletedDateTime;
            userPrincipalName = $user.onPremisesUserPrincipalName;
            ODfBUPN = $ODfBUPN;
            DisplayName = $user.displayName;
            Department = $user.department;
            Company = $user.companyName;
            Mail = $user.mail;
            SamAccountName = $user.onPremisesSamAccountName
        }
        $deletedUsers_DB.Add($ODfBUPN,$deletedUserObject)
    }
}
Write-Log "deletedUsers_DB: $($deletedUsers_DB.count)"

$O365TeamGroup_DB = Import-CSVtoHashDB -Path $DBFileTeamsChannelsOwners -Keyname "FilesFolderUrl"
$AADGuest_DB = Import-CSVtoHashDB -Path $DBFileGuests -KeyName "mail"

while ($true) {
    [int]$currentCount = 0
    [int]$pageCount = 0
    [int]$currentLabeledAll = 0
    [int]$currentLabeledSNE = 0
    [int]$currentAccessedDEL = 0
    [int]$indexErrorCountCycle = 0
    $sessionID = [Guid]::NewGuid().ToString()
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
                    . $IncFile_StdLogEndBlock
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
                    
                    $upn = $result.UserIds.Trim().ToLower()
                    #write-host $result.UserIds -ForegroundColor Yellow
                    $SiteURL = $null
                    $auditData = ConvertFrom-Json -InputObject $result.AuditData
                    if ($auditData.SiteUrl) {
                        $SiteURL = ($auditData.SiteUrl).Trim().ToLower().TrimEnd("/")
                    }

                    #deleted users
                    if ($SiteUrl) {
                        if (($SiteUrl.StartsWith($RootODfBURL)) -and (-not($auditData.UserId.Contains("app@sharepoint")))) {
                            $UrlUpn = $SiteURL.Substring($RootODfBURL.Length + 1)
                            if ($UrlUpn.IndexOf("/") -gt 0) {
                                $UrlUpn = $UrlUpn.Substring(0,$UrlUpn.IndexOf("/"))
                            }
                            if ($deletedUsers_DB.ContainsKey($UrlUpn)) {
                                $deletedUser = $deletedUsers_DB[$UrlUpn]
                                $DaysSinceDeleted = (New-TimeSpan -Start $deletedUser.deletedDateTime -End $now).Days
                                $auditObjectDEL = [pscustomobject]@{
                                    DateTime        = $auditData.CreationTime;
                                    UserId          = $auditData.UserId;
                                    ClientIP        = $auditData.ClientIP;
                                    OperationType   = $auditData.Operation;
                                    URL             = $SiteURL;
                                    File            = $auditData.SourceFileName;
                                    ObjectId        = $auditData.ObjectId;
                                    DelUsr_UPN          = $deletedUser.userPrincipalName;
                                    DelUsr_DeletedDate  = $deletedUser.deletedDateTime;
                                    DelUsr_DaysSinceDel = $DaysSinceDeleted;
                                    DelUsr_DisplayName  = $deletedUser.DisplayName;
                                    DelUsr_Company      = $deletedUser.Company;
                                    DelUsr_Department   = $deletedUser.Department;
                                    DelUsr_Mail         = $deletedUser.Mail;
                                    DelUsr_KPJM         = $deletedUser.SamAccountName
                                }
                                $ReportSPOAuditLogDEL += $auditObjectDEL
                                $totalAccessedDEL++
                                $currentAccessedDEL++
                                write-host "$($auditData.UserId) - file:$($auditData.SourceFileName) owner:$($deletedUser.userPrincipalName) (deleted:$($deletedUser.deletedDateTime))" -ForegroundColor Yellow -BackgroundColor Red
                            }
                        }
                    }
                                        
                    #searching for activities by B2B guest accounts - UPN ends with #ext#@xyz.onmicrosoft.com
                    if ($upn.Contains($GuestUPNSuffix)) {
                        #write-host $result.UserIds -ForegroundColor Magenta
                        $totalGuestsALL++
                        if ($auditData.SensitivityLabelId) {
                            if ($AIPLabelDB.ContainsKey($auditData.SensitivityLabelId)) {
                                $sensitivityLabelName = $AIPLabelDB.Item($auditData.SensitivityLabelId)
                            }
                            else {
                                $sensitivityLabelName = "UNKNOWN"
                            }
                        }
                        if ($SiteUrl) {
                            $TeamName = $ChannelName = $TeamId = $SiteOwners = [string]::Empty
                            $currentTeam = $O365TeamGroup_DB.Item($SiteURL)
                            if ($currentTeam) {
                                $TeamName = $currentTeam.teamName
                                $SiteOwners = $currentTeam.Owners
                                $TeamId     = $currentTeam.TeamId
                                if ($currentTeam.Level -eq "Channel") {
                                    $ChannelName = $currentTeam.channelName
                                }
                            }
                        }
                        
                        $mail = Get-MailFromGuestUPN -GuestUPN $auditData.UserId
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
                            extAADdefaultDomain	    = $guestData.ExtAADdefaultDomain;
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
                            TeamName                = $TeamName;
                            ChannelName             = $ChannelName;
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
                        $ReportSPOAuditLogAll += $auditObject
                        if ($null -ne $auditData.SensitivityLabelId) {
                            $totalLabeledAll++
                            $currentLabeledAll++
                            if ($AIPLabelDBSNE.Contains($auditData.SensitivityLabelId)) {
                                $ReportSPOAuditLogSNE += $auditObject
                                $totalLabeledSNE++
                                $currentLabeledSNE++
                                #Write-Host "$($result.CreationDate) - $(GetMailFromGuestUPN $auditData.UserId) - Label: $($sensitivityLabelName)" -ForegroundColor Red
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

Write-Log "totalGuestsALL: $($totalGuestsALL) $($ReportSPOAuditLogAll.count)"
Write-Log "totalLabeledAll: $($totalLabeledAll)"
Write-Log "totalLabeledSNE: $($totalLabeledSNE) $($ReportSPOAuditLogSNE.count)"
Write-Log "totalAccessedDEL: $($totalAccessedDEL) $($ReportSPOAuditLogDEL.count)"
Write-Log "END. Total count: " -NoNewline
Write-Log "$($totalCount)" -ForegroundColor Yellow
Write-Log "Index errors: $($indexErrorCountTotal)"

Export-Report "all SPO file access" -Report $ReportSPOAuditLogAll -Path $OutputFileAll
Export-Report "guests SNE file access" -Report $ReportSPOAuditLogSNE -Path $OutputFileSNE
Export-Report "deleted ODfB file access" -Report $ReportSPOAuditLogDEL -Path $OutputFileDEL

Write-log "Deleted ODfB files access summary report:"
$DelODfBUsers = $ReportSPOAuditLogDEL.userId | Sort-Object -Unique
foreach ($UPN in $DelODfBUsers) {
    [array]$accessedFilesReport = @()
    foreach ($Line in $ReportSPOAuditLogDEL) {
        if ($Line.userId -eq $UPN) {
            $reportObject = [pscustomobject]@{
                DateTime        = $Line.DateTime;
                File            = $Line.SourceFileName;
                ObjectId        = $Line.ObjectId;
                DelUsr_UPN          = $Line.DelUsr_UPN;
                DelUsr_DeletedDate  = $Line.DelUsr_DeletedDate;
                DelUsr_DaysSinceDel = $Line.DelUsr_DaysSinceDel;
                DelUsr_DisplayName  = $Line.DelUsr_DisplayName;
                DelUsr_Company      = $Line.DelUsr_Company;
                DelUsr_Department   = $Line.DelUsr_Department;
                DelUsr_Mail         = $Line.DelUsr_Mail;
                DelUsr_KPJM         = $Line.DelUsr_KPJM
            }
            $accessedFilesReport += $reportObject
        }
        write-log "UPN: $($UPN) - accessed files: $($accessedFilesReport.count)"
    }
}

#######################################################################################################################

. $IncFile_StdLogEndBlock
