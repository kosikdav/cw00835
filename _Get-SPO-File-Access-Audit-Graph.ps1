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
$LogFilePrefix		    = "spo-file-access-audit-graph"

$OutputFolder           = "spo\audit"
$OutputFilePrefix       = "file-access-graph"

$OutputFileSuffixAll    = "guests-all"
$OutputFileSuffixSNE    = "guests-sensitive-noenc"
$OutputFileSuffixDEL    = "users-deleted-odfb"

$daysBackOffset = 0
If (-not [Environment]::UserInteractive) {
    $daysBackOffset = 0
}
$fileDateDayOffset = -1 - $daysBackOffset

#setting date variables
[datetime]$Start = ((Get-Date -Hour 0 -Minute 0 -Second 0)).AddDays(-1 -$daysBackOffset)
[datetime]$End = $Start.AddDays(1)
$StartUTC = $Start.ToString("yyyy-MM-ddTHH:mm:ssZ")
$EndUTC = $End.ToString("yyyy-MM-ddTHH:mm:ssZ")

#setting unified log query parameters
$DisplayName = "Audit_FileAccess" + $StartUTC + "_" + $EndUTC
$recordTypeFilters = @("SharePointFileOperation")
$operationFilters = @("FileAccessed","FileAccessedExtended","FileDownloaded","FileSyncDownloadedFull")

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1
. $IncFile_AIP_labels
. $IncFile_Functions_Audit

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputFileAll = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixAll -FileDateDayOffset $fileDateDayOffset -Ext "csv"
$OutputFileSNE = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixSNE -FileDateDayOffset $fileDateDayOffset -Ext "csv"
$OutputFileDEL = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixDEL -FileDateDayOffset $fileDateDayOffset -Ext "csv"
$OutputFileDbg = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix "dbg-access" -FileDateDayOffset $fileDateDayOffset -Ext "csv"

[hashtable]$deletedUsers_DB = @{}
[int]$totalCount = 0
[int]$totalGuestsALL = 0
[int]$totalLabeledAll = 0
[int]$totalLabeledSNE = 0
[int]$totalAccessedDEL = 0

[array]$ReportSPOAuditLogAll = @()
[array]$ReportSPOAuditLogSNE = @()
[array]$ReportSPOAuditLogDEL = @()
[array]$AuditRecords = @()

$now = Get-Date
$twoDaysAgo = (Get-Date).AddDays(-2).ToString("yyyy-MM-ddTHH:mm:ssZ")

#######################################################################################################################

. $IncFile_StdLogStartBlock

Write-Log "Manual auth: $($ManualAuth)"
Write-Log "Query displaName: $($DisplayName)"
Write-Log "recordTypeFilters: $($recordTypeFilters)"
Write-Log "operationFilters: $($operationFilters)"
Write-Log "DaysBackOffset: $($daysBackOffset)"
Write-Log "Query start: $($start) ($($start.DayOfWeek))" -ForegroundColor Yellow
Write-Log "Query end:   $($end)" -ForegroundColor Yellow
Write-Log "Output file ALL: $($OutputFileAll)"
Write-Log "Output file SNE: $($OutputFileSNE)"
Write-Log "Output file DEL: $($OutputFileDEL)"
Write-Log "Output file DBG: $($OutputFileDbg)" -ForegroundColor Cyan
Write-Log "Guest UPN suffix: $($guestUPNSuffix)"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

##############################################################################
# create audit log search
$UriResource = "security/auditLog/queries"
$URI = New-GraphUri -Resource $UriResource -Version "beta"
$SearchParameters = @{
    displayName	        = $DisplayName
    filterStartDateTime = $StartUTC
    filterEndDateTime 	= $EndUTC
    recordTypeFilters 	= $recordTypeFilters	
    operationFilters	= $operationFilters
} | ConvertTo-Json
$Headers = $AuthDB[$AppReg_LOG_READER].AuthHeaders
$CreatedQuery = Invoke-RestMethod -Method "POST" -Uri $Uri -Body $SearchParameters -Headers $Headers -ContentType $ContentTypeJSON
$QueryId = $CreatedQuery.Id
Write-Log "QueryId: $($QueryId)"
If (-not $QueryId){
    Write-Log "No QueryId - aborting script" -MessageType "ERR"
    . $IncFile_StdLogEndBlock
    Exit
}

##############################################################################
# load DB hashtables from files
$O365TeamGroup_DB = Import-CSVtoHashDB -Path $DBFileTeamsChannelsOwners -Keyname "FilesFolderUrl"
$AADGuest_DB = Import-CSVtoHashDB -Path $DBFileGuestsStd -KeyName "mail"

##############################################################################
# deleted users
$UriResource = "directory/deletedItems/microsoft.graph.user"
$UriSelect = "userPrincipalName,deletedDateTime,department,companyName,displayName,mail,onPremisesSamAccountName,onPremisesUserPrincipalName"
$UriFilter = "deletedDateTime+le+$twoDaysAgo"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect -Filter $UriFilter
$deletedUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON

foreach ($user in $deletedUsers) {
    if ($user.userPrincipalName -and -not($user.userPrincipalName.EndsWith($GuestUPNSuffix,$CCIgnoreCase))) {
        $OriginalUPN = $user.userPrincipalName.Substring($DelUPNPrefixLength)
        $ODfBUPN = $OriginalUPN -replace '\.', '_' -replace '@', '_'
        $DaysSinceDeleted = (New-TimeSpan -Start $User.deletedDateTime -End $now).Days
        $deletedUserObject = [pscustomobject]@{
            deletedDateTime = $user.deletedDateTime;
            daysSinceDeleted = $DaysSinceDeleted;
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

##############################################################################
# check audit log search status
$UriResource = "security/auditLog/queries/$($QueryId)"
$URI = New-GraphUri -Resource $UriResource -Version "beta"
Write-Host "Checking AuditLog search status" -NoNewline
Do {
    Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
    $Headers = $AuthDB[$AppReg_LOG_READER].AuthHeaders
    $AuditLogQuery = Invoke-RestMethod -Method "GET" -Uri $Uri -Headers $Headers -ContentType $ContentTypeJSON
    $Status = $AuditLogQuery.Status
    $DisplayName = $AuditLogQuery.DisplayName
    Switch ($Status) {
        "notStarted" {
            Write-Host "." -ForegroundColor DarkGray -NoNewLine
            Start-Sleep -Seconds 60
        }
        "running" {
            Write-Host "." -ForegroundColor Green -NoNewLine
            Start-Sleep -Seconds 60
        }
        "succeeded" {
            Write-Host "succeeded" -ForegroundColor Green
        }
        "failed" {
            Write-Host "." -ForegroundColor Red -NoNewline
            Write-Log "AuditLog search failed - exiting" -MessageType "ERR"
            . $IncFile_StdLogEndBlock	
            Exit
        }
    }
} Until ($Status -eq "succeeded")

##############################################################################
# get audit log search results
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

$Uri = "https://graph.microsoft.com/beta/security/auditLog/queries/$($SearchId)/records?`$Top=999"
$Headers = $AuthDB[$AppReg_LOG_READER].AuthHeaders
Write-Host $Uri
Write-Host "Reading records ($($DisplayName))" -NoNewline
Do {
    $Retry = $false
    Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30 -Silent
    $Query = Invoke-WebRequest -Headers $Headers -Uri $Uri -Method "GET"
    $QueryRecord = $Query | ConvertFrom-Json
    $AuditRecords += $QueryRecord.value    
    Write-Host "." -NoNewline
    if ($query.Headers.'Retry-After') {
        $Retry = $true
        Start-Sleep -Seconds $query.Headers.'Retry-After'
    }
    if (-not $Retry) {
        $Uri = $QueryRecord.'@odata.NextLink'   
    }
} Until (-not $Uri)

##############################################################################
# process audit log search results
Write-Log "Processing $($AuditRecords.Count) records"
foreach ($result in $AuditRecords) {
    try {
        $upn = $result.userPrincipalName.ToLower()
    }
    catch {
        Write-Host $result -ForegroundColor Red
        Write-host $result.AuditData -ForegroundColor Red
    }
    $SiteURL = $null
    $auditData = $result.AuditData
    if ($auditData.SiteUrl) {
        $SiteURL = ($auditData.SiteUrl).Trim().ToLower().TrimEnd("/")
        
        #deleted users
        if (($SiteUrl.StartsWith($RootODfBURL)) -and (-not($auditData.UserId.Contains("app@sharepoint")))) {
            $UrlUpn = $SiteURL.Substring($RootODfBURL.Length + 1)
            if ($UrlUpn.IndexOf("/") -gt 0) {
                $UrlUpn = $UrlUpn.Substring(0,$UrlUpn.IndexOf("/"))
            }
            if ($deletedUsers_DB.ContainsKey($UrlUpn)) {
                $deletedUser = $deletedUsers_DB[$UrlUpn]
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
                    DelUsr_DaysSinceDel = $deletedUser.daysSinceDeleted;
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
        
        #B2B guest accounts - UPN ends with #ext#@xyz.onmicrosoft.com
        if ($upn -like "*$($GuestUPNSuffix)") {
            $totalGuestsALL++
            $mail = Get-MailFromGuestUPN -GuestUPN $auditData.UserId
            Try{
                $sensitivityLabelName = $AIPLabelDB.Item($auditData.SensitivityLabelId)
            } Catch {
                $sensitivityLabelName = [string]::Empty
            }
    
            Try {
                $currentTeam = $O365TeamGroup_DB.Item($SiteURL)
            } Catch {
                $currentTeam = $null
            }
            
            Try {
                $GuestData = $AADGuest_DB.Item($mail)
            } Catch {
                $guestData = $null
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
                guestId                 = $guestData.Id;
                displayName             = $guestData.displayName;
                createdDateTime         = $guestData.createdDateTime;
                createdBy               = $guestData.EmployeeType;
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
                TeamName                = $currentTeam.teamName;
                ChannelName             = $currentTeam.channelName;
                TeamId                  = $currentTeam.TeamId;
                Owners                  = $currentTeam.Owners;
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
                }
            }
        }
    }
}#foreach ($result in $results)

Write-Log "totalGuestsALL: $($totalGuestsALL)"
Write-Log "totalLabeledAll: $($totalLabeledAll)"
Write-Log "totalLabeledSNE: $($totalLabeledSNE)"
Write-Log "totalAccessedDEL: $($totalAccessedDEL)"
Write-Log "END. Total count: " -NoNewline
Write-Log "$($totalCount)" -ForegroundColor Yellow

Export-Report "all guest SPO file access" -Report $ReportSPOAuditLogAll -Path $OutputFileAll
Export-Report "guests SNE file access" -Report $ReportSPOAuditLogSNE -Path $OutputFileSNE
Export-Report "deleted ODfB file access" -Report $ReportSPOAuditLogDEL -Path $OutputFileDEL
Export-Report "all auditRecords" -Report $AuditRecords -Path $OutputFileDbg

#######################################################################################################################
<#
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
#>


#######################################################################################################################

. $IncFile_StdLogEndBlock
