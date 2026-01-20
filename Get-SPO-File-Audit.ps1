#######################################################################################################################
# Get-SPO-File-Audit
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1

#######################################################################################################################

$LogFolder			    = "exports"
$LogFilePrefix		    = "spo-file-audit"

$LogFolderGAT		= "aad-guests-audit-trace"
$LogFilePrefixGAT	= "aad-guests-audit-trace"

$OutputFolder           = "spo\audit"
$OutputFilePrefix       = "spo-file-audit"

$OutputFileSuffixAccAll = "access-b2b-all"
$OutputFileSuffixAccSNE = "access-b2b-sensitive-noenc"
$OutputFileSuffixAccDEL = "access-deleted-odfb"
$OutputFileSuffixShrAll = "sharing-all"
$OutputFileSuffixShrB2B = "sharing-b2b"

#setting unified log query parameters
[array]$auditedOperationsAccess = @(
    "FileAccessed",
    "FileAccessedExtended",
    "FileDownloaded",
    "FileSyncDownloadedFull",
    "FilePreviewed"
)
[array]$auditedOperationsSharing = @(
    "CompanyLinkCreated",
    "SharingSet",
    "SecureLinkCreated",
    "AddedToSecureLink",
    "SharingInvitationCreated"
)
[array]$ignoredFileTypes = @("aspx","spcolor","sptheme","DS_Store","url","conf","options","icp","modules","dat","version","png","jpg","jpeg","gif","svg")
[hashtable]$guestAuditRecords_DB = @{}

. $ScriptPath\include-Script-StdIncBlock.ps1
. $IncFile_AIP_labels
. $IncFile_Functions_Audit

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$LogFileGAT = New-OutputFile -RootFolder $RLF -Folder $LogFolderGAT -Prefix $LogFilePrefixGAT -Ext "log"

[hashtable]$deletedUsers_DB = @{}
[hashtable]$AuditSPOblobs_DB = @{}

[array]$ReportSPOAuditLogAccAll = @()
[array]$ReportSPOAuditLogAccSNE = @()
[array]$ReportSPOAuditLogAccDEL = @()
[array]$ReportSPOAuditLogShrAll = @()
[array]$ReportSPOAuditLogShrB2B = @()

[array]$ProcessedBlobs = @()
[array]$ToBeDeletedBlobs = @()

$now = Get-Date
$UserDeletedDaysAgoLimitODfB = 3
$DB_changed = $false
$ProgressPreference = "SilentlyContinue"

##############################################################################
. $IncFile_StdLogStartBlock

# load DB Audit_SPO_blobs_DB from file or initialize empty
if (test-path $DBFileMGMTAPI_Audit_SPO) {
    Try {
        $AuditSPOblobs_DB = Import-Clixml -Path $DBFileMGMTAPI_Audit_SPO
        Write-Log "DB file $($DBFileMGMTAPI_Audit_SPO) imported successfully, $($AuditSPOblobs_DB.count) records found"
    } 
    Catch {
        Write-Log "Error importing $($DBFileMGMTAPI_Audit_SPO), creating empty DB" -MessageType "Error"
        [hashtable]$AuditSPOblobs_DB = @{}
        $DB_changed = $true
    }
}
else {
    Write-Log "DB file $($DBFileMGMTAPI_Audit_SPO) not found, creating empty DB" -MessageType "Error"
    [hashtable]$AuditSPOblobs_DB = @{}
    $DB_changed = $true
}

##############################################################################
# load DB hashtables from files
$O365TeamGroup_DB = Import-CSVtoHashDB -Path $DBFileTeamsChannelsOwners -Keyname "FilesFolderUrl"
$AADGuest_DB = Import-CSVtoHashDB -Path $DBFileGuestsStd -KeyName "userPrincipalName"

##############################################################################
# deleted users
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "directory/deletedItems/microsoft.graph.user"
$UriSelect = "userPrincipalName,deletedDateTime,department,companyName,displayName,mail,onPremisesSamAccountName,onPremisesUserPrincipalName"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
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
# get audit log search results
$Resource = "https://manage.office.com"
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30 -Resource $Resource -Authority "login.windows.net" -Force
$PageCount = 0
$Headers = $AuthDB[$AppReg_LOG_READER].AuthHeaders
$Uri = "https://manage.office.com/api/v1.0/$($tenantId)/activity/feed/subscriptions/content?contentType=Audit.SharePoint"
write-host "Reading available content blobs" -NoNewline
Do {
    write-host $Uri -ForegroundColor Cyan
    $Request = Invoke-WebRequest -Headers $Headers -Uri $Uri -Method "GET"
    Try{
        $QueryRecords = $Request | ConvertFrom-Json
        #write-host "." -NoNewline
        $PageCount++
    }
    Catch {
        Write-Host "Error converting JSON" -ForegroundColor Red
    } 
    $AvailableContentBlobs += $QueryRecords
    $Uri = $Request.Headers.NextPageUri
    Clear-Variable Request
    Clear-Variable QueryRecords
} Until (-not $Uri)
write-host "done ($($AvailableContentBlobs.Count))"
write-log "AvailableContentBlobs: $($AvailableContentBlobs.Count)"
##############################################################################
# download blobs and process audit records
$ProcessedBlobCount = 0
$IgnoredBlobCount = 0
foreach ($Blob in $AvailableContentBlobs){
    if ($AuditSPOblobs_DB.ContainsKey($Blob.contentId)) {
        Write-Host "skipping $($Blob.contentId)" -ForegroundColor Yellow -BackgroundColor DarkGray
        $IgnoredBlobCount++
        continue
    }
    $ProcessedBlobCount++
    Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30 -Resource $Resource -Authority "login.windows.net"
    $AuditRecords = @()
    Try {
        $Request = Invoke-WebRequest -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -Uri $Blob.ContentUri -Method "GET"
        Try{
            $AuditRecords = $Request | ConvertFrom-Json
            Clear-Variable Request
        }
        Catch {
            Write-Host "Error converting JSON" -ForegroundColor Red
            Continue
        }
    } 
    Catch {
        Write-Log "Error reading $($Blob.contentId) ($($ProcessedBlobCount))" -MessageType "Error" -ForegroundColor Red
        Continue
    }
    #Write-Host "AuditRecords: $($AuditRecords.count)" -ForegroundColor Yellow -BackgroundColor DarkGray
    
    foreach ($Record in $AuditRecords){
        $currentTeam = $null
        if ($Record.SiteUrl) {
            $SiteURL = ($Record.SiteUrl).Trim().ToLower().TrimEnd("/")
            Try{
                $sensitivityLabelName = $AIPLabelDB.Item($Record.SensitivityLabelId)
            } Catch {
                $sensitivityLabelName = [string]::Empty
            }
    
            Try {
                $currentTeam = $O365TeamGroup_DB.Item($SiteURL)
            } Catch {
                $currentTeam = $null
            }
        }
        else {
            Continue
        }
        if ($Record.UserId -like ("app@sharepoint")) {
            Continue
        }
        
        ##############################################################################
        # access
        if ($Record.Operation -in $auditedOperationsAccess) {
            #deleted users
            if ($SiteUrl.StartsWith($RootODfBURL)) {
                $UrlUpn = $SiteURL.Substring($RootODfBURL.Length + 1)
                if ($UrlUpn.IndexOf("/") -gt 0) {
                    $UrlUpn = $UrlUpn.Substring(0,$UrlUpn.IndexOf("/"))
                }
                if ($deletedUsers_DB.ContainsKey($UrlUpn)) {
                    $deletedUser = $deletedUsers_DB[$UrlUpn]
                    $UserDeletedDaysAgo = (New-TimeSpan -Start $deletedUser.deletedDateTime -End $Record.CreationTime).Days
                    if ($UserDeletedDaysAgo -ge $UserDeletedDaysAgoLimitODfB) {
                        $auditObjectDEL = [pscustomobject]@{
                            Id              = $Record.Id;
                            CorrelationId   = $Record.CorrelationId;
                            CreationTime    = $Record.CreationTime;
                            CreationDate    = $Record.CreationTime.Substring(0,10);
                            UserId          = $Record.UserId;
                            ClientIP        = $Record.ClientIP;
                            OperationType   = $Record.Operation;
                            URL             = $SiteURL;
                            File            = $Record.SourceFileName;
                            ObjectId        = $Record.ObjectId;
                            DelUsr_UPN          = $deletedUser.userPrincipalName;
                            DelUsr_DeletedDate  = $deletedUser.deletedDateTime;
                            DelUsr_DaysSinceDel = $UserDeletedDaysAgo;
                            DelUsr_DisplayName  = $deletedUser.DisplayName;
                            DelUsr_Company      = $deletedUser.Company;
                            DelUsr_Department   = $deletedUser.Department;
                            DelUsr_Mail         = $deletedUser.Mail;
                            DelUsr_KPJM         = $deletedUser.SamAccountName
                        }
                        $ReportSPOAuditLogAccDEL += $auditObjectDEL
                        write-host "$('{0:d4}' -f $ProcessedBlobCount)" -ForegroundColor Yellow -BackgroundColor Green -NoNewline
                        write-host " $($Record.CreationTime) $($Record.UserId) - file:$($Record.SourceFileName) owner:$($deletedUser.userPrincipalName) $($UserDeletedDaysAgo)" -ForegroundColor Yellow -BackgroundColor Red
                    }
                }
            }

            #B2B guest accounts - UPN ends with #ext#@xyz.onmicrosoft.com
            if ($Record.UserId -like "*$($GuestUPNSuffix)" -and ($Record.SourceFileExtension -notin $ignoredFileTypes)) {
                $guestData = $mail = $null
                $currentEventProperties = [PSCustomObject]@{
                    Operation   = $Record.Operation;
                    UserId      = $Record.UserId;
                    ObjectId    = $Record.ObjectId
                }
                if ($lastEventProperties -and (compare-object $currentEventProperties $lastEventProperties)) {
                    Write-Host "skipping $($Record.UserId) - $($Record.SourceFileName)" -ForegroundColor Yellow
                    continue
                }
                Try {
                    $GuestData = $AADGuest_DB.Item($Record.UserId)
                    $mail = $GuestData.mail
                } Catch {
                    $mail = Get-MailFromGuestUPN -GuestUPN $Record.UserId
                }

                write-host "$('{0:d4}' -f $ProcessedBlobCount)" -ForegroundColor Yellow -BackgroundColor Green -NoNewline
                write-host " $($Record.CreationTime) $($Record.Operation) - $($Mail) - $($Record.SourceFileName)" -ForegroundColor Yellow
                $auditObject = [PSCustomObject]@{
                    #blobId = $Blob.contentId
                    #blobCreated = $Blob.contentCreated
                    #blobExpiration = $Blob.contentExpiration
                    Id              = $Record.Id;
                    CorrelationId   = $Record.CorrelationId;
                    AADSessionId    = $Record.AppAccessContext.AADSessionId;
                    CreationTime    = $Record.CreationTime;
                    CreationDate    = $Record.CreationTime.Substring(0,10);
                    Operation       = $Record.Operation;
                    UserId          = $Record.UserId;
                    Mail            = $mail
                    #UserType               = Get-AuditUserTypeFromCode $Record.UserType;
                    guestId                 = $guestData.Id;
                    displayName             = $guestData.displayName;
                    createdDateTime         = $guestData.createdDateTime;
                    createdBy               = $guestData.EmployeeType;
                    mailDomain              = $guestData.MailDomain;
                    extAADTenantId		    = $guestData.ExtAADTenantId;
                    extAADDisplayName	    = $guestData.ExtAADDisplayName;
                    extAADdefaultDomain	    = $guestData.ExtAADdefaultDomain;
                    #RecordType = $Record.RecordType;
                    #Version = $Record.Version;
                    ClientIP = $Record.ClientIP;
                    GeoLocation = $Record.GeoLocation;
                    AuthenticationType = $Record.AuthenticationType;
                    BrowserName = $Record.BrowserName;
                    BrowserVersion = $Record.BrowserVersion;
                    IsManagedDevice = $Record.IsManagedDevice;
                    ItemType = $Record.ItemType;
                    ListId = $Record.ListId;
                    ListItemUniqueId = $Record.ListItemUniqueId;
                    Platform = $Record.Platform;
                    Site = $Record.Site;
                    UserAgent = $Record.UserAgent;
                    WebId = $Record.WebId;
                    DeviceDisplayName = $Record.DeviceDisplayName;
                    ListBaseType = $Record.ListBaseType;
                    ListServerTemplate = $Record.ListServerTemplate;
                    SensitivityLabelId  = $Record.SensitivityLabelId;
                    SensitivityLabelName = $sensitivityLabelName;
                    SiteSensitivityLabelId = $Record.SiteSensitivityLabelId;
                    SourceFileExtension = $Record.SourceFileExtension;
                    SensitivityLabelOwnerEmail = $Record.SensitivityLabelOwnerEmail;
                    SiteUrl             = $SiteUrl;
                    ObjectId = $Record.ObjectId;

                    TeamName            = $currentTeam.teamName;
                    ChannelName         = $currentTeam.channelName;
                    TeamId              = $currentTeam.TeamId;
                    Owners              = $currentTeam.Owners;

                    SourceRelativeUrl = $Record.SourceRelativeUrl;
                    SourceFileName = $Record.SourceFileName;
                    ApplicationDisplayName = $Record.ApplicationDisplayName
                }

                $ReportSPOAuditLogAccAll += $auditObject
                if (($Record.SensitivityLabelId) -and ($AIPLabelDBSNE.Contains($Record.SensitivityLabelId))) {
                    $ReportSPOAuditLogAccSNE += $auditObject
                    write-host "$($Record.CreationTime) $($Record.Operation) $($Mail) - $($sensitivityLabelName) - $($Record.SourceFileName)" -ForegroundColor Red
                }
                $lastEventProperties = $currentEventProperties
                Clear-Variable auditObject
            }
        }
        
        ##############################################################################
        # sharing
        if ($Record.Operation -in $auditedOperationsSharing) {
            $GuestData = $mail = $null
            if ($Record.TargetUserOrGroupName -like "*$($GuestUPNSuffix)") {
                Try {
                    $GuestData = $AADGuest_DB.Item($Record.TargetUserOrGroupName)
                    $mail = $GuestData.mail
                } Catch {
                    $mail = Get-MailFromGuestUPN -GuestUPN $Record.UserId
                }
            }
            write-host "$('{0:d4}' -f $ProcessedBlobCount)" -ForegroundColor Yellow -BackgroundColor Green -NoNewline
            write-host " $($Record.CreationTime) $($Record.Operation) - $($Record.UserId)=>$($Record.TargetUserOrGroupName) - $($Record.SourceFileName)" -ForegroundColor DarkYellow

            $auditObjectShare = [PSCustomObject]@{
                Id = $Record.Id
                CorrelationId = $Record.CorrelationId
                UniqueSharingId = $Record.UniqueSharingId
                CreationTime    = $Record.CreationTime
                CreationDate    = $Record.CreationTime.Substring(0,10)
                Operation       = $Record.Operation
                EventData       = $Record.EventData
                EventSource     = $Record.EventSource
                Workload        = $Record.Workload
                UserId          = $Record.UserId
                Mail            = $mail

                TargetUserOrGroupType = $Record.TargetUserOrGroupType
                TargetUserOrGroupName = $Record.TargetUserOrGroupName
                guestId                 = $guestData.Id
                displayName             = $guestData.displayName
                createdDateTime         = $guestData.createdDateTime
                createdBy               = $guestData.EmployeeType
                mailDomain              = $guestData.MailDomain
                extAADTenantId		    = $guestData.ExtAADTenantId
                extAADDisplayName	    = $guestData.ExtAADDisplayName
                extAADdefaultDomain	    = $guestData.ExtAADdefaultDomain

                AuthenticationType  = $Record.AuthenticationType
                ClientIP            = $Record.ClientIP
                GeoLocation         = $Record.GeoLocation
                DeviceDisplayName   = $Record.DeviceDisplayName
                IsManagedDevice     = $Record.IsManagedDevice
                Platform            = $Record.Platform

                TeamName            = $currentTeam.teamName;
                ChannelName         = $currentTeam.channelName;
                TeamId              = $currentTeam.TeamId;
                Owners              = $currentTeam.Owners;

                SensitivityLabelId = $Record.SensitivityLabelId
                SensitivityLabelName = $sensitivityLabelName
                Site = $Record.Site
                WebId = $Record.WebId
                SiteUrl = $Record.SiteUrl
                SourceFileExtension = $Record.SourceFileExtension
                SourceFileName = $Record.SourceFileName
                SourceRelativeUrl = $Record.SourceRelativeUrl
                ObjectId = $Record.ObjectId
                ItemType = $Record.ItemType
                ListId = $Record.ListId
                ListItemUniqueId = $Record.ListItemUniqueId

                ApplicationDisplayName = $Record.ApplicationDisplayName
                ApplicationId = $Record.ApplicationId
                UserAgent = $Record.UserAgent
                BrowserName = $Record.BrowserName
                BrowserVersion = $Record.BrowserVersion
            }
            $ReportSPOAuditLogShrAll += $auditObjectShare
            if ($Record.TargetUserOrGroupName -like "*$($GuestUPNSuffix)") {
                $ReportSPOAuditLogShrB2B += $auditObjectShare
            }
        } # sharing

        if ($Record.UserId -like "*$($GuestUPNSuffix)") {
            Update-GuestAuditRecordDB -Id $Record.UserId -DateTime $Record.CreationTime -hashtableDB $guestAuditRecords_DB
        }
        Clear-Variable auditObjectShare
    }
    
    $blobRecord = [PSCustomObject]@{
        contentId           = $Blob.contentId; 
        contentCreated      = $Blob.contentCreated; 
        contentExpiration   = $Blob.contentExpiration;
        processedDate       = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
    $ProcessedBlobs += $blobRecord
    Clear-Variable AuditRecords
    Clear-Variable BlobRecord
    <#
    if ($ProcessedBlobs.count -ge 500) {
        break
    }
    #>
}

Write-Log "Processed blobs: $($ProcessedBlobs.count)"
Write-Log "Ignored blobs: $($IgnoredBlobCount)"

write-log "ReportSPOAuditLogAccAll: $($ReportSPOAuditLogAccAll.Count)"
write-log "ReportSPOAuditLogAccSNE: $($ReportSPOAuditLogAccSNE.Count)"
write-log "ReportSPOAuditLogAccDEL: $($ReportSPOAuditLogAccDEL.Count)"
write-log "ReportSPOAuditLogShrAll: $($ReportSPOAuditLogShrAll.Count)"
write-log "ReportSPOAuditLogShrB2B: $($ReportSPOAuditLogShrB2B.Count)"

$DatesAcc = @($ReportSPOAuditLogAccAll.CreationDate | Sort-Object -Unique)
$DatesShr = @($ReportSPOAuditLogShrAll.CreationDate | Sort-Object -Unique)
$AuditLogEventDates = @($DatesAcc + $DatesShr | Sort-Object -Unique)
Write-Log "AuditLogEventDates: $($AuditLogEventDates.count)"

write-host "-----------------------------------------------------------------------------------------" -ForegroundColor Green
foreach ($Date in $AuditLogEventDates) {
    $CurrentReportAccAll = $CurrentReportAccSNE = $CurrentReportAccDEL = $CurrentReportShrAll = $CurrentReportShrB2B = $null
    write-host "---Processing date: $($Date) ----------------------------------------------------------" -ForegroundColor Green


    $CurrentReportAccAll = @($ReportSPOAuditLogAccAll | Where-Object { $_.CreationDate -eq $Date })
    $CurrentReportAccSNE = @($ReportSPOAuditLogAccSNE | Where-Object { $_.CreationDate -eq $Date })
    $CurrentReportAccDEL = @($ReportSPOAuditLogAccDEL | Where-Object { $_.CreationDate -eq $Date })
    $CurrentReportShrAll = @($ReportSPOAuditLogShrAll | Where-Object { $_.CreationDate -eq $Date })
    $CurrentReportShrB2B = @($ReportSPOAuditLogShrB2B | Where-Object { $_.CreationDate -eq $Date })

    $CurrentOutputFileAccAll = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixAccAll -SpecificDate $Date -Ext "csv"
    $CurrentOutputFileAccSNE = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixAccSNE -SpecificDate $Date -Ext "csv"
    $CurrentOutputFileAccDEL = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixAccDEL -SpecificDate $Date -Ext "csv"
    $CurrentOutputFileShrAll = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixShrAll -SpecificDate $Date -Ext "csv"
    $CurrentOutputFileShrB2B = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixShrB2B -SpecificDate $Date -Ext "csv"

    if ($CurrentReportAccAll.Count -gt 0) {
        Export-Report "$($Date) - file access (all B2B)" -Report $CurrentReportAccAll -Path $CurrentOutputFileAccAll -SortProperty "CreationTime" -Append $true
    }
    if ($CurrentReportAccSNE.Count -gt 0) {
        Export-Report "$($Date) - file access (B2B SNE)" -Report $CurrentReportAccSNE -Path $CurrentOutputFileAccSNE -SortProperty "CreationTime" -Append $true
    }
    if ($CurrentReportAccDEL.Count -gt 0) {
        Export-Report "$($Date) - deleted ODfB file access" -Report $CurrentReportAccDEL -Path $CurrentOutputFileAccDEL -SortProperty "CreationTime" -Append $true
    }
    if ($CurrentReportShrAll.Count -gt 0) {
        Export-Report "$($Date) - sharing operations (all)" -Report $CurrentReportShrAll -Path $CurrentOutputFileShrAll -SortProperty "CreationTime" -Append $true
    }
    if ($CurrentReportShrB2B.Count -gt 0) {
        Export-Report "$($Date) - sharing operations (B2B)" -Report $CurrentReportShrB2B -Path $CurrentOutputFileShrB2B -SortProperty "CreationTime" -Append $true
    }
    write-host "-----------------------------------------------------------------------------------------" -ForegroundColor Green
}

#add processed blobs to DB
foreach ($blobRecord in $ProcessedBlobs) {
    $AuditSPOblobs_DB.Add($blobRecord.contentId, $blobRecord)
    $DB_changed = $true
}

#find expired blobs in DB
foreach ($blobId in $AuditSPOblobs_DB.Keys) {
    $contentExpiration = [datetime]$AuditSPOblobs_DB[$blobId].contentExpiration
    if ($contentExpiration -lt $now) {
        $ToBeDeletedBlobs += $blobId
        #write-host "Expired blob: $($blobId) $($AuditSPOblobs_DB[$blobId].contentExpiration) " -ForegroundColor Red
    }
}
Write-Log "Expired blobs in DB: $($ToBeDeletedBlobs.Count)"

#delete expired blobs from DB
foreach ($blobId in $ToBeDeletedBlobs) {
    $AuditSPOblobs_DB.Remove($blobId)
    $DB_changed = $true
}

#saving DB XML if needed
if (($AuditSPOblobs_DB.count -gt 0) -and ($DB_changed)){
    Try {
        $AuditSPOblobs_DB | Export-Clixml -Path $DBFileMGMTAPI_Audit_SPO
        Write-Log "DB file $($DBFileMGMTAPI_Audit_SPO) exported successfully, $($AuditSPOblobs_DB.count) records saved"
    }
    Catch {
        Write-Log "Error exporting $($DBFileMGMTAPI_Audit_SPO)" -MessageType "Error"
    }
}

#write guest audit records to Entra attribute "employeeHireDate"
Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
Write-GuestAuditRecordDBToEntra -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -hashtableDB $guestAuditRecords_DB -EntraAttribute "employeeHireDate" -AuditType "SPOAudit" -LogFile $LogFileGAT

#######################################################################################################################

. $IncFile_StdLogEndBlock