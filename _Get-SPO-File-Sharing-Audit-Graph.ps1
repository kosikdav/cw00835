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
$LogFilePrefix		    = "spo-file-sharing-audit-graph"

$OutputFolder           = "spo\audit"
$OutputFilePrefix       = "file-access-graph"
$OutputFileSuffix       = "sharing"

$daysBackOffset = 0
If (-not [Environment]::UserInteractive) {
    $daysBackOffset = 0
}
$fileDateDayOffset = -1 - $daysBackOffset

#setting date variables
[DateTime]$Start = ((Get-Date -Hour 0 -Minute 0 -Second 0)).AddDays(-1 -$daysBackOffset)
[DateTime]$End = $Start.AddDays(1)
$StartUTC = $Start.ToString("yyyy-MM-ddTHH:mm:ssZ")
$EndUTC = $End.ToString("yyyy-MM-ddTHH:mm:ssZ")

#setting unified log query parameters
$recordTypeFilters = @("SharePointSharingOperation")
$operationFilters = @("CompanyLinkCreated","SharingSet","SecureLinkCreated","AddedToSecureLink","SharingInvitationCreated")

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1
. $IncFile_AIP_labels
. $IncFile_Functions_Audit

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputFileAll = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffix -FileDateDayOffset $fileDateDayOffset -Ext "csv"
$OutputFileDbg = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix "dbg-sharing" -FileDateDayOffset $fileDateDayOffset -Ext "csv"

[array]$SPOSharingAuditLogAll = @()
[int]$totalCount = 0
[int]$totalGuestsALL = 0

#######################################################################################################################

. $IncFile_StdLogStartBlock

Write-Log "Manual auth: $($ManualAuth)"
Write-Log "recordTypeFilters: $($recordTypeFilters)"
Write-Log "operationFilters: $($operationFilters)"
Write-Log "DaysBackOffset: $($daysBackOffset)"
Write-Log "Query start: $($start) ($($start.DayOfWeek))" -ForegroundColor Yellow
Write-Log "Query end:   $($end)" -ForegroundColor Yellow
Write-Log "Output file ALL: $($OutputFileAll)"
Write-Log "Output file DBG: $($OutputFileDbg)" -ForegroundColor Cyan
Write-Log "Guest UPN suffix: $($guestUPNSuffix)"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

##############################################################################
# create audit log search
$DisplayName = "SPOAuditLogSearch_" + (Get-Date -Format "yyyyMMdd_HHmm")
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
# check audit log search status
$UriResource = "security/auditLog/queries/$($QueryId)"
$URI = New-GraphUri -Resource $UriResource -Version "beta"
Do {
    Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
    $Headers = $AuthDB[$AppReg_LOG_READER].AuthHeaders
    $AuditLogQuery = Invoke-RestMethod -Method "GET" -Uri $Uri -Headers $Headers -ContentType $ContentTypeJSON
    $Status = $AuditLogQuery.Status
    Switch ($Status) {
        "notStarted" {
            Write-Host "." -ForegroundColor DarkGray -nonewline
            Start-Sleep -Seconds 60
        }
        "running" {
            Write-Host "." -ForegroundColor Green -nonewline
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
$UriResource = "security/auditLog/queries/$($QueryId)/records"
$Uri = New-GraphUri -Resource $UriResource -Version "beta"
$AccessToken = $AuthDB[$AppReg_LOG_READER].AccessToken
$AuditRecords = Get-GraphOutputREST -Uri $Uri -AccessToken $AccessToken -ContentType $ContentTypeJSON -ProgressDots -Text "AuditRecords"
Write-Log "Audit records found: $($AuditRecords.Count)" -ForegroundColor Cyan

##############################################################################
# load DB hashtables from files
$O365TeamGroup_DB = Import-CSVtoHashDB -Path $DBFileTeamsChannelsOwners -Keyname "FilesFolderUrl"
$AADGuest_DB = Import-CSVtoHashDB -Path $DBFileGuests -KeyName "UserPrincipalName"

##############################################################################
# process audit log search results
foreach ($result in $AuditRecords) {
    $guestData = $null
    $SiteURL = $targetUserMail = [string]::Empty
    $auditData = $result.AuditData
    if ($auditData.SiteUrl) {
        $SiteURL = ($auditData.SiteUrl).Trim().ToLower().TrimEnd("/")
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
        if ($auditData.TargetUserOrGroupName) {
            if ($auditData.TargetUserOrGroupName.Contains("@")) {
                $TargetUserMail = $auditData.TargetUserOrGroupName.Trim().ToLower()
            }
            if ($AADGuest_DB.ContainsKey($auditData.TargetUserOrGroupName)) {
                $guestData = $AADGuest_DB[$auditData.TargetUserOrGroupName]
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
    }   
}#foreach ($result in $results)


Write-Log "totalGuestsALL: $($totalGuestsALL)"
Write-Log "$($totalCount)" -ForegroundColor Yellow
Write-Log "Index errors: $($indexErrorCountTotal)"

Export-Report -Text "SPOSharingAuditLogAll report" -Report $SPOSharingAuditLogAll -Path $OutputFileAll
Export-Report "all auditRecords" -Report $AuditRecords -Path $OutputFileDbg

#######################################################################################################################

. $IncFile_StdLogEndBlock
