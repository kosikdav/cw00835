#######################################################################################################################
# Get-EXO-Mailboxes-Reports
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    $IncludeMailboxPermissions = $false
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder		= "exports"
$LogFilePrefix  = "exo-reports"

$OutputFolder       = "exo\reports"
$OutputFilePrefix	= "exo"
$OutputFileSuffix	= "mbx-list"

$OutputFolderQF		= "exo-80pct-full"
$OutputFilePrefixQF	= "exo-mbx-quota-80-pct-full"

$OutputFolderTNR		= "tenaur\exo"
$OutputFilePrefixTNR	= "exo-mbx-list-tenaur"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputFile     = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffix -Ext "csv"
$OutputFileQF   = New-OutputFile -RootFolder $ROF -Folder $OutputFolderQF -Prefix $OutputFilePrefixQF -Ext "csv"
$OutputFileTNR  = New-OutputFile -RootFolder $ROF -Folder $OutputFolderTNR -Prefix $OutputFilePrefixTNR -Ext "csv"

[hashtable]$MailboxStats_DB = @{}
[hashtable]$AADUsers_DB = @{}
[array]$MailboxReport = @()
[array]$MailboxReportTNR = @()

if ($IncludeMailboxPermissions) {
    $EXOMbxReportPermissions = $true
}

#######################################################################################################################

. $ScriptPath\include-Script-StartLog-Generic.ps1

Write-Log "Output file - quota 80% full: $($OutputFileQF)"
if ($EXOMbxReportTNR) {
    Write-Log "Output file - TENAUR: $($OutputFileTNR)"
}

#$AADUsers_DB = Import-CSVtoHashDB -Path $DBFileUsersMemMin -KeyName "userPrincipalName"
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriFilter = "userType+eq+'Member'"
$UriSelect = "id,companyName,department,userPrincipalName"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Filter $UriFilter -Select $UriSelect -Top 999
[array]$AADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "AAD users" -ProgressDots
foreach ($AADUser in $AADUsers) {
    $UserObject = [pscustomobject]@{
        Id              = $AADUser.id;
        CompanyName     = $AADUser.companyName;
        Department      = $AADUser.department;
        UserPrincipalName = $AADUser.userPrincipalName;
    }
    $AADUsers_DB.Add($AADUser.userPrincipalName, $UserObject)
}
Write-Log "AADUsers_DB: $($AADUsers_DB.Count)"
Remove-Variable AADUsers

# get mailbox statistics #####################################
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "reports/getMailboxUsageDetail"
$UriReportPeriod = "D180"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -ReportPeriod $UriReportPeriod
[array]$MailboxStats = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeCSV
Write-Log "$($MailboxStats.Count) mailbox stat records returned by query"

foreach ($MailboxStat in $MailboxStats) {
    $isDel  = $MailboxStat."Is Deleted"
    $upn    = $MailboxStat."User Principal Name"
    $su     = $MailboxStat."Storage Used (Byte)"
    $du     = $MailboxStat."Deleted Item Size (Byte)"
    $MailboxStatsObject = [pscustomobject]@{
		ItemCount	        = $MailboxStat."Item Count"
        StorageUsed         = $su
        StorageUsedGB       = "$("{0:N2}" -f ($su / 1073741824)) GB ($("{0:N0}" -f $su) bytes)"
        DeletedItemCount    = $MailboxStat."Deleted Item Count"
        DeletedItemSize     = $du
        DeletedItemSizeGB   = "$("{0:N2}" -f ($du / 1073741824)) GB ($("{0:N0}" -f $du) bytes)"
    }
    if ($isDel -eq $false) {
        $MailboxStats_DB.Add($upn, $MailboxStatsObject)
    }
}
Remove-Variable MailboxStats

# retrieve mailboxes ##################################### 
Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 120
[array]$Mailboxes = Get-EXOMailbox -ResultSize Unlimited -PropertySets All
Write-Log "$($Mailboxes.Count) mailboxes returned by query"
Write-Log "EXOMbxReportPermissions: $($EXOMbxReportPermissions)"

ForEach ($Mailbox in $Mailboxes) {
    $proxyAddresses = $smtpAddresses = $x500Addresses = $sipAddresses = $SPOAddresses = [string]::Empty
    $ArchiveGuid = $SMTPDomain = $FullAccessPerms = $SendAsPerms = [string]::Empty
    $user = $stat = $qtaBytes = $quotaPctUsed =  $null

    if ($MailboxStats_DB.ContainsKey($mailbox.UserPrincipalName)) {
        $stat = $MailboxStats_DB.Item($mailbox.UserPrincipalName)
    }   
    if ($AADUsers_DB.ContainsKey($mailbox.UserPrincipalName)) {
        $User = $AADUsers_DB.Item($mailbox.UserPrincipalName)
    }

    $qta = $mailbox.ProhibitSendReceiveQuota
    $qtaBytes = ($qta.Substring($qta.IndexOf("(")+1,$qta.Length-$qta.IndexOf("(")-7)).Replace(",","")
    [decimal]$quotaPctUsed = "{0:N2}" -f [decimal](($stat.StorageUsed / $qtaBytes) * 100)
    
    if ($Mailbox.PrimarySmtpAddress) {
        $SMTPDomain = $mailbox.PrimarySmtpAddress.Split("@")[1];
    }

    if ($mailbox.EmailAddresses) {
        foreach ($EmailAddress in $mailbox.EmailAddresses) {
            if ($EmailAddress -like "smtp:*" -and $EmailAddress -notlike "*onmicrosoft.com") {
                $smtpAddresses += $EmailAddress.Split(":")[1].ToLower() + ";"
            }
            if ($EmailAddress -like "X500:*") {
                $x500Addresses += $EmailAddress.Split(":")[1] + ";"
            }
            if ($EmailAddress -like "SIP:*") {
                $sipAddresses += $EmailAddress.Split(":")[1].ToLower() + ";"
            }
            if ($EmailAddress -like "SPO:*") {
                $SPOAddresses += $EmailAddress.Split(":")[1].ToLower() + ";"
            }

        }
        $smtpAddresses = $smtpAddresses.TrimEnd(";")
        $x500Addresses = $x500Addresses.TrimEnd(";")
        $sipAddresses = $sipAddresses.TrimEnd(";")
        $SPOAddresses = $SPOAddresses.TrimEnd(";")

        $proxyAddresses = $mailbox.EmailAddresses -join ";"
    }
    if ($mailbox.ArchiveGuid -and $mailbox.ArchiveGuid -ne "00000000-0000-0000-0000-000000000000") {
        $ArchiveGuid = $mailbox.ArchiveGuid
    }

    if ($EXOMbxReportPermissions) {
        #Write-Host "# Get-EXOMailboxPermission"
        Try {
            $FullAccessPermissions = Get-EXOMailboxPermission -Identity $mailbox.UserPrincipalName | Where-Object {$_.User -Like "*@*" }
            foreach ($permission in $FullAccessPermissions) {
                if ($permission.AccessRights -contains("FullAccess")) {
                    $FullAccessPerms += $permission.User.ToLower() + ";"
                }
            }
            $FullAccessPerms = $FullAccessPerms.TrimEnd(";")
        }
        Catch {
            write-host "Get-EXOMailboxPermission failed for $($mailbox.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Red
        }

        #Write-Host "# Get-ExoRecipientPermission"
        Try {
            $SendAsPermissions = Get-ExoRecipientPermission -Identity $mailbox.UserPrincipalName | Where-Object {$_.Trustee -ne "NT AUTHORITY\SELF"}
            foreach ($permission in $SendAsPermissions) {
                if ($permission.AccessRights -contains("SendAs") -and $permission.Trustee.Contains("@")) {
                    $SendAsPerms += $permission.Trustee.ToLower() + ";"
                }
            }
            $SendAsPerms = $SendAsPerms.TrimEnd(";")
        }
        Catch {
            write-host "Get-ExoRecipientPermission failed for $($mailbox.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    $MailboxObject = [pscustomobject]@{
        Id						    = $mailbox.id
        UserPrincipalName		    = $mailbox.UserPrincipalName
        DisplayName                 = $mailbox.DisplayName
        UPNdomain                   = $mailbox.UserPrincipalName.Split("@")[1]
        PrimarySmtpAddress          = $mailbox.PrimarySmtpAddress
        SMTPdomain                  = $SMTPDomain
        smtpAddresses               = $smtpAddresses
        x500Addresses               = $x500Addresses
        sipAddresses                = $sipAddresses
        SPOAddresses                = $SPOAddresses
        CompanyName                 = $User.CompanyName
        Department                  = $User.Department
        IsDirSynced                 = $mailbox.IsDirSynced
        Alias                       = $mailbox.Alias

        ExchangeObjectId            = $mailbox.ExchangeObjectId
        ExchangeGuid                = $mailbox.ExchangeGuid
        DistinguishedName           = $mailbox.DistinguishedName
        ExternalDirectoryObjectId   = $mailbox.ExternalDirectoryObjectId
        Identity                    = $mailbox.Identity
        Name                        = $mailbox.Name
        LegacyExchangeDN            = $mailbox.LegacyExchangeDN
        WindowsLiveId               = $mailbox.WindowsLiveId
        NetId                       = $mailbox.NetId
        SamAccountName              = $mailbox.SamAccountName

        RecipientType               = $mailbox.RecipientType
        RecipientTypeDetails        = $mailbox.RecipientTypeDetails
        WhenCreated                 = $mailbox.WhenCreated
        WhenMailboxCreated          = $mailbox.WhenMailboxCreated
        WhenChanged                 = $mailbox.WhenChanged
        RequireSenderAuth           = $mailbox.RequireSenderAuthenticationEnabled
        AuditEnabled                = $mailbox.AuditEnabled
        HiddenFromAddressLists      = $mailbox.HiddenFromAddressListsEnabled
        GrantSendOnBehalfTo         = $mailbox.GrantSendOnBehalfTo
        FullAccessPermissions       = $FullAccessPerms
        SendAsPermissions           = $SendAsPerms
        
        <#
        CustomAttribute1        = $mailbox.CustomAttribute1
        CustomAttribute2        = $mailbox.CustomAttribute2
        CustomAttribute3        = $mailbox.CustomAttribute3
        CustomAttribute4        = $mailbox.CustomAttribute4
        CustomAttribute5        = $mailbox.CustomAttribute5
        CustomAttribute6        = $mailbox.CustomAttribute6
        CustomAttribute7        = $mailbox.CustomAttribute7
        CustomAttribute8        = $mailbox.CustomAttribute8
        CustomAttribute9        = $mailbox.CustomAttribute9
        CustomAttribute10       = $mailbox.CustomAttribute10
        CustomAttribute11       = $mailbox.CustomAttribute11
        CustomAttribute12       = $mailbox.CustomAttribute12
        CustomAttribute13       = $mailbox.CustomAttribute13
        CustomAttribute14       = $mailbox.CustomAttribute14
        CustomAttribute15       = $mailbox.CustomAttribute15    
        #>

        QuotaPercentUsed            = $quotaPctUsed
        StorageUsed                 = $stat.StorageUsedGB
        StorageUsedBytes            = $stat.StorageUsed
        ItemCount                   = $stat.ItemCount
        DeletedItemCount            = $stat.DeletedItemCount
        DeletedItemSize             = $stat.DeletedItemSizeGB

        IssueWarningQuota               = $mailbox.IssueWarningQuota.Replace(",","")
        ProhibitSendQuota               = $mailbox.ProhibitSendQuota.Replace(",","")
        ProhibitSendReceiveQuota        = $mailbox.ProhibitSendReceiveQuota.Replace(",","")
        RecoverableItemsQuota           = $mailbox.RecoverableItemsQuota.Replace(",","")
        RecoverableItemsWarningQuota    = $mailbox.RecoverableItemsWarningQuota.Replace(",","")
        RulesQuota                      = $mailbox.RulesQuota.Replace(",","")
        RecipientLimits                 = $mailbox.RecipientLimits
        UseDatabaseQuotaDefaults        = $mailbox.UseDatabaseQuotaDefaults



        ArchiveStatus               = $mailbox.ArchiveStatus
        ArchiveGuid                 = $ArchiveGuid
        ArchiveState                = $mailbox.ArchiveState
        ArchiveName                 = $mailbox.ArchiveName
        AutoExpandingArchiveEnabled = $mailbox.AutoExpandingArchiveEnabled
        SingleItemRecoveryEnabled   = $mailbox.SingleItemRecoveryEnabled

        AntispamBypassEnabled       = $mailbox.AntispamBypassEnabled

        ExternalOofOptions          = $mailbox.ExternalOofOptions
        DeliverToMailboxAndForward  = $mailbox.DeliverToMailboxAndForward
        ForwardingAddress           = $mailbox.ForwardingAddress
        ForwardingSmtpAddress       = $mailbox.ForwardingSmtpAddress

        DelayHoldApplied            = $mailbox.DelayHoldApplied
        DelayReleaseHoldApplied     = $mailbox.DelayReleaseHoldApplied
        LitigationHoldDate          = $mailbox.LitigationHoldDate
        LitigationHoldDuration      = $mailbox.LitigationHoldDuration
        LitigationHoldEnabled       = $mailbox.LitigationHoldEnabled
        LitigationHoldOwner         = $mailbox.LitigationHoldOwner
        
        MailboxPlan                 = $mailbox.MailboxPlan
        SKUAssigned                 = $mailbox.SKUAssigned
        
        MaxReceiveSize              = $mailbox.MaxReceiveSize
        MaxSendSize                 = $mailbox.MaxSendSize

        MessageCopyForSendOnBehalfEnabled       = $mailbox.MessageCopyForSendOnBehalfEnabled
        MessageCopyForSentAsEnabled             = $mailbox.MessageCopyForSentAsEnabled
        MessageCopyForSMTPCltSubmEnabled        = $mailbox.MessageCopyForSMTPClientSubmissionEnabled
        MessageRecallProcessingEnabled          = $mailbox.MessageRecallProcessingEnabled
        MessageTrackingReadStatusEnabled        = $mailbox.MessageTrackingReadStatusEnabled
        
        RecipientThrottlingThreshold    = $mailbox.RecipientThrottlingThreshold
        RetainDeletedItemsFor           = $mailbox.RetainDeletedItemsFor
        RetainDeletedItemsUntilBackup   = $mailbox.RetainDeletedItemsUntilBackup
        UseDatabaseRetentionDefaults    = $mailbox.UseDatabaseRetentionDefaults
        RetentionComment                = $mailbox.RetentionComment
        RetentionHoldEnabled            = $mailbox.RetentionHoldEnabled
        RetentionPolicy                 = $mailbox.RetentionPolicy
        SharingPolicy                   = $mailbox.SharingPolicy
        EmailAddressPolicyEnabled   = $mailbox.EmailAddressPolicyEnabled
        WasInactiveMailbox          = $mailbox.WasInactiveMailbox
        proxyAddresses              = $proxyAddresses

    }
    $MailboxReport += $MailboxObject
    
    if ($EXOMbxReportTNR) {
        if ($Mailbox.EmailAddresses -like "*tenaur.cz*") {
            $MailboxReportTNR += $MailboxObject
        }
    }

}

#######################################################################################################################

Export-Report "mailbox report" -Report $MailboxReport -Path $OutputFile -SortProperty "UserPrincipalName"
Export-Report "mailbox report - quota 80% full" -Report $MailboxReport.Where({$_.QuotaPercentUsed -ge 80}) -Path $OutputFileQF -SortProperty "UserPrincipalName"
if ($EXOMbxReportTNR) {
    Export-Report "mailbox report - TENAUR" -Report $MailboxReportTNR -Path $OutputFileTNR -SortProperty "UserPrincipalName"
}

. $ScriptPath\include-Script-EndLog-generic.ps1
