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
$LogFilePrefix  = "exo-recipients"

$OutputFolder       = "exo\reports"
$OutputFilePrefixRcpt	= "exo"
$OutputFileSuffixRcpt	= "rcpt-list"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$OutputFileRcpt   = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefixRcpt -Suffix $OutputFileSuffixRcpt -Ext "csv"

[hashtable]$AADUsers_DB = @{}
[array]$RecipientReport = @()

#######################################################################################################################
. $IncFile_StdLogStartBlock
write-log "Output file: $($OutputFileRcpt)"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriFilter = "userType eq 'Member'"
$UriSelect1 = "id,companyName,department,userPrincipalName,onPremisesSamAccountName,onPremisesDistinguishedName"
$UriSelect2 = "onPremisesExtensionAttributes,extension_008a5d3f841f4052ac1283ff4782c560_employeeNumber"
$UriSelect = $UriSelect1,$UriSelect2 -join ","
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Filter $UriFilter -Select $UriSelect -Top 999
[array]$AADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "AAD users" -ProgressDots
foreach ($AADUser in $AADUsers) {
    $UserObject = [pscustomobject]@{
        id              = $AADUser.id
        companyName     = $AADUser.companyName
        department      = $AADUser.department
        employeeNumber  = $AADUser.extension_008a5d3f841f4052ac1283ff4782c560_employeeNumber
        ext10           = $AADUser.onPremisesExtensionAttributes.ExtensionAttribute10
        samAccountName  = $AADUser.onPremisesSamAccountName
        dn              = $AADUser.onPremisesDistinguishedName
        OU              = $AADUser.onPremisesDistinguishedName -replace '^CN=[^,]+,'
    }
    $AADUsers_DB.Add($AADUser.id, $UserObject)
}
Write-Log "AADUsers_DB: $($AADUsers_DB.Count)"
Remove-Variable AADUsers

#######################################################################################################################
Write-Interactive "Reading EXO recipients..." -NoNewline
[array]$EXOREcipients = Get-EXORecipient -ResultSize Unlimited -PropertySets All
Write-Interactive "done ($($EXOREcipients.Count))"

foreach ($Recipient in $EXOREcipients) {
    Write-Interactive $Recipient.PrimarySmtpAddress
    $proxyAddresses = $smtpAddresses = $x500Addresses = $sipAddresses = $SPOAddresses = [string]::Empty
    $ArchiveGuid = $SMTPDomain = $FullAccessPerms = $SendAsPerms = [string]::Empty
    $user = $stat = $qtaBytes = $quotaPctUsed =  $null

    if ($AADUsers_DB.ContainsKey($Recipient.ExternalDirectoryObjectId)) {
        $User = $AADUsers_DB.Item($Recipient.ExternalDirectoryObjectId)
    }
        if ($Recipient.PrimarySmtpAddress) {
        $SMTPDomain = $Recipient.PrimarySmtpAddress.Split("@")[1];
    }

    if ($Recipient.EmailAddresses) {
        foreach ($EmailAddress in $Recipient.EmailAddresses) {
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

        $proxyAddresses = $Recipient.EmailAddresses -join ";"
    }
    if ($Recipient.ArchiveGuid -and $Recipient.ArchiveGuid -ne "00000000-0000-0000-0000-000000000000") {
        $ArchiveGuid = $Recipient.ArchiveGuid
    }

    $recipientObject = [pscustomobject]@{
        Id                          = $Recipient.Id
        DisplayName                 = $Recipient.DisplayName
        RecipientType               = $Recipient.RecipientType
        RecipientTypeDetails        = $Recipient.RecipientTypeDetails
        UserPrincipalName           = $Recipient.UserPrincipalName
        PrimarySmtpAddress          = $Recipient.PrimarySmtpAddress
        SMTPdomain                  = $SMTPDomain
        smtpAddresses               = $smtpAddresses
        x500Addresses               = $x500Addresses
        sipAddresses                = $sipAddresses
        SPOAddresses                = $SPOAddresses
        userId                      = $User.id
        companyName                 = $User.companyName
        department                  = $User.department
        employeeNumber              = $User.employeeNumber
        ext10                       = $User.ext10
        dn                          = $User.dn
        OU                          = $User.OU
        SamAccountName              = $User.samAccountName
        EmailAddresses              = $Recipient.EmailAddresses -join ";"
        ExternalDirectoryObjectId   = $Recipient.ExternalDirectoryObjectId
        Identity                    = $Recipient.Identity
        Alias                       = $Recipient.Alias
        FirstName                   = $Recipient.FirstName
        LastName                    = $Recipient.LastName
        Name                        = $Recipient.Name
        ArchiveGuid                 = $Recipient.ArchiveGuid
        AuthenticationType          = $Recipient.AuthenticationType
        City                        = $Recipient.City
        Notes                       = $Recipient.Notes
        Company                     = $Recipient.Company
        CountryOrRegion             = $Recipient.CountryOrRegion
        PostalCode                  = $Recipient.PostalCode
        CustomAttribute1            = $Recipient.CustomAttribute1
        CustomAttribute2            = $Recipient.CustomAttribute2
        CustomAttribute3            = $Recipient.CustomAttribute3
        CustomAttribute4            = $Recipient.CustomAttribute4
        CustomAttribute5            = $Recipient.CustomAttribute5
        CustomAttribute6            = $Recipient.CustomAttribute6
        CustomAttribute7            = $Recipient.CustomAttribute7
        CustomAttribute8            = $Recipient.CustomAttribute8
        CustomAttribute9            = $Recipient.CustomAttribute9
        CustomAttribute10           = $Recipient.CustomAttribute10
        CustomAttribute11           = $Recipient.CustomAttribute11
        CustomAttribute12           = $Recipient.CustomAttribute12
        CustomAttribute13           = $Recipient.CustomAttribute13
        CustomAttribute14           = $Recipient.CustomAttribute14
        CustomAttribute15           = $Recipient.CustomAttribute15
        ExtensionCustomAttribute1   = $Recipient.ExtensionCustomAttribute1
        ExtensionCustomAttribute2   = $Recipient.ExtensionCustomAttribute2
        ExtensionCustomAttribute3   = $Recipient.ExtensionCustomAttribute3
        ExtensionCustomAttribute4   = $Recipient.ExtensionCustomAttribute4
        ExtensionCustomAttribute5   = $Recipient.ExtensionCustomAttribute5
        Database                    = $Recipient.Database
        ArchiveDatabase             = $Recipient.ArchiveDatabase
        DatabaseName                = $Recipient.DatabaseName
        ManagedFolderMailboxPolicy  = $Recipient.ManagedFolderMailboxPolicy
        ExpansionServer             = $Recipient.ExpansionServer
        ExternalEmailAddress        = $Recipient.ExternalEmailAddress
        HiddenFromAddressListsEnabled = $Recipient.HiddenFromAddressListsEnabled
        EmailAddressPolicyEnabled   = $Recipient.EmailAddressPolicyEnabled
        ResourceType                = $Recipient.ResourceType
        ManagedBy                   = $Recipient.ManagedBy
        Manager                     = $Recipient.Manager
        ActiveSyncMailboxPolicy     = $Recipient.ActiveSyncMailboxPolicy
        ActiveSyncMailboxPolicyIsDefaulted = $Recipient.ActiveSyncMailboxPolicyIsDefaulted
        Office                      = $Recipient.Office
        ObjectCategory              = $Recipient.ObjectCategory
        OrganizationalUnit          = $Recipient.OrganizationalUnit
        Phone                       = $Recipient.Phone
        PoliciesIncluded            = $Recipient.PoliciesIncluded
        PoliciesExcluded            = $Recipient.PoliciesExcluded
        ServerLegacyDN              = $Recipient.ServerLegacyDN
        ServerName                  = $Recipient.ServerName
        StateOrProvince             = $Recipient.StateOrProvince
        StorageGroupName            = $Recipient.StorageGroupName
        Title                       = $Recipient.Title
        UMMailboxPolicy             = $Recipient.UMMailboxPolicy
        UMRecipientDialPlanId       = $Recipient.UMRecipientDialPlanId
        WindowsLiveID               = $Recipient.WindowsLiveID
        HasActiveSyncDevicePartnership = $Recipient.HasActiveSyncDevicePartnership
        AddressListMembership       = $Recipient.AddressListMembership
        OwaMailboxPolicy            = $Recipient.OwaMailboxPolicy
        AddressBookPolicy           = $Recipient.AddressBookPolicy
        SharingPolicy               = $Recipient.SharingPolicy
        RetentionPolicy             = $Recipient.RetentionPolicy
        ShouldUseDefaultRetentionPolicy = $Recipient.ShouldUseDefaultRetentionPolicy
        MailboxMoveTargetMDB        = $Recipient.MailboxMoveTargetMDB
        MailboxMoveSourceMDB        = $Recipient.MailboxMoveSourceMDB
        MailboxMoveFlags            = $Recipient.MailboxMoveFlags
        MailboxMoveRemoteHostName   = $Recipient.MailboxMoveRemoteHostName
        MailboxMoveBatchName        = $Recipient.MailboxMoveBatchName
        MailboxMoveStatus           = $Recipient.MailboxMoveStatus
        MailboxRelease              = $Recipient.MailboxRelease
        ArchiveRelease              = $Recipient.ArchiveRelease
        IsValidSecurityPrincipal    = $Recipient.IsValidSecurityPrincipal
        LitigationHoldEnabled       = $Recipient.LitigationHoldEnabled
        Capabilities                = $Recipient.Capabilities
        ArchiveState                = $Recipient.ArchiveState
        SKUAssigned                 = $Recipient.SKUAssigned
        WhenMailboxCreated          = $Recipient.WhenMailboxCreated
        UsageLocation               = $Recipient.UsageLocation
        ExchangeGuid                = $Recipient.ExchangeGuid
        ArchiveStatus               = $Recipient.ArchiveStatus
        SafeSendersHash             = $Recipient.SafeSendersHash
        SafeRecipientsHash          = $Recipient.SafeRecipientsHash
        BlockedSendersHash          = $Recipient.BlockedSendersHash
        WhenSoftDeleted             = $Recipient.WhenSoftDeleted
        ExchangeVersion             = $Recipient.ExchangeVersion
        DistinguishedName           = $Recipient.DistinguishedName
        ObjectClass                 = $Recipient.ObjectClass
        WhenChanged                 = $Recipient.WhenChanged
        WhenCreated                 = $Recipient.WhenCreated
        WhenChangedUTC              = $Recipient.WhenChangedUTC
        WhenCreatedUTC              = $Recipient.WhenCreatedUTC
        ExchangeObjectId            = $Recipient.ExchangeObjectId
        OrganizationId              = $Recipient.OrganizationId
    }
    $RecipientReport += $RecipientObject
}

#######################################################################################################################

Export-Report "EXO recipient report" -Report $RecipientReport -Path $OutputFileRcpt -SortProperty "UserPrincipalName"

#######################################################################################################################
. $IncFile_StdLogEndBlock
