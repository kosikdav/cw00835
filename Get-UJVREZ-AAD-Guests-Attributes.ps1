
# Set-AAD-Guests-Attributes
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1
. $ScriptPath\include-Script-StdIncBlock.ps1


$OutputFolder 		= "t2t-ujvrez\mailusers"
$OutputFilePrefix	= "mailusers"

$OutputFile = New-OutputFile -RootFolder $REF -Folder $OutputFolder -Prefix $OutputFilePrefix -Ext "csv"

$timeDiffTolerance  = 60
$sleepShort = 60
$sleepLong = 120

[array]$MailUserReport = @()


#######################################################################################################################
$AADEXTTenant_DB = Import-CSVtoHashDB -Path $DBFileExtAADTenants -KeyName "domain"
##############################################################################################
# read Guests from Graph 
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriFilter = "UserType eq 'Guest'"
$UriSelect = "id,userPrincipalName,userType,displayName,mail,onPremisesExtensionAttributes"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Filter $UriFilter -Select $UriSelect
[array]$AllAADGuests = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ProgressDots -Text "AAD guest users"

##############################################################################################
# Process all guest accounts

foreach ($Guest in $AllAADGuests) {
    Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 60
    $MailDomain = $ExtTenant = $PartnerTenant = $AADExtCompanyName = $CurrentCompanyName = $CurrentEmployeeType = $MailUser = $null
    $InboundSync = $XTSync = $false
    $emailAddresses = [string]::Empty
    $upn = $Guest.UserPrincipalName
    $ext15 = $Guest.onPremisesExtensionAttributes.extensionAttribute15
    $UriResource = "users/$($Guest.id)"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
    #check for cross tenant sync accounts - ext15 attr = XTSync_tenantId
    if ($ext15 -and $ext15.StartsWith("XTSync_")) {
        $XTSync = $true
        $XTSyncCounter++
    }
    if ($Guest.Mail) {
        $MailDomain	= ($Guest.Mail.Split("@")[1]).ToLower()
        $ExtTenant = $AADExtTenant_DB[$MailDomain]
        if ($CustomDomainsUJVREZ -contains $MailDomain) {
            write-host "$($Guest.mail) - custom domain match for UJVREZ: $($MailDomain)"
            $ExtTenant = $AADExtTenant_DB["ujv.cz"]
        }
    }
    
    if ($XTSync) {
        #guest is synced via T2T sync - extensionAttribute15 is set
        $tenantId = ($ext15 -Split "_", 2)[1]
        #fix possible _ to - in tenantId - ELENG :)
        $tenantId = $tenantId -replace "_","-"
        if ($ExtTenant -and ($ExtTenant.tenantId -ne $tenantId)) {
            write-log "$($Guest.mail) - tenantId mismatch $($ExtTenant.tenantId) vs $($tenantId)" -ForegroundColor "Red"
            Continue
        }
        if ($tenantId -eq $TenantId_UJVREZ) {
            $MailUser = Get-MailUser -Identity $Guest.mail -ErrorAction SilentlyContinue
            if ($MailUser.EmailAddresses) {
                $emailAddresses = $MailUser.EmailAddresses -join ";"
            }
            $ReportObject = [PSCustomObject]@{
                ExternalEmailAddress = $MailUser.ExternalEmailAddress
                UserPrincipalName = $Guest.UserPrincipalName
                AccountDisabled = $MailUser.AccountDisabled
                OtherMail = $MailUser.OtherMail
                IsDirSynced = $MailUser.IsDirSynced
                Alias = $MailUser.Alias
                CustomAttribute15 = $MailUser.CustomAttribute15
                DisplayName = $MailUser.DisplayName
                EmailAddresses = $emailAddresses
                ExternalDirectoryObjectId = $MailUser.ExternalDirectoryObjectId
                LegacyExchangeDN = $MailUser.LegacyExchangeDN
                PrimarySmtpAddress = $MailUser.PrimarySmtpAddress
                RecipientType = $MailUser.RecipientType
                RecipientTypeDetails = $MailUser.RecipientTypeDetails
                Identity = $MailUser.Identity
                Id = $MailUser.Id
                ExchangeObjectId = $MailUser.ExchangeObjectId
                Guid = $MailUser.Guid
            }
            $MailUserReport += $ReportObject
        }
    }
}

Export-Report -Report $MailUserReport -Path $OutputFile -SortProperty "ExternalEmailAddress"

