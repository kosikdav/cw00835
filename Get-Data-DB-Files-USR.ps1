#######################################################################################################################
# Get-Data-DB-Files
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    [string]$workloads
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "db"
$LogFilePrefix		= "get-data-db-files-usr"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

[array]$DBReportGuestsStd       = $null
[array]$DBReportGuestsName      = $null
[array]$DBReportGuestsSIA       = $null

[array]$DBReportUsersAllStd     = $null
[array]$DBReportUsersAllMin     = $null
[array]$DBReportUsersAllName    = $null
[array]$DBReportUsersAllSIA     = $null

[array]$DBReportUsersMemStd     = $null
[array]$DBReportUsersMemMin     = $null
[array]$DBReportUsersMemName    = $null
[array]$DBReportUsersMemSIA     = $null
[array]$DBReportUsersMemLic     = $null

[int]$ThrottlingDelayPerUserInMsec = 200

[hashtable]$ADSyncedUsers_DB =@{}

if (-not $interactiveRun) {
    $ADCredential = Import-Clixml -Path $aad_grp_mgmt_cred
}

#######################################################################################################################
. $IncFile_StdLogStartBlock

$AADTenantDomain_DB = Import-CSVToHashDB -Path $DBFileExtAADTenants -KeyName "domain"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$ADFilter = "msExchExtensionAttribute29 -like `"*`""
$ADProperties = @(
    "userPrincipalName",
    "DisplayName",
    "msExchExtensionAttribute40"
)
if ($interactiveRun) {
    $SyncedADUsers = Get-ADUser -Filter $ADFilter -Properties $ADProperties 
}
else {
    $SyncedADUsers = Get-ADUser -Credential $ADCredential -Filter $ADFilter -Properties $ADProperties
}
Write-Log "SyncedADUsers: $($SyncedADUsers.Count)"

foreach ($user in $SyncedADUsers) {
    $userObject = [pscustomobject]@{
        UserPrincipalName = $user.UserPrincipalName;
        DisplayName = $user.DisplayName;
        msExchExtensionAttribute40 = $user.msExchExtensionAttribute40
    }
    $ADSyncedUsers_DB.Add($user.UserPrincipalName, $userObject)
}
Write-Log "ADSyncedUsers_DB: $($ADSyncedUsers_DB.Count)"

$UriResource = "users"
$UriSelect1 = "id,userPrincipalName,displayName,userType,employeeType,accountEnabled,mail,companyName,department,jobTitle,mobilePhone,preferredLanguage"
$UriSelect2 = "createdDateTime,creationType,externalUserState,externalUserStateChangeDateTime"
$UriSelect3 = "onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesSamAccountName,onPremisesDistinguishedName,onPremisesImmutableId,signInActivity"
$UriSelect = $UriSelect1,$UriSelect2,$UriSelect3 -join ","
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 99 -Select $UriSelect
[array]$AllAADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "AAD users" -ProgressDots
foreach ($User in $AllAADUsers) {
    Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
    write-host "$($User.UserPrincipalName.PadRight(40)) " -ForegroundColor Yellow -NoNewline
    $msExchExtensionAttribute40 = $onpremDisplayName = $lsi = $lsiNI = $Tenant	= $null
    $UserRecordLic = $null
    
    $LastSignInDateTime = $LastSignInDateTime_NI = "never"
    $Mail = $MailDomain = $mobilePhone = $UPNExtMail = $UPNExtMailDomain = [string]::Empty
    $onPremisesImmutableId = $onPremisesGUID = $CreatedDate = [string]::Empty
    $AADPremLicense = $CopilotLicense = $EXOLicense = $SPOLicense = $TMSLicense = $IntuneLicense = $PwrAutLicense = $PwrAppLicense = [string]::Empty
    $M365E3SKU = $M365E3UATSKU = $M365E5SKU = $M365F3SKU = $M365CopilotSKU = $M365E5SecSKU = $TeamsPremSKU = $PwrAutPremSKU = $PwrAppPremSKU = [string]::Empty
    $AADPremLicenseNeeded = $false
    $DaysSinceCreated = -1
    $UPN = $User.UserPrincipalName.ToLower()
    if ($ADSyncedUsers_DB.ContainsKey($UPN)) {
        $msExchExtensionAttribute40 = $ADSyncedUsers_DB[$UPN].msExchExtensionAttribute40
        $onpremDisplayName = $ADSyncedUsers_DB[$UPN].DisplayName
    }

    if ($user.Mail) {
        $Mail = $User.Mail.ToLower()
        $MailDomain	= Get-DomainFromAddress -Address $Mail
    }
    
    if ($User.mobilePhone) {
        $mobilePhone = "Tel:" + $User.mobilePhone.Trim()
    }
    
    if ($User.onPremisesImmutableId) {
        $onPremisesImmutableId = $User.onPremisesImmutableId.Trim()
        $onPremisesGUID = ([Guid]([Convert]::FromBase64String($onPremisesImmutableId))).Guid
    }
    
    if ($User.UserType -eq "Member") {
        $UriResource = "users/$($User.id)/licenseDetails"
        $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
        try {
            $Query = Invoke-RestMethod -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON
            $licensedSKUs = $Query.Value
        }
        Catch {
            $licensedSKUs = $null
        }
        if ($licensedSKUs.count -gt 0) {
            foreach ($sku in $licensedSKUs) {
                if ($M365E3SKUIds.Contains($sku.skuId))         {$M365E3SKU = $true}
                if ($M365E3UATSKUIds.Contains($sku.skuId))      {$M365E3UATSKU = $true}
                if ($M365E5SKUIds.Contains($sku.skuId))         {$M365E5SKU = $true}
                if ($M365F3SKUIds.Contains($sku.skuId))         {$M365F3SKU = $true}
                if ($M365CopilotSKUIds.Contains($sku.skuId))    {$M365CopilotSKU = $true}
                if ($M365E5SecSKUIds.Contains($sku.skuId))      {$M365E5SecSKU = $true}
                if ($TeamsPremSKUIds.Contains($sku.skuId))      {$TeamsPremSKU = $true}
                if ($PwrAutPremSKUIds.Contains($sku.skuId))     {$PwrAutPremSKU = $true}
                if ($PwrAppPremSKUIds.Contains($sku.skuId))     {$PwrAppPremSKU = $true}

                foreach ($plan in $sku.servicePlans) {
                    if (($AADP1LicensePlans.Contains($plan.servicePlanId)) -and ([string]::IsNullOrEmpty($AADPremLicense))) {$AADPremLicense = "P1"}
                    if ($AADP2LicensePlans.Contains($plan.servicePlanId))       {$AADPremLicense = "P2"}
                    if ($EXOLicensePlans.Contains($plan.servicePlanId))         {$EXOLicense = $true}
                    if ($SPOLicensePlans.Contains($plan.servicePlanId))         {$SPOLicense = $true}
                    if ($TMSLicensePlans.Contains($plan.servicePlanId))         {$TMSLicense = $true}
                    if ($IntuneLicensePlans.Contains($plan.servicePlanId))      {$IntuneLicense = $true}
                    if ($PwrAutLicensePlans.Contains($plan.servicePlanId))      {$PwrAutLicense = $true}
                    if ($PwrAppLicensePlans.Contains($plan.servicePlanId))      {$PwrAppLicense = $true}
                    if ($M365CopilotLicensePlans.Contains($plan.servicePlanId)) {$CopilotLicense = $true}
                } #plan
            } #sku
            if ([string]::IsNullOrEmpty($AADPremLicense)) {
                $AADPremLicenseNeeded = $true
            }
        }
        $UserRecordLic = [pscustomobject]@{
            Id					= $User.id;
            RecordCreated       = $CurrentDate;
            UserPrincipalName 	= $UPN;
            DisplayName			= $User.displayName;
            CompanyName			= $User.companyName;
            Department			= $User.department;
            SKUCount            = $licensedSKUs.Count;
            M365E3SKU           = $M365E3SKU;
            M365E3UATSKU        = $M365E3UATSKU;
            M365E5SKU           = $M365E5SKU;
            M365F3SKU           = $M365F3SKU;
            M365CopilotSKU      = $M365CopilotSKU;
            M365E5SecSKU        = $M365E5SecSKU;
            TeamsPremSKU        = $TeamsPremSKU;
            PwrAutPremSKU       = $PwrAutPremSKU;
            PwrAppPremSKU       = $PwrAppPremSKU;

            AADPremLicense      = $AADPremLicense;
            EXOLicense          = $EXOLicense;
            SPOLicense          = $SPOLicense;
            TMSLicense          = $TMSLicense;
            IntuneLicense       = $IntuneLicense;
            PwrAutLicense       = $PwrAutLicense;
            PwrAppLicense       = $PwrAppLicense;
            CopilotLicense      = $CopilotLicense;
            AADPremLicenseNeeded = $AADPremLicenseNeeded
        }
    }

    write-host 
    # created date
    $CreatedDate = (([DateTime]$User.CreatedDateTime).ToUniversalTime()).ToString("yyyy-MM-dd")
    $DaysSinceCreated = (New-TimeSpan -Start $createdDate -End $currentDate).Days

    # last signin
    $lsi = $User.SignInActivity.LastSignInDateTime
    $lsiNI = $User.SignInActivity.LastNonInteractiveSignInDateTime
    if (-not (($null -eq $lsi) -or ($lsi -eq "0001-01-01T00:00:00Z"))) {
        $LastSignInDateTime	= [DateTime]$lsi
        $DaysSinceLastSignIn = (New-TimeSpan -Start $LastSignInDateTime.ToString("yyyy-MM-dd") -End $CurrentDate).Days
    }
    if (-not(($null -eq $lsiNI) -or ($lsiNI -eq "0001-01-01T00:00:00Z"))) {
        $LastSignInDateTime_NI	= [DateTime]$lsiNI
        $DaysSinceLastSignIn_NI = (New-TimeSpan -Start $LastSignInDateTime_NI.ToString("yyyy-MM-dd") -End $CurrentDate).Days
    }
    $DaysSinceLastSignIn,$DaysSinceLastSignIn_NI = $DaysSinceCreated

    # userType = guest
    if ($user.userType -eq "guest") {
        if ($AADTenantDomain_DB.ContainsKey($MailDomain)) {
            $Tenant = $AADTenantDomain_DB.Item($MailDomain)
        }
        $UPNExtMail = Get-MailFromGuestUPN -GuestUPN $UPN
        $UPNExtMailDomain = Get-DomainFromGuestUPN -GuestUPN $UPN
    }

    # records
    $UserRecordName = [pscustomobject]@{
        Id						    = $User.id
        RecordCreated               = $CurrentDate
        UserPrincipalName 			= $UPN
        Mail 					    = $Mail
        DisplayName 				= $User.DisplayName
    }
    
    $UserRecordMin = [pscustomobject]@{
        Id  						= $User.id
        RecordCreated               = $CurrentDate
        accountEnabled				= $User.accountEnabled
        UserType 					= $User.userType
        UserPrincipalName 			= $UPN
        Mail                        = $Mail
        DisplayName 				= $User.displayName
        CompanyName					= $User.companyName
        Department					= $User.department
        onPremisesSyncEnabled       = $User.onPremisesSyncEnabled
    }
    
    $UserRecordSIA = [pscustomobject]@{
        Id  						= $User.id
        RecordCreated               = $CurrentDate
        accountEnabled				= $User.accountEnabled
        UserType 					= $User.userType
        UserPrincipalName 			= $UPN
        Mail                        = $Mail
        DisplayName 				= $User.displayName
        CreatedDateTime 			= $User.createdDateTime
        LastSignIn				    = $LastSignInDateTime
        LastSignIn_NI			    = $LastSignInDateTime_NI
    }

    $UserRecordStd = [pscustomobject]@{
        Id  						= $User.id
        RecordCreated               = $CurrentDate
        UserPrincipalName 			= $UPN
        UPNDomain					= Get-DomainFromAddress -Address $UPN
        DisplayName 				= $User.DisplayName
        UserType 					= $User.UserType
        AccountEnabled				= $User.AccountEnabled
        Mail 						= $Mail
        MailDomain 				    = $MailDomain
        UPNExtMail				    = $UPNExtMail
        UPNExtMailDomain		    = $UPNExtMailDomain
        CompanyName					= $User.companyName
        Department					= $User.department
        JobTitle					= $User.jobTitle
        EmployeeType				= $User.employeeType
        MobilePhone 				= $MobilePhone
        preferredLanguage			= $User.preferredLanguage
        CreatedDateTime 			= $User.createdDateTime
        DaysSinceCreated			= $DaysSinceCreated
        LastSignIn				    = $LastSignInDateTime
        LastSignIn_NI			    = $LastSignInDateTime_NI
        DaysSinceLastSignIn		    = $DaysSinceLastSignIn
        DaysSinceLastSignIn_NI 	    = $DaysSinceLastSignIn_NI
        onPremisesSyncEnabled 		= $User.onPremisesSyncEnabled
        onPremisesLastSyncDateTime 	= $User.onPremisesLastSyncDateTime
        onPremisesSamAccountName	= $User.onPremisesSamAccountName
        onPremisesDistinguishedName	= $User.onPremisesDistinguishedName
        onPremisesImmutableId		= $onPremisesImmutableId
        onPremisesGUID				= $onPremisesGUID
        msExchExtensionAttribute40	= $msExchExtensionAttribute40
        onpremDisplayName			= $onpremDisplayName
        ExternalUserState 		    = $User.ExternalUserState
        ExternalUserStateChangeDateTime = $User.ExternalUserStateChangeDateTime
        ExtAADTenantId			    = $Tenant.tenantId
        ExtAADDisplayName		    = $Tenant.displayName
        ExtAADdefaultDomain		    = $Tenant.defaultDomainName
    }

    # usertype member
    if ($User.UserType -eq "member") {
        $DBReportUsersMemName += $UserRecordName
        $DBReportUsersMemMin += $UserRecordMin
        $DBReportUsersMemStd += $UserRecordStd
        $DBReportUsersMemSIA += $UserRecordSIA
        $DBReportUsersMemLic += $UserRecordLic
    }

    # usertype guest
    else {
        $DBReportGuestsName += $UserRecordName
        $DBReportGuestsStd += $UserRecordStd
        $DBReportGuestsSIA += $UserRecordSIA
    }      

    # all users
    $DBReportUsersAllName += $UserRecordName
    $DBReportUsersAllMin += $UserRecordMin
    $DBReportUsersAllStd += $UserRecordStd
    $DBReportUsersAllSIA += $UserRecordSIA
    
    Start-Sleep -Milliseconds $ThrottlingDelayPerUserInMsec
}

Export-Report "member users (name only)" -Report $DBReportUsersMemName -Path $DBFileUsersMemName -SortProperty UserPrincipalName
Export-Report "member users (minimal)" -Report $DBReportUsersMemMin -Path $DBFileUsersMemMin -SortProperty UserPrincipalName
Export-Report "member users (standard)" -Report $DBReportUsersMemStd -Path $DBFileUsersMemStd -SortProperty UserPrincipalName
Export-Report "member users (SIA)" -Report $DBReportUsersMemSIA -Path $DBFileUsersMemSIA -SortProperty UserPrincipalName
Export-Report "member users (Lic)" -Report $DBReportUsersMemLic -Path $DBFileUsersMemLic -SortProperty UserPrincipalName

Export-Report "guest users (name only)" -Report $DBReportGuestsName -Path $DBFileGuestsName -SortProperty UserPrincipalName
Export-Report "guest users (standard)" -Report $DBReportGuestsStd -Path $DBFileGuestsStd -SortProperty UserPrincipalName
Export-Report "guest users (SIA)" -Report $DBReportGuestsSIA -Path $DBFileGuestsSIA -SortProperty UserPrincipalName

Export-Report "all users (name only)" -Report $DBReportUsersAllName -Path $DBFileUsersAllName -SortProperty UserPrincipalName
Export-Report "all users (minimal)" -Report $DBReportUsersAllMin -Path $DBFileUsersAllMin -SortProperty UserPrincipalName
Export-Report "all users (standard)" -Report $DBReportUsersAllStd -Path $DBFileUsersAllStd -SortProperty UserPrincipalName
Export-Report "all users (SIA)" -Report $DBReportUsersAllSIA -Path $DBFileUsersAllSIA -SortProperty UserPrincipalName

. $IncFile_StdLogEndBlock
