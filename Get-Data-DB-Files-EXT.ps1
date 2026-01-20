#######################################################################################################################
# Get-Data-DB-Files-EXT
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
$LogFilePrefix		= "get-data-db-files-ext"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

[array]$DBReportTenantDomains       = @()
[array]$DBReportAADPartnerTenants   = @()
[array]$DBReportExtAADTenants       = @()

[int]$ThrottlingDelayPerTenantInMsec = 500
[int]$ThrottlingDelayPerDomainInMsec = 500

$TTL = 30

#######################################################################################################################

. $IncFile_StdLogStartBlock

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
[array]$GuestDomainArray = @()
[hashtable]$AADPartnerTenants_DB = @{}

$UriResource = "domains"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
[array]$TenantDomains = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
ForEach ($Domain in $TenantDomains) {
    $Email = $Intune = $OrgIdAuthentication = $OfficeCommunicationsOnline = $false

    if ($Domain.supportedServices -contains "Email") {
        $Email = $true
    }
    if ($Domain.supportedServices -contains "Intune") {
        $Intune = $true
    }
    if ($Domain.supportedServices -contains "OrgIdAuthentication") {
        $OrgIdAuthentication = $true
    }
    if ($Domain.supportedServices -contains "OfficeCommunicationsOnline") {
        $OfficeCommunicationsOnline = $true
    }

    $DomainObject = [pscustomobject]@{
        domain              = $Domain.id
        isVerified          = $Domain.isVerified
        isRoot              = $Domain.isRoot
        isInitial           = $Domain.isInitial
        isAdminManaged      = $Domain.isAdminManaged
        isDefault           = $Domain.isDefault
        authenticationType  = $Domain.authenticationType
        email              = $Email
        intune             = $Intune
        officeCommunicationsOnline = $OfficeCommunicationsOnline
        orgIdAuthentication = $OrgIdAuthentication
        supportedServices   = $Domain.supportedServices -join ";"
        passwordNotificationWindowInDays = $Domain.passwordNotificationWindowInDays
        passwordValidityPeriodInDays       = $Domain.passwordValidityPeriodInDays
    }
    $DBReportTenantDomains += $DomainObject
}

Export-Report "DBReportTenantDomains" -Report $DBReportTenantDomains -Path $DBFileTenantDomains -SortProperty "domain"

$UriResource = "policies/crossTenantAccessPolicy/partners"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
[array]$PartnerTenants = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
ForEach ($Tenant in $PartnerTenants) {
    Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
    $UriResource = "tenantRelationships/findTenantInformationByTenantId"
    $UriParam = "tenantId='$($Tenant.tenantId)'"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -EqualsParam $UriParam
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $TenantInfo = Invoke-RestMethod -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON -ErrorAction Stop

    $UriResource = "policies/crossTenantAccessPolicy/partners/$($Tenant.tenantId)/identitySynchronization"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
    Try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $CrossTenantSyncInfo = Invoke-RestMethod -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON -ErrorAction Stop
        $InboundSyncAllowed = $CrossTenantSyncInfo.userSyncInbound.isSyncAllowed
        If (-not $InboundSyncAllowed) {
            $InboundSyncAllowed = $false
        }
    }
    Catch {
        $InboundSyncAllowed = $false
    }
    $automaticConsent_inbound = $Tenant.automaticUserConsentSettings.inboundAllowed;
    If (-not $automaticConsent_inbound) {
        $automaticConsent_inbound = $false
    }
    $automaticConsent_outbound = $Tenant.automaticUserConsentSettings.outboundAllowed;
    If (-not $automaticConsent_outbound) {
        $automaticConsent_outbound = $false
    }

    $TenantData = [pscustomobject]@{
        tenantId                        = $Tenant.tenantId;
        displayName                     = $TenantInfo.displayName;
        defaultDomainName               = $TenantInfo.defaultDomainName;
        b2bCollaborationOutbound        = $Tenant.b2bCollaborationOutbound;
        b2bCollaborationInbound         = $Tenant.b2bCollaborationInbound;
        
        inboundTrust_MFAAccepted        = $Tenant.inboundTrust.isMfaAccepted;
        inboundTrust_ComplDevAccepted   = $Tenant.inboundTrust.isCompliantDeviceAccepted;
        inboundTrust_HAADJAccepted      = $Tenant.inboundTrust.isHybridAzureADJoinedDeviceAccepted;
        
        b2bDCOutbound_UsrGrp_accessType = $Tenant.b2bDirectConnectOutbound.usersAndGroups.accessType;
        b2bDCOutbound_UsrGrp_target     = $Tenant.b2bDirectConnectOutbound.usersAndGroups.targets.target;
        b2bDCOutbound_UsrGrp_targetType = $Tenant.b2bDirectConnectOutbound.usersAndGroups.targets.targetType;
        b2bDCOutbound_Apps_accessType   = $Tenant.b2bDirectConnectOutbound.applications.accessType;
        b2bDCOutbound_Apps_target       = $Tenant.b2bDirectConnectOutbound.applications.targets.target;
        b2bDCOutbound_Apps_targetType   = $Tenant.b2bDirectConnectOutbound.applications.targets.targetType;

        b2bDCInbound_UsrGrp_accessType  = $Tenant.b2bDirectConnectInbound.usersAndGroups.accessType;
        b2bDCInbound_UsrGrp_target      = $Tenant.b2bDirectConnectInbound.usersAndGroups.targets.target;
        b2bDCInbound_UsrGrp_targetType  = $Tenant.b2bDirectConnectInbound.usersAndGroups.targets.targetType;
        b2bDCInbound_Apps_accessType    = $Tenant.b2bDirectConnectInbound.applications.accessType;
        b2bDCInbound_Apps_target        = $Tenant.b2bDirectConnectInbound.applications.targets.target;
        b2bDCInbound_Apps_targetType    = $Tenant.b2bDirectConnectInbound.applications.targets.targetType;

        automaticConsent_inbound        = $automaticConsent_inbound;
        automaticConsent_outbound       = $automaticConsent_outbound;

        InboundSyncAllowed              = $InboundSyncAllowed
    }
    
    $DBReportAADPartnerTenants += $TenantData 
    $AADPartnerTenants_DB.Add($Tenant.tenantId, $TenantData)
    Start-Sleep -Milliseconds $ThrottlingDelayPerTenantInMsec
}

Export-Report "DBReportAADPartnerTenants" -Report $DBReportAADPartnerTenants -Path $DBFileAADPartnerTenants

$UriResource = "users"
$UriFilter = "UserType+eq+'Guest'"
$UriSelect = "mail"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Filter $UriFilter -Select $UriSelect
[array]$AllAADGuests = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
ForEach ($Guest in $AllAADGuests) {
    if ($Guest.mail) {
        $Domain = $Guest.Mail.Split("@")[1]
        $GuestDomainArray += $Domain.ToLower()
    }
}
$GuestDomainArray = $GuestDomainArray | Sort-Object -unique

ForEach ($Domain in $GuestDomainArray) {
    Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
    $queryError,$queryErrorParentDomain,$domainNotFound,$3rdLevelDomain,$triedParentDomain = $false
    [string]$parentDomain = $null
    if (([regex]::matches($domain,".").count) -ge 2) {
        $3rdLevelDomain = $true
        $parentDomain = $domain.Substring($domain.IndexOf(".") + 1)
    }
    
    $UriResource = "tenantRelationships/findTenantInformationByDomainName"
    $UriParam = "domainName='$($Domain)'"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -EqualsParam $UriParam
    Try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $Tenant = Invoke-RestMethod -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON -ErrorAction Stop
        If ($Tenant) {
            $InboundSyncAllowed = $false
            If ($AADPartnerTenants_DB.Contains($Tenant.tenantId)) {
                $InboundSyncAllowed = $AADPartnerTenants_DB[$Tenant.tenantId].InboundSyncAllowed
                If (-not $InboundSyncAllowed) {
                    $InboundSyncAllowed = $false
                }
            }
        }
        $AADDomainData = [pscustomobject]@{
            domain              = $domain;
            tenantId            = $Tenant.tenantId;
            displayName         = $Tenant.displayName.Trim();
            defaultDomainName   = $Tenant.defaultDomainName;
            InboundSyncAllowed  = $InboundSyncAllowed
        }
        $DBReportExtAADTenants += $AADDomainData
    }
    Catch { 
        $QueryError = $true
        $QueryErrorMsg = $_.Exception.Message
        if ($_.Exception.Message.Contains("(404) Not Found")) {
            $domainNotFound = $true
        }
        if (($_.Exception.Message.Contains("(500) Internal Server Error")) -and $3rdLevelDomain -and $parentDomain) {
            $triedParentDomain = $true
            $UriResource = "tenantRelationships/findTenantInformationByDomainName"
            $UriParam = "domainName='$($parentDomain)'"
            $Uri = New-GraphUri -Version "beta" -Resource $UriResource -EqualsParam $UriParam
            Try {
                $QueryParentDomain = Invoke-RestMethod -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON -ErrorAction Stop
                $AADDomainData = [pscustomobject]@{
                    domain              = $domain;
                    tenantId            = "???" + $QueryParentDomain.tenantId;
                    displayName         = "??? " + $QueryParentDomain.displayName.Trim();
                    defaultDomainName   = "???" + $QueryParentDomain.defaultDomainName;
                    InboundSyncAllowed  = "n/a"
                }
                $DBReportExtAADTenants += $AADDomainData
                write-host "$($domain)" -ForegroundColor Magenta -NoNewline
                write-host ": successfully read parent domain data: $($QueryParentDomain.displayName), $($QueryParentDomain.defaultDomainName)"
            }
            Catch {
                $QueryErrorParentDomain = $true
                Write-Log "ERR: ($($parentDomain)) $($_.Exception.Message)" -MessageType Error
            }
        }
        if ((-not $domainNotFound) -and (-not $triedParentDomain)){ 
            Write-Log "ERR: ($($Domain)) $($QueryErrorMsg)" -MessageType Error	
        }
    }
    Start-Sleep -Milliseconds $ThrottlingDelayPerDomainInMsec
}

Export-Report "DBReportExtAADTenants" -Report $DBReportExtAADTenants -Path $DBFileExtAADTenants -SortProperty "domain"

. $IncFile_StdLogEndBlock
