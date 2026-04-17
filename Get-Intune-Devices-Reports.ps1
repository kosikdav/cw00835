#######################################################################################################################
# Get-Intune-Devices-Reports
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "exports"
$LogFilePrefix		= "intune-devices-report" 
$OutputFolder		= "intune-devices\reports"
$OutputFilePrefix	= "intune-devices"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile    = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Ext "csv"

[array]$DevicesReport = @()
#[array]$PhysicalIdList = @("[GID]","[HWID]","[USER-GID]","[USER-HWID]","[ZTDID]","[OrderId]")
#######################################################################################################################

. $IncFile_StdLogStartBlock

Write-Log "Getting AAD devices list as of: $(Get-Date -Format "yyyy-MM-dd HH:MM:ss")"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "deviceManagement/managedDevices"
$UriSelect1 = "id,userId,deviceName,managedDeviceOwnerType,enrolledDateTime,lastSyncDateTime,operatingSystem,complianceState,jailBroken,managementAgent,osVersion"
$UriSelect2 = "easActivated,easDeviceId,azureADRegistered,deviceEnrollmentType,emailAddress,azureADDeviceId,deviceRegistrationState"
$UriSelect3 = "deviceCategoryDisplayName,isSupervised"
$UriSelect4 = "isEncrypted,userPrincipalName,model,manufacturer,imei,complianceGracePeriodExpirationDateTime,serialNumber,phoneNumber,androidSecurityPatchLevel,userDisplayName"
$UriSelect5 = "configurationManagerClientEnabledFeatures,wiFiMacAddress,subscriberCarrier,meid,totalStorageSpaceInBytes,freeStorageSpaceInBytes,managedDeviceName"
$UriSelect6 = "requireUserEnrollmentApproval,managementCertificateExpirationDate,hardwareInformation"
$UriSelect = $UriSelect1,$UriSelect2,$UriSelect3,$UriSelect4,$UriSelect5,$UriSelect6 -join ","
$Uri = New-GraphUri -Version "beta" -Resource $UriResource -Select $UriSelect
$Devices = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "Intune devices" -ProgressDots

Write-Log "Total devices: $($Devices.Count)"

ForEach ($Device in $Devices) {
    $enrolledDateTime = [datetime]$Device.enrolledDateTime
    $lastSyncDateTime,$enrolledDateTime,$imei,$meid = $null
        
    $daysSinceLastSync,$daysSinceEnrolled = "n/a"
    
    if ($Device.enrolledDateTime) {
        $enrolledDateTime = [datetime]$Device.enrolledDateTime
        $DaysSinceEnrolled = (New-TimeSpan -Start $enrolledDateTime -End $Today).Days
    }
    if ($device.lastSyncDateTime) {
        $lastSyncDateTime = [datetime]$Device.lastSyncDateTime
        $DaysSinceLastSync = (New-TimeSpan -Start $lastSyncDateTime -End $Today).Days
    }
    
    if ($Device.complianceGracePeriodExpirationDateTime -eq "9999-12-31T23:59:59Z") {
        $complianceGracePeriodExpiration = $null
    }
    else {
        $complianceGracePeriodExpiration = [datetime]$Device.complianceGracePeriodExpirationDateTime
    }

    if ($Device.imei) {
        $imei = [char]34 + $Device.imei + [char]34
    }
    
    if ($Device.meid) {
        $meid = [char]34 + $Device.meid + [char]34
    }
    $HWInfo = $Device.hardwareInformation

    $DeviceObject = [pscustomobject]@{
        id = $Device.id;
        deviceName = $Device.deviceName;
        managedDeviceName = $Device.managedDeviceName;
        userId = $Device.userId;
        userPrincipalName = $Device.userPrincipalName;
        emailAddress = $Device.emailAddress;
        userDisplayName = $Device.userDisplayName;
        enrolledDateTime = $enrolledDateTime;
        daysSinceEnrolled = $DaysSinceEnrolled;
        lastSyncDateTime = $lastSyncDateTime;
        daysSinceLastSync = $DaysSinceLastSync;
        deviceEnrollmentType = $Device.deviceEnrollmentType;
        managedDeviceOwnerType = $Device.managedDeviceOwnerType;
        operatingSystem = $Device.operatingSystem;
        osVersion = $Device.osVersion;
        androidSecurityPatchLevel = $Device.androidSecurityPatchLevel;
        complianceState = $Device.complianceState;
        jailBroken = $Device.jailBroken;
        managementAgent = $Device.managementAgent;
        easActivated = $Device.easActivated;
        easDeviceId = $Device.easDeviceId;
        azureADRegistered = $Device.azureADRegistered;
        azureADDeviceId = $Device.azureADDeviceId;
        deviceRegistrationState = $Device.deviceRegistrationState;
        deviceCategoryDisplayName = $Device.deviceCategoryDisplayName;
        isSupervised = $Device.isSupervised;
        isEncrypted = $Device.isEncrypted;
        manufacturer = $Device.manufacturer;
        model = $Device.model;
        imei = $imei;
        meid = $meid;
        serialNumber = $Device.serialNumber;
        complianceGracePeriodExpiration = $complianceGracePeriodExpiration;
        phoneNumber = $Device.phoneNumber;
        configurationManagerClientEnabledFeatures = $Device.configurationManagerClientEnabledFeatures;
        wiFiMacAddress = $Device.wiFiMacAddress;
        subscriberCarrier = $Device.subscriberCarrier;
        totalStorageSpaceInBytes = $Device.totalStorageSpaceInBytes;
        freeStorageSpaceInBytes = $Device.freeStorageSpaceInBytes;
        requireUserEnrollmentApproval = $Device.requireUserEnrollmentApproval;
        managementCertificateExpirationDate = $Device.managementCertificateExpirationDate

        hw_sn = $HWInfo.serialNumber
        hw_totalStorageSpace = $HWInfo.totalStorageSpace;
        hw_freeStorageSpace = $HWInfo.freeStorageSpace;
        hw_imei= $HWInfo.imei;
        hw_meid= $HWInfo.meid;
        hw_manufacturer= $HWInfo.manufacturer;
        hw_model= $HWInfo.model;
        hw_phoneNumber= $HWInfo.phoneNumber;
        hw_subscriberCarrier= $HWInfo.subscriberCarrier;
        hw_cellularTechnology= $HWInfo.cellularTechnology;
        hw_wifiMac = $HWInfo.wifiMac
        hw_operatingSystemLanguage = $HWInfo.operatingSystemLanguage;
        hw_isSupervised = $HWInfo.isSupervised;
        hw_isEncrypted = $HWInfo.isEncrypted;
        hw_batterySerialNumber = $HWInfo.batterySerialNumber;
        hw_batteryHealthPercentage = $HWInfo.batteryHealthPercentage;
        hw_batteryChargeCycles = $HWInfo.batteryChargeCycles;
        hw_isSharedDevice = $HWInfo.isSharedDevice;
        hw_tpmSpecificationVersion = $HWInfo.tpmSpecificationVersion;
        hw_operatingSystemEdition = $HWInfo.operatingSystemEdition;
        hw_deviceFullQualifiedDomainName = $HWInfo.deviceFullQualifiedDomainName;
        hw_dgrdVsecHWReqState= $HWInfo.deviceGuardVirtualizationBasedSecurityHardwareRequirementState;
        hw_dgrdVsecState= $HWInfo.deviceGuardVirtualizationBasedSecurityState;
        hw_dgrdLSACredGuardState= $HWInfo.deviceGuardLocalSystemAuthorityCredentialGuardState;
        hw_osBuildNumber= $HWInfo.osBuildNumber;
        hw_operatingSystemProductType= $HWInfo.operatingSystemProductType;
        hw_ipAddressV4 = $HWInfo.ipAddressV4;
        hw_subnetAddress = $HWInfo.subnetAddress;
        hw_esimIdentifier = $HWInfo.esimIdentifier;
        hw_systemManagementBIOSVersion = $HWInfo.systemManagementBIOSVersion;
        hw_tpmManufacturer = $HWInfo.tpmManufacturer;
        hw_tpmVersion = $HWInfo.tpmVersion;
        hw_wiredIPv4Addresses = $HWInfo.wiredIPv4Addresses -join ";";
        hw_batteryLevelPercentage = $HWInfo.batteryLevelPercentage;
        hw_residentUsersCount = $HWInfo.residentUsersCount;
        hw_productName = $HWInfo.productName;
        hw_deviceLicensingStatus = $HWInfo.deviceLicensingStatus;
        hw_deviceLicensingLastErrorCode = $HWInfo.deviceLicensingLastErrorCode;
        hw_deviceLicensingLastErrorDescription = $HWInfo.deviceLicensingLastErrorDescription;
        hw_sharedDeviceCachedUsers = $HWInfo.sharedDeviceCachedUsers -join ";"
    }
    $DevicesReport += $DeviceObject
}

Export-Report -Text "Intune device report" -Report $DevicesReport -Path $OutputFile -SortProperty "displayName"

#######################################################################################################################

. $IncFile_StdLogStartBlock
