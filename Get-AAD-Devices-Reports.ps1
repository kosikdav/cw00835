#######################################################################################################################
# Get-AAD-Devices-Reports
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "exports"
$LogFilePrefix		= "aad-devices-report" 
$OutputFolder		= "aad-devices\reports"
$OutputFilePrefix	= "aad-devices"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile    = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Ext "csv"

[array]$DevicesReport = @()
#[array]$PhysicalIdList = @("[GID]","[HWID]","[USER-GID]","[USER-HWID]","[ZTDID]","[OrderId]")
#######################################################################################################################

. $IncFile_StdLogBeginBlock

Write-Log "Getting AAD devices list as of: $(Get-Date -Format "yyyy-MM-dd HH:MM:ss")"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "devices"
$UriSelect1 = "id,accountEnabled,approximateLastSignInDateTime,complianceExpirationDateTime,deviceCategory,deviceId,deviceOwnership,deviceVersion"
$UriSelect2 = "displayName,enrollmentProfileName,isCompliant,isManaged,manufacturer,mdmAppId,model,onPremisesLastSyncDateTime,onPremisesSyncEnabled"
$UriSelect3 = "operatingSystem,operatingSystemVersion,physicalIds,profileType,registrationDateTime,systemLabels,trustType"
$UriSelect = $UriSelect1,$UriSelect2,$UriSelect3 -join ","
$Expand = "registeredOwners"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect -Expand $Expand
Write-Host $Uri
$Devices = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "AAD devices" -ProgressDots

Write-Log "Total devices: $($Devices.Count)"

Initialize-ProgressBarMain -Activity "Building AAD device report" -Total $Devices.Count
ForEach ($Device in $Devices) {
	Update-ProgressBarMain
    $Owners = [string]::Empty
    $LastSignInDateTime,$mdmAppName = $null
    $DaysSinceLastSignIn = "n/a"
    $GID,$HWID,$USER_GID,$USER_HWID,$ZTDID,$OrderId = $null

    if ($device.physicalIds) {
        foreach ($id in $device.physicalIds) {
            $index = $id.IndexOf(":")
            $prefix = $id.Substring(0,($Index))
            $value = $id.Substring($Index + 1)
            switch ($prefix) {
                "[GID]" { $GID = $value }
                "[HWID]" { $HWID = $value }
                "[USER-GID]" { $USER_GID = $value }
                "[USER-HWID]" { $USER_HWID = $value }
                "[ZTDID]" { $ZTDID = $value }
                "[OrderId]" { $OrderId = $value }
            }
        }
    }
    
    if ($Device.registeredOwners) {
        $Owners = $Device.registeredOwners.userPrincipalName -join ";"
    }
    
    if ($device.approximateLastSignInDateTime) {
        $LastSignInDateTime = [datetime]$Device.approximateLastSignInDateTime
        $DaysSinceLastSignIn = (New-TimeSpan -Start $LastSignInDateTime -End $Today).Days
    }

    if ($Device.mdmAppId) {
        if ($mdmApp_DB.ContainsKey($Device.mdmAppId)) {
            $mdmAppName = $mdmApp_DB[$Device.mdmAppId]
        }
    }

    $DeviceObject = [pscustomobject]@{
        id = $Device.id
        deviceId = $Device.deviceId
        displayName = $Device.displayName
        owners = $Owners
        accountEnabled = $Device.accountEnabled
        lastSignInDateTime = $LastSignInDateTime
        daysSinceLastSignIn = $DaysSinceLastSignIn
        complianceExpirationDateTime = $Device.complianceExpirationDateTime
        deviceCategory = $Device.deviceCategory
        deviceOwnership = $Device.deviceOwnership
        deviceVersion = $Device.deviceVersion
        enrollmentProfileName = $Device.enrollmentProfileName
        profileType = $Device.profileType
        registrationDateTime = $Device.registrationDateTime
        trustType = $Device.trustType
        isCompliant = $Device.isCompliant
        isManaged = $Device.isManaged
        manufacturer = $Device.manufacturer
        model = $Device.model
        mdmAppId = $Device.mdmAppId
        mdmAppName = $mdmAppName
        onPremisesLastSyncDateTime = $Device.onPremisesLastSyncDateTime
        onPremisesSyncEnabled = $Device.onPremisesSyncEnabled
        operatingSystem = $Device.operatingSystem
        operatingSystemVersion  = $Device.operatingSystemVersion
        GID = $GID
        HWID = $HWID
        USER_GID = $USER_GID
        USER_HWID = $USER_HWID
        ZTDID = $ZTDID
        OrderId = $OrderId
    }
    $DevicesReport += $DeviceObject
}

Export-Report -Text "AAD devices report" -Report $DevicesReport -Path $OutputFile -SortProperty "displayName"

#######################################################################################################################

. $IncFile_StdLogEndBlock
