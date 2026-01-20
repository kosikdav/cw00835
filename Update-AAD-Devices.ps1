#######################################################################################################################
# Update-AAD-Devices
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder          = "aad-device-mgmt"
$LogFilePrefix      = "aad-device-mgmt"
$LogFileFreq        = "YMD"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Suffix $LogFileSuffix -Ext "log"

$ADCredentialPath = $aad_grp_mgmt_cred

[hashtable]$AADDevice_DB = @{}

[array]$ASIS_VPN_ADM_Devices = @()
[array]$TOBE_VPN_ADM_Devices = @()
[array]$TOBE_VPN_ADM_Devices_Whitelist = @(
    #NB319420-VM
    "398bb09f-54fd-4aab-b087-86eec7b78b25"     
)

#######################################################################################################################

. $IncFile_StdLogStartBlock

if (-not $interactiveRun) {
    Write-Log "AD credential file: $($ADCredentialPath)"
    $ADCredential = Import-Clixml -Path $ADCredentialPath
}

Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30

$UriResource = "devices"
$UriSelect = "id,displayName,deviceId,operatingSystem,trustType,extensionAttributes"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect -Top 999
[array]$AllAADDevices = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -ContentType $ContentTypeJSON -text "AAD devices" -ProgressDots
Write-Log "Total AAD devices: $($AllAADDevices.Count)"

$HAADJ_Devices = $AllAADDevices | Where-Object { $_.operatingSystem -eq "Windows" -and $_.trustType -eq "serverAd" }
Write-Log "Total HAADJ devices: $($HAADJ_Devices.Count)"

foreach ($device in $HAADJ_Devices) {
    $deviceObject = [PSCustomObject]@{
        deviceId    = $device.deviceId
        displayName = $device.displayName.Trim()
    }
    $AADDevice_DB.Add($device.deviceId, $deviceObject)
    if ($device.extensionAttributes.extensionAttribute14 -eq "vpn-adm") {
        $ASIS_VPN_ADM_Devices += $device.deviceId
    }
}

write-host "Total AAD devices in DB: $($AADDevice_DB.Count)"

if (-not $interactiveRun) {
    $TOBE_VPN_ADM_Devices = (Get-ADGroupMember -Identity $VPN_ADM_DeviceGroup -Credential $ADCredential).objectGuid
}
else {
    $TOBE_VPN_ADM_Devices = (Get-ADGroupMember -Identity $VPN_ADM_DeviceGroup).objectGuid
}
$TOBE_VPN_ADM_Devices = $TOBE_VPN_ADM_Devices + $TOBE_VPN_ADM_Devices_Whitelist
$missingDevices = $TOBE_VPN_ADM_Devices | Where-Object { $_ -notin $ASIS_VPN_ADM_Devices }
$extraDevices = $ASIS_VPN_ADM_Devices | Where-Object { $_ -notin $TOBE_VPN_ADM_Devices }
$devicesToProcess = $missingDevices + $extraDevices

Write-Log "ASIS_VPN_ADM_Devices: $($ASIS_VPN_ADM_Devices.Count)"
Write-Log "TOBE_VPN_ADM_Devices: $($TOBE_VPN_ADM_Devices.Count)"
Write-Log "missingDevices: $($missingDevices.Count)"
Write-Log "extraDevices: $($extraDevices.Count)"

foreach ($deviceId in $devicesToProcess) {
    if ($missingDevices.contains($deviceId)) {
        $ext14 = "vpn-adm"
    }
    else {
        $ext14 = $null
    }
    if ($AADDevice_DB.ContainsKey($deviceId)) {
        $device = $AADDevice_DB[$deviceId]
        $UriResource = "/devices(deviceId='$($deviceId)')"
        $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
        $Body = @{
            extensionAttributes = @{
                extensionAttribute14 = $ext14
            }
        } | ConvertTo-Json
        $Body = [System.Text.Encoding]::UTF8.GetBytes($Body)
        Try {
            $ResultPATCH = Invoke-RestMethod -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $Body -Method "PATCH" -ContentType $ContentTypeJSON
            Write-Log "$($device.displayName) - ext14 set to `"$($ext14)`"" -ForegroundColor "Green"
        }
        Catch {
            Write-Log $_.Exception.Message -MessageType "Error"
        }
    }
    else {
        Write-Log "$($deviceId) - not found in HAADJ devices" -ForegroundColor "Red"
    }
}

#######################################################################################################################
# Sync Intune Serial Numbers to Entra Device extensionAttribute13
#######################################################################################################################

Write-Log $string_divider
Write-Log "Syncing Intune serial numbers to Entra extensionAttribute13"

# Build lookup of Entra devices by deviceId with current ext13 value
[hashtable]$EntraDevice_Ext13_DB = @{}
foreach ($device in $AllAADDevices) {
    $EntraDevice_Ext13_DB[$device.deviceId] = @{
        id = $device.id
        displayName = $device.displayName
        ext13 = $device.extensionAttributes.extensionAttribute13
    }
}

# Get all Intune managed devices
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "deviceManagement/managedDevices"
$UriSelect = "id,deviceName,serialNumber,azureADDeviceId"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect -Top 999
[array]$IntuneDevices = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "Intune devices" -ProgressDots

Write-Log "Total Intune devices: $($IntuneDevices.Count)"

# Filter devices with serial number and valid Azure AD device ID
$IntuneDevicesWithSerial = $IntuneDevices | Where-Object {
    $_.serialNumber -and
    ($_.serialNumber.Trim() -ne "0" -and $_.serialNumber -ne 0 -and $_.serialNumber.Trim() -ne "") -and
    $_.azureADDeviceId -and
    $_.azureADDeviceId -ne "00000000-0000-0000-0000-000000000000"
}

Write-Log "Intune devices with serial number: $($IntuneDevicesWithSerial.Count)"

[int]$serialUpdated = 0
[int]$serialSkipped = 0
[int]$serialNotFound = 0

foreach ($intuneDevice in $IntuneDevicesWithSerial) {
    $azureADDeviceId = $intuneDevice.azureADDeviceId
    $serialNumber = $intuneDevice.serialNumber.Trim()

    if ($EntraDevice_Ext13_DB.ContainsKey($azureADDeviceId)) {
        $entraDevice = $EntraDevice_Ext13_DB[$azureADDeviceId]

        # Check if update is needed
        if ($entraDevice.ext13 -eq $serialNumber) {
            $serialSkipped++
            continue
        }

        # Update extensionAttribute13 with serial number
        $UriResource = "devices/$($entraDevice.id)"
        $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
        $Body = @{
            extensionAttributes = @{
                extensionAttribute13 = $serialNumber
            }
        } | ConvertTo-Json
        $Body = [System.Text.Encoding]::UTF8.GetBytes($Body)

        Try {
            $ResultPATCH = Invoke-RestMethod -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $Body -Method "PATCH" -ContentType $ContentTypeJSON
            Write-Log "$($entraDevice.displayName) - ext13 set to `"$($serialNumber)`"" -ForegroundColor "Green"
            $serialUpdated++
        }
        Catch {
            Write-Log "Error updating $($entraDevice.displayName): $($_.Exception.Message)" -MessageType "Error"
        }
    }
    else {
        $serialNotFound++
    }
}

Write-Log "Serial number sync complete - Updated: $($serialUpdated), Skipped (already set): $($serialSkipped), Not found in Entra: $($serialNotFound)"

#######################################################################################################################

. $IncFile_StdLogEndBlock
