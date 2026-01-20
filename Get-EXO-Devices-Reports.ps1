#######################################################################################################################
# Get-EXO-Devices-Reports
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder		= "exports"
$LogFilePrefix  = "exo-reports"

$OutputFolder       = "exo\reports"
$OutputFilePrefix	= "exo"
$OutputFileSuffixBasic	= "device-list-basic"
$OutputFileSuffixFull	= "device-list-full"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputFileBasic    = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixBasic -Ext "csv"
$OutputFileFull     = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixFull -Ext "csv"

[hashtable]$MailboxStats_DB = @{}
[hashtable]$AADUsers_DB = @{}
[array]$EASReport = @()
[array]$DeviceReportFull = @()
[array]$DeviceReportBasic = @()


#######################################################################################################################

. $ScriptPath\include-Script-StartLog-Generic.ps1

<#
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
#>

# retrieve devices ##################################### 
Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 120
write-host "Get-MobileDevice..." -NoNewline
[array]$Devices = Get-MobileDevice -ResultSize Unlimited
write-host "done ($($Devices.Count))"
$EASDevices = $Devices | Where-Object {($_.ClientType -eq 'EAS' -or $_.ClientType -match 'ActiveSync')}
write-host "All devices: $($Devices.Count)"
write-host "EAS devices: $($EASDevices.Count)"

$Counter = 0
ForEach ($Device in $Devices) {
    #$Counter++
    #write-host "$($Counter)/$($Devices.Count) - $Device.Identity" -ForegroundColor Yellow
    $ReportObject = [pscustomobject]@{
        FriendlyName = $Device.FriendlyName
        DeviceId = $Device.DeviceId
        DeviceImei = $Device.DeviceImei
        DeviceMobileOperator = $Device.DeviceMobileOperator
        DeviceOS = $Device.DeviceOS
        DeviceOSLanguage = $Device.DeviceOSLanguage
        DeviceTelephoneNumber = $Device.DeviceTelephoneNumber
        DeviceType = $Device.DeviceType
        DeviceUserAgent = $Device.DeviceUserAgent
        DeviceModel = $Device.DeviceModel
        FirstSyncTime = $Device.FirstSyncTime
        UserDisplayName = $Device.UserDisplayName
        DeviceAccessState = $Device.DeviceAccessState
        DeviceAccessStateReason = $Device.DeviceAccessStateReason
        DeviceAccessControlRule = $Device.DeviceAccessControlRule
        ClientVersion = $Device.ClientVersion
        ClientType = $Device.ClientType
        IsManaged = $Device.IsManaged
        IsCompliant = $Device.IsCompliant
        IsDisabled = $Device.IsDisabled
        AdminDisplayName = $Device.AdminDisplayName
        ExchangeVersion = $Device.ExchangeVersion
        Name = $Device.Name
        DistinguishedName = $Device.DistinguishedName
        Identity = $Device.Identity
        ObjectCategory = $Device.ObjectCategory
        ObjectClass = $Device.ObjectClass
        WhenChanged = $Device.WhenChanged
        WhenCreated = $Device.WhenCreated
        WhenChangedUTC = $Device.WhenChangedUTC
        WhenCreatedUTC = $Device.WhenCreatedUTC
        ExchangeObjectId = $Device.ExchangeObjectId
        OrganizationalUnitRoot = $Device.OrganizationalUnitRoot
        OrganizationId = $Device.OrganizationId
        Guid = $Device.Guid
        OriginatingServer = $Device.OriginatingServer
        IsValid = $Device.IsValid
        ObjectState = $Device.ObjectState
    }
    $DeviceReportBasic += $ReportObject
}
Export-Report "device report - basic" -Report $DeviceReportBasic -Path $OutputFileBasic

$Timer = New-Object System.Diagnostics.Stopwatch
$Timer.Start()
$Counter = 0
ForEach ($Device in $Devices) {
    $Counter++
    $elapsed = $Timer.Elapsed.TotalSeconds
    $avgTimePerDevice = $elapsed / $Counter
    $remainingDevices = $Devices.Count - $Counter
    $estimatedRemainingSeconds = [math]::Round($avgTimePerDevice * $remainingDevices,2)
    $remainingTime = [TimeSpan]::FromSeconds($estimatedRemainingSeconds)
    write-host "$($Counter)/$($Devices.Count) - $($Device.Identity) remaining time: $($remainingTime)" -ForegroundColor Yellow
    $stat = $null
    Try{
        Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 120
        $stat = Get-MobileDeviceStatistics -Identity $Device.Id
    }
    Catch{
        Write-Log "ERROR retrieving statistics for device: $($Device.Identity)" -MessageType "ERR"
    }

    $ReportObject = [pscustomobject]@{
        FriendlyName = $Device.FriendlyName
        DeviceId = $Device.DeviceId
        DeviceOS = $Device.DeviceOS
        DeviceOSLanguage = $Device.DeviceOSLanguage
        DeviceType = $Device.DeviceType
        DeviceUserAgent = $Device.DeviceUserAgent
        DeviceFriendlyName = $stat.DeviceFriendlyName
        DeviceModel = $Device.DeviceModel
        UserDisplayName = $Device.UserDisplayName
        DeviceAccessStateReason = $Device.DeviceAccessStateReason
        ClientVersion = $Device.ClientVersion
        ClientType = $Device.ClientType
        ExchangeVersion = $Device.ExchangeVersion
        Name = $Device.Name
        Identity = $Device.Identity
        WhenCreated = $Device.WhenCreated
        FirstSyncTime = $Device.FirstSyncTime
        LastPolicyUpdateTime = $stat.LastPolicyUpdateTime
        LastSyncAttemptTime = $stat.LastSyncAttemptTime
        LastSuccessSync = $stat.LastSuccessSync
        LastPingHeartbeat = $stat.LastPingHeartbeat
        ExchangeObjectId = $Device.ExchangeObjectId
        Guid = $Device.Guid

        IsRemoteWipeSupported = $stat.IsRemoteWipeSupported
        Status = $stat.Status
        StatusNote = $stat.StatusNote
        DevicePolicyApplied = $stat.DevicePolicyApplied
        DevicePolicyApplicationStatus = $stat.DevicePolicyApplicationStatus
        NumberOfFoldersSynced = $stat.NumberOfFoldersSynced
        DistinguishedName = $Device.DistinguishedName
        ObjectState = $Device.ObjectState
        WhenChanged = $Device.WhenChanged        
    }
    $ReportObject | Export-Csv -Path $OutputFileFull -Append -NoTypeInformation -Encoding "UTF8" -Delimiter ","
}

#######################################################################################################################

. $ScriptPath\include-Script-EndLog-generic.ps1
