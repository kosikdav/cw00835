#######################################################################################################################
# Set-AAD-MFA-Phone-From-IDM
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    [switch]$FullRun
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder          = "aad-mfa"
$LogFilePrefix		= "aad-mfa"
$LogFileSuffix		= "set-phone-from-IDM"
$LogFileFreq		= "YMD"

$OutputFolder		= "aad-mfa\reports"
$OutputFilePrefix	= "aad-mfa"
$OutputFileSuffix	= "set-phone-from-IDM"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Suffix $LogFileSuffix -Ext "log"
$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffix -Ext "csv"

$DoNotConfigureFromIDM = @()

[array]$MFAPhoneReport = @()
[array]$InvalidPhoneNumberList = @()
[int]$ThrottlingDelayPerUserinMsec = 300
[int]$PhoneNumberPropagationDelayinSec = 30

######################################################################################################################

. $IncFile_StdLogStartBlock

# load DB mfa-mgmt from file or initialize empty
if (test-path $DBFileMFAMgmt) {
    Try {
        $MFAMgmt_DB = Import-Clixml -Path $DBFileMFAMgmt
        Write-Log "DB file $($DBFileMFAMgmt) imported successfully, $($MFAMgmt_DB.count) records found"
    } 
    Catch {
        Write-Log "Error importing $($DBFileMFAMgmt), creating empty DB" -MessageType "Error"
        [hashtable]$MFAMgmt_DB = @{}
        $DB_changed = $true
    }
}
else {
    Write-Log "DB file $($DBFileMFAMgmt) not found, creating empty DB" -MessageType "Error"
    [hashtable]$MFAMgmt_DB = @{}
    $DB_changed = $true
}

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriFilter = "userType eq 'Member' and accountEnabled eq true and onPremisesSyncEnabled eq true"
$UriSelect1 = "id,userPrincipalName,mail,displayName,onPremisesSyncEnabled,onpremisesSamAccountName,onPremisesDistinguishedName,mobilePhone"
$UriSelect2 = "extension_008a5d3f841f4052ac1283ff4782c560_cEZIntuneMFAAuthMobile"
$UriSelect3 = "extension_008a5d3f841f4052ac1283ff4782c560_msExchExtensionAttribute40"
$UriSelect = $UriSelect1, $UriSelect2, $UriSelect3 -join ","
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Filter $UriFilter -Select $UriSelect
$AADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ProgressDots -Text "AAD member users"

#######################################################################################################################
foreach ($User in $AADUsers) {
    Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
    $CurrentMFADBRecord = $null
    $UPN = $samAccountName = $null
    $operation = $sysPrefEnabled = $usrPrefMethod = $sysPrefMethod = $targetMethod = $null
    $phoneNumbersMatch = $false
    $phoneMethodSetSuccessfully = $false
    $signInPreferencesSetSuccessfully = $false
    $CurrentMFAPhone = $IDMAuthPhone = $mobile =$null

    $UPN = $User.UserPrincipalName
    $samAccountName = $User.onpremisesSamAccountName
    $DN = $User.onPremisesDistinguishedName

    if ($UPN -and ($UPN.Substring(0,2) -in $NoMFAPhoneMgmtAccountPrefixes)) {
        Continue
    }

    if ($UPN -and ($UPN.EndsWith($DefaultUPNSuffix))) {
        Continue
    }

    if ($samAccountName -and ($samAccountName.Substring(0,2) -in $NoMFAPhoneMgmtAccountPrefixes)) {
        Continue
    }

    if ($DN -and ($DN.EndsWith($OU_ServiceAccounts))) {
        Continue
    }

    $CurrentMFADBRecord = $null
    $operation = $sysPrefEnabled = $usrPrefMethod = $sysPrefMethod = $targetMethod = $null
    $phoneNumbersMatch = $false
    $phoneMethodSetSuccessfully = $false
    $signInPreferencesSetSuccessfully = $false

    if ($User.MobilePhone) {
        $mobile = Get-IntlFormatPhoneNumber -PhoneNumber $User.MobilePhone -EntraMFAFormat
    }
    else {
        $mobile = "none"
    }

    if ($User.extension_008a5d3f841f4052ac1283ff4782c560_cEZIntuneMFAAuthMobile) {
        $IDMAuthPhone = Get-IntlFormatPhoneNumber -PhoneNumber $User.extension_008a5d3f841f4052ac1283ff4782c560_cEZIntuneMFAAuthMobile -EntraMFAFormat
    }
    else {
        $IDMAuthPhone = "none"
    }

    if ($MFAMgmt_DB.ContainsKey($User.Id)) {
        $CurrentMFADBRecord = $MFAMgmt_DB[$User.Id]
        if (($CurrentMFADBRecord.MFAphone -eq $IDMAuthPhone) -or ($CurrentMFADBRecord.MFAphone -eq $mobile)) {
            #skipping, DB phone and IDM phones match or DB phone and mobile match, nothing to do
            if (-not $FullRun) {
                continue
            }
        }
    }

    if (($mobile -eq "none") -and ($IDMAuthPhone -eq "none")) {
        # no mobile number in AAD and no auth phone in IDM, skip user
        $operation = "skip-no-numbers"
        $match = "SKIP"
        $color = "DarkGray"
        <#
        if ($interactiveRun) {
            write-host "$($User.UserPrincipalName.PadRight(40," ")) $($samAccountName.PadRight(14," ")) IDM:$($IDMAuthPhone.PadRight(15," ")) mobile:$($mobile.PadRight(15," ")) " -NoNewline -ForegroundColor DarkGray
            write-host $match -ForegroundColor $color
        }
        #>
        continue
    }
    #at this point we know that at least one of the two numbers (mobile in AAD or auth phone in IDM) is populated

    $UriResource = "users/$($User.Id)/authentication/phoneMethods/$($mobilePhoneMethodId)"
    $Uri = New-GraphUri -Version "beta" -Resource $UriResource
    Try {
        $ErrorMessageGET = "success"
        $ResponseGET = Invoke-WebRequest -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON -UseBasicParsing
        $MobilePhoneMethod = $ResponseGET | ConvertFrom-Json
        $CurrentMFAPhone = Get-IntlFormatPhoneNumber -PhoneNumber $MobilePhoneMethod.phoneNumber -EntraMFAFormat
    }
    Catch {
    $ErrorMessageGET = $_.Exception.Message
        If ($ErrorMessageGET.Contains("(404)")) {
            $CurrentMFAPhone = "none"
        }
        Else {
            Write-Log "$($UPN) ($($User.displayName)): Error reading MFA phoneMethods: $($ErrorMessageGET)" -MessageType "ERROR" -ForceOnScreen
            Continue
        }
    }

    #at this point we have the current MFA phone number (if any) from AAD, 
    #the auth phone number from IDM and the mobile phone number from AAD (if any), 
    #let's compare and decide what to do
    
    if ($CurrentMFAPhone -eq "none") {
        # AAD MFA empty
        $operation = "new"
        $match = "NONE"
        $color = "Yellow"
    }
    else {
        If (($CurrentMFAPhone -eq $IDMAuthPhone) -or ($CurrentMFAPhone -eq $mobile)) {
            # numbers in IDM and AAD MFA match, or mobile and AAD MFA match = nothing to do
            $phoneNumbersMatch = $true
            $operation = "ok-skip"
            $match = "OK"
            $color = "Green"
        }
        Else {
            # numbers do not match
            $match = "DIFF"
            $color = "Red"
        }
    }

    if ($interactiveRun) {
        Write-Host "$($User.UserPrincipalName.PadRight(40," ")) $($samAccountName.PadRight(14," ")) IDM:$($IDMAuthPhone.PadRight(15," ")) mobile:$($mobile.PadRight(15," ")) AAD-MFA:$($CurrentMFAPhone.PadRight(15," ")) " -NoNewline
        Write-Host $match -ForegroundColor $color -NoNewline
    }

    If (($match -eq "NONE") -or ($match -eq "DIFF")) {
        # if there is a phone number in IDM use that, otherwise fallback to mobile (which might be the same number but at least we tried)
        if ($IDMAuthPhone -ne "none") {
            $targetNumber = $IDMAuthPhone
        }
        else {
            $targetNumber = $mobile
        }
        #need to create or update MFA phone method
        If ($CurrentMFAPhone -eq "none") {
            #configure new MFA number
            $operation = "new"
        }
        else {
            #update existing MFA number
            If ($IDMAuthPhone -ne $CurrentMFAPhone) {
                #update - different MFA number
                $operation = "update-number"
            } 
            else {
                #update - current number OK but incorrect format
                $operation = "update-format"
            }
            #current number needs to be deleted first
            Try {
                #write-host "Invoke-WebRequest -Uri $Uri -Method DELETE -UseBasicParsing" -ForegroundColor Magenta
                $ResponseDELETE = Invoke-WebRequest -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Method "DELETE" -UseBasicParsing
                if($interactiveRun) {
                    Write-Host
                }
                Write-Log "$($UPN) ($($User.displayName)): MFA phone $($CurrentMFAPhone) deleted"
            }
            Catch {
                $errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                Write-Log "$($UPN) ($($User.displayName)): Error deleting MFA phone $($CurrentMFAPhone): $($errObj.error.code)" -MessageType "ERROR" -ForceOnScreen
                Continue
            }
        }

        $UriResource = "users/$($user.Id)/authentication/phoneMethods"
        $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
        $GraphBody = [pscustomobject]@{
            phoneNumber = $targetNumber
            phoneType = "mobile"
        } | ConvertTo-Json
        Try {
            #configure MFA number - auth phone
            #write-host "Invoke-WebRequest -Uri $Uri -Method POST -ContentType $ContentTypeJSON -UseBasicParsing -Body $GraphBody" -ForegroundColor Magenta
            $ResponsePOST = Invoke-WebRequest -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $GraphBody -Method "POST" -ContentType $ContentTypeJSON -UseBasicParsing
            if($interactiveRun) {
                Write-Host
            }
            Write-Log "$($UPN) ($($User.displayName)): MFA phone configured: $($targetNumber) ($($operation))"
            $phoneNumberConfigured = $targetNumber
            $phoneMethodSetSuccessfully = $true
        }
        Catch {
            $errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
            if ($errObj.error.code -eq "invalidPhoneNumber") {
                $InvalidPhoneNumberList += "$($UPN) - $($IDMAuthPhone)"
            }
            if($interactiveRun) {
                Write-Host
            }
            Write-Log "$($UPN) ($($User.displayName)): Error configuring MFA phone: $($IDMAuthPhone) ($($operation)) - $($errObj.error.code)" -MessageType "ERROR" -ForceOnScreen
            # if the error is invalidPhoneNumber and we have a mobile number, try to use that instead
            if (($errObj.error.code -eq "invalidPhoneNumber") -and ($mobile -ne "none") -and ($mobile -ne $IDMAuthPhone)) {
                Write-Log "$($UPN) ($($User.displayName)): auth phone number invalid, fallback to mobile"
                $GraphBody = [pscustomobject]@{
                    phoneNumber = $mobile
                    phoneType = "mobile"
                } | ConvertTo-Json
                Try {
                    #configure MFA number - mobile phone
                    $ResponsePOST = Invoke-WebRequest -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $GraphBody -Method "POST" -ContentType $ContentTypeJSON -UseBasicParsing
                    Write-Log "$($UPN) ($($User.displayName)): MFA phone configured: $($IDMAuthPhone) ($($operation))"
                    $phoneNumberConfigured = $mobile
                    $phoneMethodSetSuccessfully = $true
                }
                Catch {
                    $errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                    Write-Log "$($UPN) ($($User.displayName)): Error configuring MFA phone: $($IDMAuthPhone) ($($operation)) - $($errObj.error.code)" -MessageType "ERROR" -ForceOnScreen
                }
            } 
        }
    }  
    
    if ($phoneMethodSetSuccessfully) {
        Start-Sleep -Seconds $PhoneNumberPropagationDelayinSec
    }

    ########################################################################################################
    ########################################################################################################
    ########################################################################################################

    $UriResource = "users/$($user.Id)/authentication/signInPreferences"
    $Uri = New-GraphUri -Version "beta" -Resource $UriResource
    Try {
        $ResponseGET = Invoke-WebRequest -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON -UseBasicParsing
        $SignInPreferences = $ResponseGET | ConvertFrom-Json
        $sysPrefEnabled = $SignInPreferences.isSystemPreferredAuthenticationMethodEnabled
        $usrPrefMethod  = $SignInPreferences.userPreferredMethodForSecondaryAuthentication
        $sysPrefMethod  = $SignInPreferences.systemPreferredAuthenticationMethod
        If (-not ($sysPrefMethod -eq $UsrToSysMethodConv_DB[$usrPrefMethod])) {
            $targetMethod = $SysToUsrMethodConv_DB[$sysPrefMethod]
            $UriResource = "users/$($user.Id)/authentication/signInPreferences"
            $Uri = New-GraphUri -Version "beta" -Resource $UriResource
            $GraphBody = [pscustomobject]@{
                userPreferredMethodForSecondaryAuthentication = $targetMethod
            } | ConvertTo-Json
            Try {
                #configure preferred auth method
                $ResponsePATCH = Invoke-WebRequest -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $GraphBody -Method "PATCH" -ContentType $ContentTypeJSON -UseBasicParsing
                Write-Log "$($UPN): userPreferredMethod set to: $($targetMethod), previous value: $($usrPrefMethod)"
                $signInPreferencesSetSuccessfully = $true
            }
            Catch {
                $errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                Write-Log "$($UPN): Error configuring preferred auth method $($targetMethod): $($errObj.error.code)" -MessageType "ERROR" -ForceOnScreen
            }
        }
        else {
            if ($interactiveRun) {
                write-host "  OK" -ForegroundColor Cyan
            }
            $signInPreferencesSetSuccessfully = $true
        }
    }
    Catch {
        $ErrorMessageGET = $_.Exception.Message
        $errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
        If (-not ($ErrorMessageGET.Contains("(404)"))) {
            Write-Log "$($UPN): Error reading signin preferences: $($errObj.error.code) - $($ErrorMessageGET)" -MessageType "ERROR" -ForceOnScreen
        }
    }

    $MFAPhoneReport += [pscustomobject]@{
        UserPrincipalName = $user.userPrincipalName
        Id                = $User.Id
        KPJM              = $User.onpremisesSamAccountName
        DisplayName       = $User.displayName
        Mail              = $User.mail
        Mail_40           = $User.extension_008a5d3f841f4052ac1283ff4782c560_msExchExtensionAttribute40
        mobile            = $User.mobile
        CurrentMFAPhone   = $CurrentMFAPhone
        IDMAuthPhone      = $IDMAuthPhone
        Operation         = $operation
        sysPrefEnabled    = $sysPrefEnabled
        usrPrefMethod     = $usrPrefMethod
        sysPrefMethod     = $sysPrefMethod
        targetMethod      = $targetMethod
    }

    if ($phoneNumbersMatch -or ($phoneMethodSetSuccessfully -and $signInPreferencesSetSuccessfully)) {
        if ($CurrentMFADBRecord) {
            # update existing record
            $NewMFADBRecord = $CurrentMFADBRecord
        }
        else {
            # create new record
            $NewMFADBRecord = [PSCustomObject]@{
            userId 	= $user.Id
            MFAphone 	= $CurrentMFAPhone
            whenConfigured = $null
            lastUpdated = $null
            }
        }
        if ($phoneMethodSetSuccessfully -and $signInPreferencesSetSuccessfully) {
            $NewMFADBRecord.MFAphone = $phoneNumberConfigured
            $NewMFADBRecord.whenConfigured = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
        $NewMFADBRecord.lastUpdated = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")

        if ($CurrentMFADBRecord) {
            $MFAMgmt_DB[$user.Id] = $NewMFADBRecord
            $DB_changed = $true
        }
        else {
            $MFAMgmt_DB.Add($user.Id, $NewMFADBRecord)
            $DB_changed = $true
        } 
    }

    if ($FullRun){
        Start-Sleep -Milliseconds $ThrottlingDelayPerUserinMsec
    }
}#foreach ($User in $AADUsers)

if ($InvalidPhoneNumberList) {
    Write-Log $string_divider
    Write-Log "Invalid phone numbers found:" -MessageType "WARNING"
    foreach ($entry in $InvalidPhoneNumberList) {
        Write-Log $entry -MessageType "WARNING"
    }
    Write-Log $string_divider
}
else {
    Write-Log "No invalid phone numbers found."
}

#saving DB XML if needed
if (($MFAMgmt_DB.count -gt 0) -and ($DB_changed)){
  Try {
      $MFAMgmt_DB | Export-Clixml -Path $DBFileMFAMgmt
      Write-Log "DB file $($DBFileMFAMgmt) exported successfully, $($MFAMgmt_DB.count) records saved"
  }
  Catch {
      Write-Log "Error exporting $($DBFileMFAMgmt)" -MessageType "Error"
  }
}

Export-Report "MFA phone report" -Report $MFAPhoneReport -SortProperty "UserPrincipalName" -Path $OutputFile

#######################################################################################################################

. $IncFile_StdLogEndBlock
