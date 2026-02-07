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
$LogFilePrefix      = "aad-mfa"
$LogFileSuffix      = "set-phone-from-IDM"
$LogFileFreq        = "YMD"

$OutputFolder       = "aad-mfa\reports"
$OutputFilePrefix	  = "aad-mfa"
$OutputFileSuffix	  = "set-phone-from-IDM"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Suffix $LogFileSuffix -Ext "log"
$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffix -Ext "csv"

$DoNotConfigureFromIDM = @()

[array]$MFAPhoneReport = @()
[array]$InvalidPhoneNumberList = @()
[int]$ThrottlingDelayPerUserinMsec = 300

#######################################################################################################################

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
$UriFilter = "UserType+eq+'Member'&accountEnabled+eq+'True'&onPremisesSyncEnabled+eq+'True'"
#$UriFilter = "startswith(userPrincipalName,'josef.mat')"
$UriSelect1 = "id,userPrincipalName,mail,displayName,onPremisesSyncEnabled,onpremisesSamAccountName,mobile"
$UriSelect2 = "extension_008a5d3f841f4052ac1283ff4782c560_cEZIntuneMFAAuthMobile"
$UriSelect3 = "extension_008a5d3f841f4052ac1283ff4782c560_msExchExtensionAttribute40"
$UriSelect = $UriSelect1, $UriSelect2, $UriSelect3 -join ","
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Filter $UriFilter -Select $UriSelect
$AADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ProgressDots -Text "AAD member users"

#######################################################################################################################
$counter = 0
foreach ($User in $AADUsers) {

  Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
  $UPN = $User.UserPrincipalName

  $CurrentMFADBRecord = $null
  $CurrentMFAPhone = $operation = $sysPrefEnabled = $usrPrefMethod = $sysPrefMethod = $targetMethod = $null
  $phoneNumbersMatch = $false
  $phoneMethodSetSuccessfully = $false
  $signInPreferencesSetSuccessfully = $false
  
    if ($User.Mobile) {
      $mobile = Get-IntlFormatPhoneNumber -PhoneNumber $User.Mobile -EntraMFAFormat
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
      if (($CurrentMFADBRecord.MFAphone -eq $IDMAuthPhone) -or (($IDMAuthPhone -eq "none") -and ($CurrentMFADBRecord.MFAphone -eq $mobile))) {
        write-host "$($User.UserPrincipalName.PadRight(40," ")) skipping, DB phone and IDM phones match"
        continue
      }
    }
    
    if (($mobile -eq "none") -and ($IDMAuthPhone -eq "none")) {
      # no mobile number in AAD and no auth phone in IDM, skip user
      $operation = "skip-no-numbers"
      $match = "SKIP"
      $clr = "DarkGray"
      write-host "$($User.UserPrincipalName.PadRight(40," ")) IDM:$($IDMAuthPhone.PadRight(20," ")) AAD-MFA:$($CurrentMFAPhone) " -NoNewline
      write-host $match -ForegroundColor $clr
      continue
    }



    #$ExchExtensionAttribute40 = $User.extension_008a5d3f841f4052ac1283ff4782c560_msExchExtensionAttribute40
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
    
    If ($CurrentMFAPhone -eq $IDMAuthPhone) {
      # numbers in IDM and AAD MFA match, nothing to do
      $phoneNumbersMatch = $true
      $operation = "ok-skip"
      $match = "OK"
      $clr = "Green"
    }
    Else {
      # numbers do not match
      $match = "DIFF"
      $clr = "Red"
    }
    
    write-host "$($User.UserPrincipalName.PadRight(40," ")) IDM:$($IDMAuthPhone.PadRight(20," ")) AAD-MFA:$($CurrentMFAPhone) " -NoNewline
    write-host $match -ForegroundColor $clr
    
    If (($ErrorMessageGET.Contains("(404)")) -or ($CurrentMFAPhone -ne $IDMAuthPhone)) {
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
          $ResponseDELETE = Invoke-WebRequest -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Method "DELETE" -ContentType $ContentTypeJSON
          Write-Log "$($UPN) ($($User.displayName)): MFA phone $($CurrentMFAPhone) deleted"
        }
        Catch {
          $errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
          Write-Log "$($UPN) ($($User.displayName)): Error deleting MFA phone $($CurrentMFAPhone): $($errObj.error.code)" -MessageType "ERROR" -ForceOnScreen
          Continue
        }
      }
      
      $UriResource = "users/$($userId)/authentication/phoneMethods"
      $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
      $GraphBody = [pscustomobject]@{
        phoneNumber = $IDMAuthPhone
        phoneType = "mobile"
      } | ConvertTo-Json
      Try {
        #configure MFA number - auth phone
        $ResponsePOST = Invoke-WebRequest -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $GraphBody -Method "POST" -ContentType $ContentTypeJSON
        Write-Log "$($UPN) ($($User.displayName)): MFA phone configured: $($IDMAuthPhone) ($($operation))"
        $phoneNumberConfigured = $IDMAuthPhone
        $phoneMethodSetSuccessfully = $true
      }
      Catch {
        $errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
        if ($errObj.error.code -eq "invalidPhoneNumber") {
          $InvalidPhoneNumberList += "$($UPN) - $($IDMAuthPhone)"
        }
        Write-Log "$($UPN) ($($User.displayName)): Error configuring MFA phone: $($IDMAuthPhone) ($($operation)) - $($errObj.error.code)" -MessageType "ERROR" -ForceOnScreen
        # if the error is invalidPhoneNumber and we have a mobile number, try to use that instead
        if (($errObj.error.code -eq "invalidPhoneNumber") -and $mobile) {
          Write-Log "$($UPN) ($($User.displayName)): auth phone number invalid, fallback to mobile"
          $GraphBody = [pscustomobject]@{
            phoneNumber = $mobile
            phoneType = "mobile"
          } | ConvertTo-Json
          Try {
            #configure MFA number - mobile phone
            $ResponsePOST = Invoke-WebRequest -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $GraphBody -Method "POST" -ContentType $ContentTypeJSON
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
    


    $UriResource = "users/$($userId)/authentication/signInPreferences"
    $Uri = New-GraphUri -Version "beta" -Resource $UriResource
    Try {
      $ResponseGET = Invoke-WebRequest -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON
      $SignInPreferences = $ResponseGET | ConvertFrom-Json
      
      $sysPrefEnabled = $SignInPreferences.isSystemPreferredAuthenticationMethodEnabled
      $usrPrefMethod  = $SignInPreferences.userPreferredMethodForSecondaryAuthentication.ToLower()
      $sysPrefMethod  = $SignInPreferences.systemPreferredAuthenticationMethod.ToLower()

      If (-not ($sysPrefMethod -eq $UsrToSysMethodConv_DB[$usrPrefMethod])) {
          $targetMethod = $SysToUsrMethodConv_DB[$sysPrefMethod]
          #Write-Log "$($UPN): current userPreferredMethod: $($usrPrefMethod) - should be: $($sysPrefMethod) ($($targetMethod))" -MessageType "WARNING"
          
          $UriResource = "users/$($userId)/authentication/signInPreferences"
          $Uri = New-GraphUri -Version "beta" -Resource $UriResource
          $GraphBody = [pscustomobject]@{
              userPreferredMethodForSecondaryAuthentication = $targetMethod
          } | ConvertTo-Json
          Write-Host $GraphBody
          Write-Host $Uri
          Try {
              #configure preferred auth method
              $ResponsePATCH = Invoke-WebRequest -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $GraphBody -Method "PATCH" -ContentType $ContentTypeJSON
              Write-Log "$($UPN): userPreferredMethod set to: $($targetMethod), previous value: $($usrPrefMethod)"
              $signInPreferencesSetSuccessfully = $true
          }
          Catch {
              $errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
              Write-Log "$($UPN): Error configuring preferred auth method $($targetMethod): $($errObj.error.code)" -MessageType "ERROR" -ForceOnScreen
          }
      }
      else {
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

  Start-Sleep -Milliseconds $ThrottlingDelayPerUserinMsec
  $counter++
  if ($counter -eq 20) {
    break
  }
}#foreach ($User in $ADUsers)

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
