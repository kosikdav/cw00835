#######################################################################################################################
# Set-AAD-MFA-Phone-From-IDM
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
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

$ADCredentialPath = $aadauthmobmgmt_cred

$DoNotConfigureFromIDM = @()

[array]$MFAPhoneReport = @()
[array]$InvalidPhoneNumberList = @()
[array]$ADUsers = $null
[hashtable]$AADUser_DB = @{}
[int]$ThrottlingDelayPerUserinMsec = 300

#######################################################################################################################

. $IncFile_StdLogStartBlock

$ADCredential = Import-Clixml -Path $ADCredentialPath
Write-Log "AD credential file: $($ADCredentialPath)"

$ADFilter = "(sAMAccountName -notlike `"qh*`") -and (cEZIntuneMFAAuthMobile -like `"*`") -and (msExchExtensionAttribute29 -like `"*`") -and (msExchExtensionAttribute40 -like `"*`")"
$ADProperties = @(
  "userPrincipalName",
  "DisplayName",
  "sAMAccountName",
  "distinguishedName",
  "mobile",
  "mail",
  "cEZIntuneMFAAuthMobile",
  "msExchExtensionAttribute40"
)
$ADUsers = Get-ADUser -Credential $ADCredential -Filter $ADFilter -Properties $ADProperties
Write-Log "AD users with configured `"cEZIntuneMFAAuthMobile`" attribute: $(Get-Count -Object $ADUsers)"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriFilter = "UserType+eq+'Member'&accountEnabled+eq+'True'&onPremisesSyncEnabled+eq+'True'"
#$UriFilter = "startswith(userPrincipalName,'josef.mat')"
$UriSelect = "id,userPrincipalName,onPremisesSyncEnabled"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Filter $UriFilter -Select $UriSelect
$AADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ProgressDots -Text "AAD member users"

foreach ($user in $AADUsers) {
  $UserObject = [PSCustomObject]@{
    id = $user.id;
    userPrincipalName = $user.userPrincipalName
  }
  Try {
    $AADUser_DB.Add($user.userPrincipalName,$UserObject)
  }
  Catch {
    Write-Host $UserObject
  }
}

#######################################################################################################################


foreach ($User in $ADUsers) {
  Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
  $UPN = $User.UserPrincipalName
  If ($AADUsers.userPrincipalName -contains($UPN)) {
    $CurrentMFAPhone = $CurrentMFAPhoneFixed = $operation = $sysPrefEnabled = $usrPrefMethod = $sysPrefMethod = $targetMethod = $null
    $userId = $AADUser_DB[$UPN].id
    $dspn = $User.DisplayName
    $sam  = $User.sAMAccountName
    $dn   = $User.distinguishedName
    $mobile = Get-MFAFormattedPhoneNumber -PhoneNumber $User.Mobile
    $OU = $dn.substring($dn.IndexOf('OU=')+3,($dn.substring(($dn.IndexOf('OU='))+3,30)).IndexOf('OU=')-1)
    $IDMAuthPhone = Get-MFAFormattedPhoneNumber -PhoneNumber $User.cEZIntuneMFAAuthMobile
    
    $UriResource = "users/$($userId)/authentication/phoneMethods/$($mobilePhoneMethodId)"
    $Uri = New-GraphUri -Version "beta" -Resource $UriResource
    Try {
      $ErrorMessageGET = "success"
      $ResponseGET = Invoke-WebRequest -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON -UseBasicParsing
      $MobilePhoneMethod = $ResponseGET | ConvertFrom-Json
      $CurrentMFAPhone = $MobilePhoneMethod.phoneNumber
      $CurrentMFAPhoneFixed = Get-MFAFormattedPhoneNumber -PhoneNumber $MobilePhoneMethod.phoneNumber
    }
    Catch {
      $ErrorMessageGET = $_.Exception.Message
      If ($ErrorMessageGET.Contains("(404)")) {
        $CurrentMFAPhone = "empty"
        $CurrentMFAPhoneFixed = "empty"
      }
      Else {
        Write-Log "$($UPN) ($($dspn)): Error reading MFA phoneMethods: $($ErrorMessageGET)" -MessageType "ERROR" -ForceOnScreen
        Continue
      }
    }
    
    If ($CurrentMFAPhone -eq $IDMAuthPhone) {
      # numbers in IDM and AAD MFA match, nothing to do
      $operation = "ok-skip"
      $match = "OK"
      $clr = "Green"
    } 
    Else {
      # numbers do not match
      $match = "DIFF"
      $clr = "Red"
    }
    
    write-host "$($User.UserPrincipalName.PadRight(35," ")) IDM:$($IDMAuthPhone.PadRight(20," ")) AAD-MFA:$($CurrentMFAPhone) " -NoNewline
    write-host $match -ForegroundColor $clr
    
    If (($ErrorMessageGET.Contains("(404)")) -or ($CurrentMFAPhone -ne $IDMAuthPhone)) {
      If ($CurrentMFAPhone -eq "empty") {
        #configure new MFA number
        $operation = "new"
      } 
      else {
        #update existing MFA number
        If ($IDMAuthPhone -ne $CurrentMFAPhoneFixed) {
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
          Write-Log "$($UPN) ($($dspn)): MFA phone $($CurrentMFAPhone) deleted"
        }
        Catch {
          $errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
          Write-Log "$($UPN) ($($dspn)): Error deleting MFA phone $($CurrentMFAPhone): $($errObj.error.code)" -MessageType "ERROR" -ForceOnScreen
          Continue
        }
      }
      
      $UriResource = "users/$($userId)/authentication/phoneMethods"
      $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
      $GraphBody = [pscustomobject]@{
        phoneNumber = $IDMAuthPhone;
        phoneType = "mobile"
      } | ConvertTo-Json
      Try {
        #configure MFA number - auth phone
        $ResponsePOST = Invoke-WebRequest -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $GraphBody -Method "POST" -ContentType $ContentTypeJSON
        Write-Log "$($UPN) ($($dspn)): MFA phone configured: $($IDMAuthPhone) ($($operation))"
      }
      Catch {
        $errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
        if ($errObj.error.code -eq "invalidPhoneNumber") {
          $InvalidPhoneNumberList += "$($UPN) - $($IDMAuthPhone)"
        }
        Write-Log "$($UPN) ($($dspn)): Error configuring MFA phone: $($IDMAuthPhone) ($($operation)) - $($errObj.error.code)" -MessageType "ERROR" -ForceOnScreen
        # if the error is invalidPhoneNumber and we have a mobile number, try to use that instead
        if (($errObj.error.code -eq "invalidPhoneNumber") -and $mobile) {
          Write-Log "$($UPN) ($($dspn)): auth phone number invalid, fallback to mobile"
          $GraphBody = [pscustomobject]@{
            phoneNumber = $mobile
            phoneType = "mobile"
          } | ConvertTo-Json
          Try {
            #configure MFA number - mobile phone
            $ResponsePOST = Invoke-WebRequest -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $GraphBody -Method "POST" -ContentType $ContentTypeJSON
            Write-Log "$($UPN) ($($dspn)): MFA phone configured: $($IDMAuthPhone) ($($operation))"
          }
          Catch {
            $errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
            Write-Log "$($UPN) ($($dspn)): Error configuring MFA phone: $($IDMAuthPhone) ($($operation)) - $($errObj.error.code)" -MessageType "ERROR" -ForceOnScreen
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
          }
          Catch {
              $errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
              Write-Log "$($UPN): Error configuring preferred auth method $($targetMethod): $($errObj.error.code)" -MessageType "ERROR" -ForceOnScreen
          }
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
      UserPrincipalName = $UPN;
      Id                = $userId;
      KPJM              = $sam;
      DisplayName       = $dspn;
      Mail              = $User.mail;
      Mail_40           = $User.msExchExtensionAttribute40;
      OU                = $OU;
      mobile            = $mobile;
      CurrentMFAPhone   = $CurrentMFAPhone;
      IDMAuthPhone      = $IDMAuthPhone;
      Operation         = $operation;
      DN                = $dn;
      sysPrefEnabled    = $sysPrefEnabled;
      usrPrefMethod     = $usrPrefMethod;
      sysPrefMethod     = $sysPrefMethod;
      targetMethod      = $targetMethod
    }
  
    Start-Sleep -Milliseconds $ThrottlingDelayPerUserinMsec
  }#If ($AADUsers.userPrincipalName -contains($UPN))
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
Export-Report "MFA phone report" -Report $MFAPhoneReport -SortProperty "UserPrincipalName" -Path $OutputFile

#######################################################################################################################

. $IncFile_StdLogEndBlock
