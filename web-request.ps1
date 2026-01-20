#######################################################################################################################
# Set-AAD-MFA-Phone-From-IDM
#######################################################################################################################

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1

#######################################################################################################################

$LogFolder          = "aad-mfa"
$LogFilePrefix      = "aad-mfa"
$LogFileSuffix      = "set-phone-from-IDM"
$LogFileFreq        = "YMD"

$OutputFolder       = "aad-mfa\reports"
$OutputFilePrefix	  = "aad-mfa"
$OutputFileSuffix	  = "set-phone-from-IDM"

#######################################################################################################################

. $ScriptPath\include-Var-Define.ps1
. $ScriptPath\include-Var-Init.ps1
. $ScriptPath\include-Functions-Common.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Suffix $LogFileSuffix -Ext "log"

$alternateMobilePhoneMethodId  = "b6332ec1-7057-4abe-9331-3d72feddfe41"
$officePhoneMethodId           = "e37fc753-ff3b-4958-9484-eaa9425c82bc"
$mobilePhoneMethodId           = "3179e48a-750b-4051-897c-87b9720928f7"

[array]$MFAPhoneReport = @()
[array]$ADUsers = $null
[hashtable]$AADUser_DB = @{}

#######################################################################################################################

$userId = "d0ffee24-32f0-455f-a99d-1b2c541ab97b"

Request-MSALToken -AppRegName "CEZ_AAD_MFA_config" -TTL 30
#$Uri = "https://graph.microsoft.com/beta/users/$($UPN)/authentication/phoneMethods/$($mobilePhoneMethodId)"
$UriResource = "users/$($userId)/authentication/phoneMethods/$($mobilePhoneMethodId)"
$Uri = New-GraphUri -Version "beta" -Resource $UriResource


$MobilePhoneMethod = Invoke-RestMethod -Headers $AuthDB["CEZ_AAD_MFA_config"].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON
write-host $mobilePhoneMethod

$Response = Invoke-WebRequest -Headers $AuthDB["CEZ_AAD_MFA_config"].AuthHeaders -Uri $Uri -Method "DELETE" -ContentType $ContentTypeJSON
write-host $Response.StatusCode
Get-Member -InputObject $Response
write-host $Response.StatusCode

Exit


if ((Get-Count -Object $ADUsers) -gt 0) {
  Initialize-ProgressBarMain -Activity "Processing MFA config for all users" -Total $ADUsers.Count
  foreach ($User in $ADUsers) {
    Request-MSALToken -AppRegName "CEZ_AAD_MFA_config" -TTL 30
    Update-ProgressBarMain

    $UPN = $User.UserPrincipalName
    If ((-not($DoNotConfigureFromIDM -contains($UPN))) -and $AADUsers.userPrincipalName -contains($UPN)) {
      $CurrentMFAPhone      = $null
      $CurrentMFAPhoneFixed = $null
      $operation            = $null
      $UpdateText           = $null
      $UPN = $User.userPrincipalName
      $userId = $AADUser_DB[$UPN].id
      $dspn = $User.DisplayName
      $sam  = $User.sAMAccountName
      $dn   = $User.distinguishedName
      $mobile = Get-MFAFormattedPhoneNumber -PhoneNumber $User.Mobile
      $OU = $dn.substring($dn.IndexOf('OU=')+3,($dn.substring(($dn.IndexOf('OU='))+3,30)).IndexOf('OU=')-1)
      $IDMAuthPhone = Get-MFAFormattedPhoneNumber -PhoneNumber $User.cEZIntuneMFAAuthMobile
      
      #$Uri = "https://graph.microsoft.com/beta/users/$($UPN)/authentication/phoneMethods/$($mobilePhoneMethodId)"
      $UriResource = "users/$($userId)/authentication/phoneMethods/$($mobilePhoneMethodId)"
      $Uri = New-GraphUri -Version "beta" -Resource $UriResource
      Try {
        $ErrorMessageGET = "success"
        $MobilePhoneMethod = Invoke-RestMethod -Headers $AuthDB["CEZ_AAD_MFA_config"].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON
        $CurrentMFAPhone = $MobilePhoneMethod.phoneNumber
        $CurrentMFAPhoneFixed = Get-MFAFormattedPhoneNumber -PhoneNumber $MobilePhoneMethod.phoneNumber
      }
      Catch {
        $ErrorMessageGET = $_.Exception.Message
        If (-not ($ErrorMessageGET.Contains("(404)"))) {
          Write-Log "$($UPN) - $($dspn) - $($sam): $($ErrorMessageGET)" -MessageType "ERROR" -ForceOnScreen
        }
      }
      If ($CurrentMFAPhone -eq $IDMAuthPhone) {
        $match = "OK"
        $clr = "Green"
      } 
      Else {
        $match = "DIFF"
        $clr = "Red"
      }
      write-host "$($User.UserPrincipalName.PadRight(35," ")) IDM:$($IDMAuthPhone.PadRight(20," ")) AAD-MFA:$($CurrentMFAPhone.PadRight(20," "))" -NoNewline
      write-host $match -ForegroundColor $clr
      
      If (($ErrorMessageGET.Contains("(404)")) -or ($CurrentMFAPhone -ne $IDMAuthPhone)) {
        If ($null -eq $CurrentMFAPhone) {
          #configure new MFA number
          $operation = "new"
        } 
        else {
          #update existing MFA number
          If ($CurrentMFAPhone -ne $CurrentMFAPhoneFixed) {
            #update - different MFA number
            $operation = "update-number"
          } 
          else {
            #update - current number OK but incorrect format
            $operation = "update-format"
          } 
          #current number needs to be deleted first
          Try {
            $ResultDELETE = Invoke-WebRequest -Headers $AuthDB["CEZ_AAD_MFA_config"].AuthHeaders -Uri $Uri -Method "DELETE" -ContentType $ContentTypeJSON
          }
          Catch {
            $ErrorMessageDELETE = $_.ErrorDetails.Message | Out-String
            Write-Log "$($dspn) ($($UPN) $($sam)): Error deleting MFA phone: $($CurrentMFAPhone): $($ErrorMessageDELETE)" -MessageType "ERROR" -ForceOnScreen
            $UpdateText = ": error deleting $($CurrentMFAPhone): $($ErrorMessageDELETE)"
          }
          If ($ResultDELETE.StatusCode -eq 204) {
            $UpdateText = ": $($CurrentMFAPhone) deleted"
          }        
        }
        
        #$Uri = "https://graph.microsoft.com/beta/users/$($UPN)/authentication/phoneMethods"
        $UriResource = "users/$($userId)/authentication/phoneMethods"
        $Uri = New-GraphUri -Version "beta" -Resource $UriResource
        $ErrorMessagePOST = "success"
        $GraphBody = [pscustomobject]@{
          phoneNumber = $IDMAuthPhone;
          phoneType = "mobile"
        } | ConvertTo-Json
        Try {
          #configure MFA number
          $ResultPOST = Invoke-RestMethod -Headers $AuthDB["CEZ_AAD_MFA_config"].AuthHeaders -Uri $Uri -Body $GraphBody -Method POST -ContentType $ContentTypeJSON
        }
        Catch {
          Write-Host $Uri
          Write-Host $AuthDB["CEZ_AAD_MFA_config"].AuthHeaders
          Write-Host $GraphBody
          Write-Host $_.ErrorDetails.Message

          $ErrorMessagePOST = $_.ErrorDetails.Message | Out-String
          Write-Log "$($dspn) ($($UPN) $($sam)): Error configuring MFA phone: $($IDMAuthPhone) ($($operation))$($UpdateText) - $($ErrorMessagePOST)" -MessageType "ERROR" -ForceOnScreen
        }
        If ($ResultPOST) {
          Write-Log "$($dspn) ($($UPN) $($sam)): MFA phone configured: $($IDMAuthPhone) ($($operation)$($UpdateText))"
        }  
      }  
      
      If ($CurrentMFAPhone -eq $IDMAuthPhone) {$operation = "ok-skip"}
          
      Try {
        $MFAPhoneReport += [pscustomobject]@{
          UserPrincipalName = $User.userPrincipalName;
          KPJM              = $User.sAMAccountName;
          DisplayName       = $User.displayName;
          Mail              = $User.mail;
          Mail_40           = $User.msExchExtensionAttribute40;
          OU                = $OU;
          Mobile            = $Mobile;
          CurrentMFAPhone   = $CurrentMFAPhone;
          IDMAuthPhone      = $IDMAuthPhone;
          Operation         = $operation;
          DN                = $dn
        }  
      }
      Catch {
        Write-Host $_.ErrorDetails.Message
      }
    }
    $UriResource = "users/$($userId)/authentication/signInPreferences"
    $Uri = New-GraphUri -Version "beta" -Resource $UriResource
    Try {
      $signInPreferences = Invoke-RestMethod -Headers $AuthDB["CEZ_AAD_MFA_config"].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON
    }
    Catch {
      $ErrorMessageGET = $_.Exception.Message
      If (-not ($ErrorMessageGET.Contains("(404)"))) {
        Write-Log "$($UPN) - $($dspn) - $($sam): $($ErrorMessageGET)" -MessageType "ERROR" -ForceOnScreen
      }
    }
  }
}
Complete-ProgressBarMain

Export-Report "MFA phone report" -Report $MFAPhoneReport -SortProperty "UserPrincipalName" -Path $OutputFile

#######################################################################################################################

. $ScriptPath\include-Script-EndLog-generic.ps1
