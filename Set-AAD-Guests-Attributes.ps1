
# Set-AAD-Guests-Attributes
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "aad-guest-mgmt"
$LogFilePrefix		= "aad-guests-attributes"
$daysBackOffset     = 30

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$timeDiffTolerance  = 60
$sleepShort = 60
$sleepLong = 120

[array]$DoNotFixGuestList = @()
[hashtable]$AuditLogEvents_DB = @{}
[array]$ConflictingProxyAddresses = @()
$ErrMsgProxyAddrConflict = "is already being used by the proxy addresses or LegacyExchangeDN"

#######################################################################################################################

. $IncFile_StdLogStartBlock

Write-Log "daysBackOffset: $($daysBackOffset) days"
Write-Log "audit log events start date: $($strYesterdayUTCStart)"
Write-Log "timeDiffTolerance: $($timeDiffTolerance) seconds"

$AADEXTTenant_DB = Import-CSVtoHashDB -Path $DBFileExtAADTenants -KeyName "domain"
$AADPartnerTenant_DB = Import-CSVtoHashDB -Path $DBFileAADPartnerTenants -KeyName "tenantId"

$guestsFixed = 0
$XTSyncCounter = 0

##############################################################################################
# read RecentAuditLogEventsInvite from Graph
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "auditLogs/directoryAudits"
$UriFilter = "activityDisplayName+eq+'Invite external user'+and+result+eq+'success'+and+activityDateTime+ge+$($strYesterdayUTCStart)"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Filter $UriFilter
$RecentAuditLogEventsInvite = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ProgressDots -Text "Audit log events"

##############################################################################################
# read Guests from Graph 
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriFilter = "UserType+eq+'Guest'"
$UriSelect = "id,userPrincipalName,userType,displayName,createdDateTime,mail,companyName,employeeType,employeeHireDate,otherMails,proxyAddresses,showInAddressList,onPremisesExtensionAttributes"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Filter $UriFilter -Select $UriSelect
[array]$AllAADGuests = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ProgressDots -Text "AAD guest users"

foreach ($AuditLogEvent in $RecentAuditLogEventsInvite) {
    $ownerUpn = $AuditLogEvent.initiatedBy.user.userPrincipalName
    if ($ownerUpn) {
        if ($AuditLogEvent.Result -eq "success" -and ($AuditLogEvent.additionalDetails.Count -ge 1)) {
            $InvitedUserMail = $($AuditLogEvent.additionalDetails | Where-Object { $_.key -eq "invitedUserEmailAddress" })[0].value
            if ($InvitedUserMail) {
                $UriResource = "users/$($ownerUpn)"
                $UriSelect = "id,userPrincipalName,onPremisesSamAccountName"
                $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
                $owner = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
                $AuditRecord = [pscustomobject]@{
                    activityDateTime    = $AuditLogEvent.activityDateTime
                    result              = $AuditLogEvent.Result
                    ownerUpn            = $ownerUpn
                    ownerSamAccountName = $owner.onPremisesSamAccountName
                    invitedUserMail     = $InvitedUserMail
                }
                if ($AuditLogEvents_DB.Contains($InvitedUserMail)) {
                    $AuditLogEvents_DB.Set_Item($InvitedUserMail,$AuditRecord)
                }
                else {
                    $AuditLogEvents_DB.Add($InvitedUserMail,$AuditRecord)
                }
            }
        }
	}
}
Remove-Variable RecentAuditLogEventsInvite

##############################################################################################
# Process all guest accounts


foreach ($Guest in $AllAADGuests) {
    Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 60
    Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
    $MailDomain = $ExtTenant = $PartnerTenant = $AADExtCompanyName = $CurrentCompanyName = $CurrentEmployeeType = $MailUser = $null
    $InboundSync = $XTSync = $false
    $upn = $Guest.UserPrincipalName
    $ext15 = $Guest.onPremisesExtensionAttributes.extensionAttribute15
    $UriResource = "users/$($Guest.id)"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
    #check for cross tenant sync accounts - ext15 attr = XTSync_tenantId
    if ($ext15 -and $ext15.StartsWith("XTSync_")) {
        $XTSync = $true
        $XTSyncCounter++
    }
    if ($Guest.Mail) {
        $MailDomain	= ($Guest.Mail.Split("@")[1]).ToLower()
        $ExtTenant = $AADExtTenant_DB[$MailDomain]
        If ($ExtTenant) {
            $AADExtCompanyName = $ExtTenant.displayName.Trim()
            $AADExtCompanyName = $AADExtCompanyName.substring(0, [System.Math]::Min(63, $AADExtCompanyName.Length))
            if ($T2TTenant_DB.ContainsKey($ExtTenant.tenantId)) {
                $AADExtCompanyName = $T2TTenant_DB[$ExtTenant.tenantId]
            }
            If ($AADPartnerTenant_DB.ContainsKey($ExtTenant.tenantId)) {
                $PartnerTenant = $AADPartnerTenant_DB[$ExtTenant.tenantId]
                $InboundSync = Convert-ValueToBool -Value $PartnerTenant.InboundSyncAllowed
            }
        }
    }
    
    if ($XTSync) {
        #guest is synced via T2T sync - extensionAttribute15 is set
        $tenantId = ($ext15 -Split "_", 2)[1]
        #fix possible _ to - in tenantId - ELENG :)
        $tenantId = $tenantId -replace "_","-"
        if ($ExtTenant -and ($ExtTenant.tenantId -ne $tenantId)) {
            write-log "$($Guest.mail) - tenantId mismatch $($ExtTenant.tenantId) vs $($tenantId)" -ForegroundColor "Red"
            Continue
        }

        if ($ExtTenant -and $T2TTenant_DB.ContainsKey($ExtTenant.tenantId)) {
            $T2TTenantOrgName = $T2TTenant_DB[$ExtTenant.tenantId]
            #if companyName is set and does not match the T2T tenant name, update it
            if ($Guest.companyName -and ($Guest.companyName.Trim() -ne $T2TTenantOrgName)) {
                write-host "$($Guest.mail) - $($T2TTenantOrgName)"
                $UriResource = "users/$($Guest.id)"
                $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
                $GraphBodyCompany = @{
                    companyName = $T2TTenantOrgName
                } | ConvertTo-Json
                $GraphBodyCompany = [System.Text.Encoding]::UTF8.GetBytes($GraphBodyCompany)
                Try {
                    $ResultPATCH = Invoke-RestMethod -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $GraphBodyCompany -Method "PATCH" -ContentType $ContentTypeJSON
                    Write-Log "$($Guest.mail) SUCCESS companyName: `"$($T2TTenantOrgName)`"" -ForegroundColor "Cyan"
                }
                Catch {
                    $ErrorMessagePATCH = $_.ErrorDetails.Message | Out-String
                    Write-Log "$($Guest.mail) ERR PATCH companyName: `"$($T2TTenantOrgName)`"" -MessageType Error
                    Write-Log $ErrorMessagePATCH -MessageType Error
                }
            }
            #if synced as member, change to guest
            if ($Guest.userType -eq "Member") {
                write-host "$($Guest.mail) - $($Guest.userType)"
                $GraphBodyCompany = @{
                    userType = "Guest"
                } | ConvertTo-Json
                $GraphBodyCompany = [System.Text.Encoding]::UTF8.GetBytes($GraphBodyCompany)
                Try {
                    $ResultPATCH = Invoke-RestMethod -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $GraphBodyCompany -Method "PATCH" -ContentType $ContentTypeJSON
                    Write-Log "$($Guest.mail) - $($Guest.userType) SUCCESS userType PATCH to guest" -ForegroundColor "Cyan"
                }
                Catch {
                    $ErrorMessagePATCH = $_.ErrorDetails.Message | Out-String
                    Write-Log "ERR PATCH userType to guest" -MessageType Error
                    Write-Log $($ErrorMessagePATCH) -MessageType Error
                }
            }
            #if showInAddressList is not set, set it to true
            $MailUser = Get-MailUser -Identity $Guest.Mail
            if ($MailUser.HiddenFromAddressListsEnabled) {
                Try {
                    Set-MailUser -Identity $Guest.Mail -HiddenFromAddressListsEnabled:$false
                    Write-Log "$($Guest.mail) SUCCESS showInAddressList: (mail domain: $($MailDomain) user object age: $($GuestAge.Days))" -ForegroundColor "Cyan"
                }
                Catch {
                    Write-Log "$($Guest.mail) ERR showInAddressList: (mail domain: $($MailDomain) user object age: $($GuestAge.Days))" -MessageType Error
                    Write-Log $_.Exception.Message -MessageType Error
                }
            }
        }
        else {
            Write-Log "$($Guest.mail) ERR ext15 attribute set but T2TTenant_DB does not contain $($ExtTenant.tenantId)" -MessageType Error
        }
    }
    else {
        #guest was manually created via invite
        $now = Get-Date
        $GuestAge = New-TimeSpan $Guest.createdDateTime $now
    
        If ($Guest.companyName) {
            $CurrentCompanyName = $Guest.companyName.Trim()
        }
        
        If ($Guest.employeeType) {
            $CurrentEmployeeType = $Guest.employeeType.Trim()
        }

        ##############################################################################################
        # AuditTrace - employeeHireDate attribute
        if (($null -eq $Guest.employeeHireDate) -and ($GuestAge.Days -le 180)) {
            #set to createdDateTime if empty and user is less than 180 days old
            $GraphBody = @{
                employeeHireDate = $Guest.createdDateTime
            } | ConvertTo-Json
            #write-host $GraphBody
            Try {
                $ResultPATCH = Invoke-RestMethod -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $GraphBody -Method "PATCH" -ContentType $ContentTypeJSON
                Write-Log "$($Guest.mail) SUCCESS employeeHireDate: `"$($Guest.createdDateTime)`"" -ForegroundColor "Green"
            }
            Catch {
                $ErrorMessagePATCH = $_.ErrorDetails.Message | Out-String
                Write-Log "$($Guest.mail) ERR PATCH employeeHireDate: `"$($Guest.createdDateTime)`"" -MessageType Error
                Write-Log $_.Exception.Message -MessageType Error
                Write-Log $ErrorMessagePATCH -MessageType Error
            }
        }

        ##############################################################################################
        # showInAddressList attribute
        if (($GuestAge.Days -gt 1) -and ($AADGuestDomainsShowInGAL.Contains($MailDomain) -or $InboundSync)) {
            $MailUser = Get-MailUser -Identity $Guest.Mail
            if ($MailUser.HiddenFromAddressListsEnabled) {
                Try {
                    Set-MailUser -Identity $Guest.Mail -HiddenFromAddressListsEnabled:$false
                    Write-Log "$($Guest.mail) SUCCESS showInAddressList: (mail domain: $($MailDomain) user object age: $($GuestAge.Days))" -ForegroundColor "Cyan"
                }
                Catch {
                    Write-Log "$($Guest.mail) ERR showInAddressList: (mail domain: $($MailDomain) user object age: $($GuestAge.Days))" -MessageType Error
                    Write-Log $_.Exception.Message -MessageType Error
                }
            }
        }
    
        ##############################################################################################
        # employeeType attribute 
        If ($AuditLogEvents_DB.Contains($Guest.mail)) {
            $AuditLogDBRecord = $AuditLogEvents_DB[$Guest.mail]
            if ($AuditLogDBRecord.ownerUpn.Length -gt 35) {
                if ($AuditLogDBRecord.ownerSamAccountName) {
                    $stampString = $AuditLogDBRecord.ownerSamAccountName + ";" + $AuditLogDBRecord.activityDateTime
                }
                else {
                    $stampString = $AuditLogDBRecord.ownerUpn.Substring(0, [System.Math]::Min(35, $my_string.Length)) + ";" + $AuditLogDBRecord.activityDateTime
                }
            }
            else {
                $stampString = $AuditLogDBRecord.ownerUpn + ";" + $AuditLogDBRecord.activityDateTime
            }
           
            If (-not $CurrentEmployeeType) {
                $diff = New-TimeSpan $AuditLogDBRecord.activityDateTime $Guest.createdDateTime
                if ([Math]::Abs($diff.Seconds) -le $timeDiffTolerance) {
                    $GraphBodyEmployeeType = @{
                        employeeType = $stampString
                    } | ConvertTo-Json
                    Try {
                        $ResultPATCH = Invoke-RestMethod -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $GraphBodyEmployeeType -Method "PATCH" -ContentType $ContentTypeJSON
                        Write-Log "$($Guest.mail) SUCCESS employeeType: `"$($stampString)`" (diff:$($diff.Seconds))" -ForegroundColor "Green"
                    }
                    Catch {
                        $ErrorMessagePATCH = $_.ErrorDetails.Message | Out-String
                        Write-Log "$($Guest.mail) ERR PATCH employeeType: `"$($stampString)`"" -MessageType Error
                        Write-Log $ErrorMessagePATCH -MessageType Error
                        write-host $_.Exception.Message
                    }
                }
                else {
                    Write-Log "$($Guest.mail) ERR TIME DIFF TOO LARGE: Invite:$($AuditLogDBRecord.activityDateTime) Created:$($Guest.createdDateTime) Diff:$($diff.Seconds)  (max allowed:$($timeDiffTolerance) sec" -MessageType Error
                }
            }
        }
    
        ##############################################################################################
        # companyName attribute 
        if ($MailDomain -and $AADExtTenant_DB.ContainsKey($MailDomain) -and ($CurrentCompanyName -ne $AADExtCompanyName)) {
            $GraphBodyCompany = @{
                companyName = $AADExtCompanyName.Trim()
            } | ConvertTo-Json
            $GraphBodyCompany = [System.Text.Encoding]::UTF8.GetBytes($GraphBodyCompany)
            Try {
                $ResultPATCH = Invoke-RestMethod -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $GraphBodyCompany -Method "PATCH" -ContentType $ContentTypeJSON
                Write-Log "$($Guest.mail) SUCCESS companyName: `"$($AADExtCompanyName)`" (previous value: `"$($CurrentCompanyName)`")" -ForegroundColor "Cyan"
            }
            Catch {
                $ErrorMessagePATCH = $_.ErrorDetails.Message | Out-String
                Write-Log "$($Guest.mail) ERR PATCH companyName: `"$($AADExtCompanyName)`" (current value: `"$($CurrentCompanyName)`")" -MessageType Error
                Write-Log $ErrorMessagePATCH -MessageType Error
                write-host $_.Exception.Message
            }
        }
    
        ##############################################################################################
        # proxyAddresses attribute
        if (-not($DoNotFixGuestList -Contains ($upn))) {
            $UPNExtMail = Get-MailFromGuestUPN -GuestUPN $upn
            if ($UPNExtMail) {
                if ($Guest.proxyAddresses -Contains(("smtp:"+$upn))) {
                    Write-Log "$($upn) - expected mail: $($UPNExtMail)"
                    if ($Guest.proxyAddresses.Count -gt 1) {
                        foreach ($proxyAddress in $Guest.proxyAddresses) {
                            if ($proxyAddress.StartsWith("SMTP:")) {
                                continue
                            }
                            else {
                                if (($proxyAddress -ne "smtp"+$UPNExtMail) -and ($proxyAddress -ne "smtp"+$upn)) {
                                    Write-Log "$($upn) - removing redundant proxyAddress $($proxyAddress)"
                                    Try {
                                        Set-MailUser $upn -EmailAddresses @{remove="$($proxyAddress)"} -ErrorAction Stop
                                        Start-SleepDots -Seconds $sleepShort
                                    }
                                    Catch {
                                        Write-Log $_.Exception.Message -MessageType Error
                                    }
                                }
                            }
                        }
                    }
    
                    if (-not($Guest.proxyAddresses -Contains(("smtp:"+$UPNExtMail)))) {
                        Write-Log "$($upn) - $($UPNExtMail) missing from proxyAddresses, adding first"
                        Try {
                            Set-MailUser $upn -EmailAddresses @{add="smtp:$($UPNExtMail)"} -ErrorAction Stop
                            Start-SleepDots -Seconds $sleepShort
                        }
                        Catch {
                            Write-Log $_.Exception.Message -MessageType Error
                            if ($_.Exception.Message -like "*$($ErrMsgProxyAddrConflict)*") {
                                $ConflictingProxyAddresses += $UPNExtMail
                            }
                            else {
                                Continue	
                            }
                            Continue
                        }
                    }
                    else {
                        Write-Log "$($upn) - $($UPNExtMail) already in proxyAddresses"
                    }
        
                    if (-not($Guest.proxyAddresses -CContains(("SMTP:"+$UPNExtMail)))) {
                        Write-Log "$($upn) - setting primary SMTP addr to $($UPNExtMail)"
                        Try {
                            Set-MailUser $upn -PrimarySmtpAddress $UPNExtMail -ErrorAction Stop
                            Start-SleepDots -Seconds $sleepLong
                        }
                        Catch {
                            Write-Log $_.Exception.Message -MessageType Error
                            Continue	
                        }
                    }
                    else {
                        Write-Log "$($upn) - $($UPNExtMail) already primary SMTP addr"
                    }
                    Write-Log "$($upn) - deleting $($upn) from proxyAddresses"
                    Try {
                        Set-MailUser $upn -EmailAddresses @{remove="$($upn)"} -ErrorAction Stop
                        $guestsFixed++
                    }
                    catch {
                        Write-Log $_.Exception.Message -MessageType Error
                        Continue	
                    }
                }
            }
            else {
                Write-Log "$($upn) - bad UPN format: $($upn)"
                continue
            }
            continue
        }
        else {
            Write-Log -LogString "$($upn) - ignoring $($upn) based on exception list" 
        }    
    }
}
if ($ConflictingProxyAddresses.Count -gt 0) {
    Write-Log $string_divider
    Write-Log "Conflicting proxy addresses found: $($ConflictingProxyAddresses.Count)" -MessageType Warn
    foreach ($address in $ConflictingProxyAddresses) {
        Write-Log $address -MessageType Warn
    }
    Write-Log $string_divider
}
else {
    Write-Log "No conflicting proxy addresses found."
}
Write-Log "Guest accounts processed: $($AllAADGuests.Count)"
Write-Log "XTSync accounts processed: $($XTSyncCounter)"
Write-Log "Guests fixed proxyAddresses: $($guestsFixed)"
#######################################################################################################################

. $IncFile_StdLogEndBlock
