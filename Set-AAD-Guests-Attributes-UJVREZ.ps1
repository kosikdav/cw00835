
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
$LogFilePrefix		= "aad-guests-attributes-ujvrez"
$daysBackOffset     = 30

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

#######################################################################################################################

. $IncFile_StdLogStartBlock

##############################################################################################
# read Guests from Graph 
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriFilter = "UserType eq 'Guest' and endsWith(mail,'@ujvgroup.cz')"
$UriSelect = "id,userPrincipalName,userType,displayName,mail,companyName,otherMails,proxyAddresses,onPremisesExtensionAttributes"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Filter $UriFilter -Select $UriSelect

[array]$AllAADGuests = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ConsistencyLevel "eventual"
write-host "AAD Guests retrieved: $($AllAADGuests.Count)" -ForegroundColor "Green"

##############################################################################################
# Process all guest accounts

foreach ($Guest in $AllAADGuests) {
    Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 60
    Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
    write-host "Processing guest $($Guest.mail) - $($Guest.displayName)" -ForegroundColor "Cyan"
    $MailDomain = $ExtTenant = $PartnerTenant = $AADExtCompanyName = $CurrentCompanyName = $CurrentEmployeeType = $MailUser = $null
    $InboundSync = $XTSync = $false
    $upn = $Guest.UserPrincipalName
    $ext15 = $Guest.onPremisesExtensionAttributes.extensionAttribute15
    $UriResource = "users/$($Guest.id)"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
    #check for cross tenant sync accounts - ext15 attr = XTSync_tenantId
    if ($ext15 -and $ext15.StartsWith("XTSync_")) {
        $XTSync = $true
    }
    if ($Guest.Mail) {
        $MailDomain	= ($Guest.Mail.Split("@")[1]).ToLower()
    }
    
    if ($XTSync) {
        #guest is synced via T2T sync - extensionAttribute15 is set
        $tenantId = ($ext15 -Split "_", 2)[1]
        #fix possible _ to - in tenantId - ELENG :)
        $tenantId = $tenantId -replace "_","-"

        if ($tenantId -eq $TenantId_UJVREZ) {
            if ($Guest.proxyAddresses) {
                foreach ($proxyAddress in $Guest.proxyAddresses) {
                    if (($proxyAddress.StartsWith("SMTP:")) -or ($proxyAddress -like "*@ujvgroup.cz")) {
                        continue
                    }
                    else {
                        Try {
                            Set-MailUser $upn -EmailAddresses @{remove="$($proxyAddress)"} -ErrorAction Stop
                            write-log "$($Guest.mail) SUCCESS removing redundant proxyAddress $($proxyAddress)" -ForegroundColor "Yellow"
                            #Start-SleepDots -Seconds $sleepShort
                        }
                        Catch {
                            Write-Log $_.Exception.Message -MessageType Error
                        }
                    }
                }
            }

            if ($Guest.otherMails) {
                $GraphBodyOther = @{
                    otherMails = @()
                } | ConvertTo-Json
                $GraphBodyOther = [System.Text.Encoding]::UTF8.GetBytes($GraphBodyOther)
                Try {
                    $ResultPATCH = Invoke-RestMethod -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $GraphBodyOther -Method "PATCH" -ContentType $ContentTypeJSON
                    Write-Log "$($Guest.mail) SUCCESS deleting otherMails" -ForegroundColor "Cyan"
                }
                Catch {
                    $ErrorMessagePATCH = $_.ErrorDetails.Message | Out-String
                    Write-Log "$($Guest.mail) ERR PATCH otherMails" -MessageType Error
                    Write-Log $ErrorMessagePATCH -MessageType Error
                    write-host $_.Exception.Message
                }
            }
            Set-MailUser -Identity $Guest.Mail -HiddenFromAddressListsEnabled:$true
        }
    }
}

#######################################################################################################################

. $IncFile_StdLogEndBlock
