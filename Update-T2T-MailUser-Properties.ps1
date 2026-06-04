#######################################################################################################################
# Set-MaiboxT2T-Properties
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1

#######################################################################################################################

$LogFolder			= "t2t"
$LogFilePrefix		= "update-mbx-properties"
$LogFileFreq		= "Y"

#######################################################################################################################

. $ScriptPath\include-Script-Start-Include.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Freq $LogFileFreq -Ext "log"

[datetime]$date = (get-date).AddDays(-$DaysBack)

$Src_PN_attr = "extension_93b54ce056df45bd8f5f398753fa17c0_employeeNumber"
$Dst_PN_attr = "employeeNumber"
$Dst_Mapping_attr = "extensionAttribute10"

$b2bmailusers = "D:\data\t2t-ujvrez\ujvrez-guests-mailusers.csv"

$MbxFilter1 = "(alias -like '*')"
$MbxFilter2 = " -and ((RecipientTypeDetails -eq 'UserMailbox')"
$MbxFilter3 = " -or (RecipientTypeDetails -eq 'SharedMailbox')"
$MbxFilter4 = " -or (RecipientTypeDetails -eq 'RoomMailbox')"
$MbxFilter5 = " -or (RecipientTypeDetails -eq 'EquipmentMailbox'))"
$userMbxFilter = $MbxFilter1 + $MbxFilter2 + $MbxFilter3 + $MbxFilter4 + $MbxFilter5

#######################################################################################################################
function Get-TargetT2TUser {
	[CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$SrcIdentity,
        [string]$SrcAccessToken,
		[string]$DstAccessToken,
		[pscredential]$DstADCredential
    )
	# main function body ##################################
	$UriResource = "users/$SrcIdentity"
	$UriSelect = "id,userPrincipalName,onpremisesSamAccountName,$($Src_PN_attr)"
	$Uri = New-GraphUri -Resource $UriResource -Select $UriSelect -Version "v1.0"
	$SrcUser = Get-GraphOutputREST -Uri $Uri -AccessToken $SrcAccessToken -ContentType $ContentTypeJSON
	if ($SrcUser) {
		$Filter = "$($Dst_Mapping_attr) -eq '$($SrcUser.$Src_PN_attr)'"
		$DstUser = Get-ADUser -Filter $Filter -Properties $Dst_Mapping_attr -Credential $DstADCredential
		if ($DstUser) {
			return $DstUser
		}
		else {
			return $null
		}
	}
	else {
		return $null
	}
}
#######################################################################################################################

. $IncFile_StdLogStartBlock


$B2BMailUsers_DB = Import-CSVtoHashDB -Path $b2bmailusers -KeyName "PrimarySmtpAddress"

if ($InteractiveRun) {
	$ADCredentialPath = "c:\cred\qp_aad_grp_mgmt\qp_aad_grp_mgmt_qskosikdav.cred"
}
else {
	$ADCredentialPath = $aad_grp_mgmt_cred
}
$ADCredential = Import-Clixml -Path $ADCredentialPath

$DstMailRoutingDomain = "cezdata.mail.onmicrosoft.com"

$Src_AppReg_LOG_READER 			= $AppReg_UJVREZ_LOG_READER
$Src_AppReg_EXO_MGMT 			= $AppReg_UJVREZ_EXO_MGMT   
$Src_T2T_EXO_MIGRATION_GROUP 	= $UJVREZ_T2T_EXO_MIGRATION_GROUP

Request-MSALToken -AppRegName $Src_AppReg_LOG_READER -TTL 30
Connect-EXOService -AppRegName $Src_AppReg_EXO_MGMT  -TTL 120
$ExchangeSession = New-PSSession -Name "OnPremExchange" -ConfigurationName "Microsoft.Exchange" -ConnectionUri "http://cw00616exch3.cezdata.corp/PowerShell/" -Authentication Kerberos
Import-PSSession $ExchangeSession -DisableNameChecking -AllowClobber

#get source T2T migration group
$T2TMigrationGroupName = Get-GroupNameFromGraphById -AccessToken $AuthDB[$Src_AppReg_LOG_READER].AccessToken -Id $Src_T2T_EXO_MIGRATION_GROUP

if ($T2TMigrationGroupName) {
	Write-Host "Source T2T migration group found: $T2TMigrationGroupName"
} else {
	Write-Host "Source T2T migration group with ID $Src_T2T_EXO_MIGRATION_GROUP not found."
	Exit
}

$UriResource = "groups/$($Src_T2T_EXO_MIGRATION_GROUP)/members"
$Uri = New-GraphUri -Resource $UriResource -Version "v1.0"
[array]$T2TMigrationGroupMembers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$Src_AppReg_LOG_READER].AccessToken
write-host "SRC T2T migration group members: $($T2TMigrationGroupMembers.count)"

write-host "SRC mailboxes total: " -NoNewline
[array]$SrcMailboxesAll = Get-EXOMailbox -ResultSize Unlimited -PropertySets All
write-host $SrcMailboxesAll.count
[array]$SrcMailboxes = $SrcMailboxesAll | Where-Object { $_.ExternalDirectoryObjectId -in $T2TMigrationGroupMembers.id }
write-host "SRC mailboxes in T2T migration group: $($SrcMailboxes.count)"
Remove-Variable SrcMailboxesAll

<#
Request-MSALToken -AppRegName $Dst_AppReg_LOG_READER -TTL 30

$UriResource = "users"
$UriSelect = "id,userPrincipalName,onpremisesSamAccountName"
$UriFilter = "userType eq 'Member' and onpremisesSyncEnabled eq true"
$Uri = New-GraphUri -Resource $UriResource -Version "v1.0" -Select $UriSelect -Filter $UriFilter -Top 999
[array]$DstAADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$Dst_AppReg_LOG_READER].AccessToken -ProgressDots -Text "AAD users"
write-host "DST AAD users: $($DstAADUsers.count)"

Connect-EXOService -AppRegName $Dst_AppReg_EXO_MGMT -TTL 120
$DstEXOMailUsers = Get-EXORecipient -ResultSize Unlimited -PropertySets "Minimum, MailboxMove" -RecipientType MailUser
write-host "DST EXO mailUsers: $($DstEXOMailUsers.count)"
Get-PSSession | Remove-PSSession

$Session = New-PSSession -Name "OnPremExchange" -ConfigurationName "Microsoft.Exchange" -ConnectionUri "http://cw00616exch3.cezdata.corp/PowerShell/" -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking -AllowClobber | Out-Null
write-host "DST AD mailUsers: " -NoNewline
$DstADMailUsers = Get-MailUser -ResultSize Unlimited
write-host $DstADMailUsers.count
#>

foreach ($SrcMailbox in $SrcMailboxes) {
	#write-host "$($SrcMailbox.PrimarySmtpAddress)" -ForegroundColor Green
	[array]$SRCsmtpAddresses = [array]$SRCx500Addresses = [array]$SRCproxyAddresses = @()
	$ArchiveGuid = $null
	
	if ($SrcMailbox.EmailAddresses) {
        foreach ($EmailAddress in $SrcMailbox.EmailAddresses) {
            if ($EmailAddress -like "smtp:*" -and $EmailAddress -notlike "*onmicrosoft.com") {
                $SRCsmtpAddresses += $EmailAddress
            }
            if ($EmailAddress -like "X500:*") {
                $SRCx500Addresses += $EmailAddress
            }
        }
        $SRCproxyAddresses = $SrcMailbox.EmailAddresses -join ";"
    }
    if ($SrcMailbox.ArchiveGuid -and $SrcMailbox.ArchiveGuid -ne "00000000-0000-0000-0000-000000000000") {
        $ArchiveGuid = $SrcMailbox.ArchiveGuid
    }

	#Write-Host "Processing mailbox: $($SrcMailbox.UserPrincipalName)"
	
	#write-host "ExchangeObjectId: $($SrcMailbox.ExchangeObjectId)"
	#write-host "ExchangeGuid: $($SrcMailbox.ExchangeGuid)"
	#write-host "ArchiveStatus: $($SrcMailbox.ArchiveStatus)"
	#write-host "ArchiveGuid: $($SrcMailbox.ArchiveGuid)"
	#write-host "LegacyExchangeDN: $($SrcMailbox.LegacyExchangeDN)"
	#write-host $SRCx500Addresses -ForegroundColor Cyan
	#write-host $SRCsmtpAddresses -ForegroundColor Green
	$DstUser = Get-TargetT2TUser -SrcIdentity $SrcMailbox.ExternalDirectoryObjectId -SrcAccessToken $AuthDB[$Src_AppReg_LOG_READER].AccessToken -DstADCredential $ADCredential
	if (-not $DstUser) {
		Write-Log "No target user found for source mailbox $($SrcMailbox.UserPrincipalName) with id $($SrcMailbox.id). Skipping..." -MessageType "ERR"
		Continue
	}
	#write-host "Target user: $($DstUser.userPrincipalName)"
	<#
	if ($B2BMailUsers_DB.ContainsKey($SrcMailbox.primarysmtpaddress)) {
		$B2BMailUser = $B2BMailUsers_DB[$SrcMailbox.primarysmtpaddress]
	}
	#>
	Try {
		$DstMailUser = Get-MailUser -Identity $DstUser.samAccountName -ErrorAction Stop
	}
	Catch {
		Try {
			$DstRemoteMailbox = Get-RemoteMailbox -Identity $DstUser.samAccountName -ErrorAction Stop
			Write-Log "User $($DstUser.samAccountName) is type RemoteMailbox" -MessageType "ERR"
			Continue
		}
		Catch {
			Try {
				Enable-MailUser -Identity $DstUser.samAccountName -PrimarySmtpAddress $DstUser.userPrincipalName -Alias $DstUser.samAccountName -ExternalEmailAddress $SrcMailbox.PrimarySmtpAddress
				Write-Log "Enabled mailUser for $($DstUser.samAccountName) with PrimarySmtpAddress $($DstUser.userPrincipalName) and ExternalEmailAddress $($SrcMailbox.PrimarySmtpAddress)"
				Start-Sleep -Seconds 30
			}
			Catch {
				Write-Log "Failed to enable mailUser for $($DstUser.samAccountName). Error: $_" -messageType "ERR"
				Continue
			}
		}
	}
	If ($DstMailUser) {
		if ($DstMailUser.ExchangeGuid -ne $SrcMailbox.ExchangeGuid) {
			write-host " setting ExchangeGUID to $($SrcMailbox.ExchangeGuid)"
			Set-MailUser -Identity $DstUser.samAccountName -ExchangeGuid $SrcMailbox.ExchangeGuid
		}
		
		if ($ArchiveGuid -and $DstMailUser.ArchiveGuid -ne $ArchiveGuid) {
			write-host " setting ArchiveGuid to $($ArchiveGuid)"
			Set-MailUser -Identity $DstUser.samAccountName -ArchiveGuid $ArchiveGuid
		}
		
		if ($DstMailUser.PrimarySmtpAddress -ne $DstUser.userPrincipalName) {
			write-host " setting PrimarySmtpAddress to $($DstUser.userPrincipalName)"
			Set-MailUser -Identity $DstUser.samAccountName -PrimarySmtpAddress $DstUser.userPrincipalName
		}
		
		if ($DstMailUser.ExternalEmailAddress -ne "SMTP:$($SrcMailbox.PrimarySmtpAddress)") {
			write-host " setting ExternalEmailAddress to $($SrcMailbox.PrimarySmtpAddress)"
			Set-MailUser -Identity $DstUser.samAccountName -ExternalEmailAddress $SrcMailbox.PrimarySmtpAddress
		}
		
		if ($DstMailUser.EmailAddresses -notcontains "smtp:$($DstUser.samAccountName)@$DstMailRoutingDomain") {
			write-host " adding smtp address $($DstUser.samAccountName)@$DstMailRoutingDomain"
			Set-MailUser -Identity $DstUser.samAccountName -EmailAddresses @{add="smtp:$($DstUser.samAccountName)@$DstMailRoutingDomain"}
		}

		if ($DstMailUser.EmailAddresses -notcontains "x500:$($SrcMailbox.LegacyExchangeDN)") {
			write-host " adding LegacyExchangeDN $($SrcMailbox.LegacyExchangeDN)"
			Set-MailUser -Identity $DstUser.samAccountName -EmailAddresses @{add="x500:$($SrcMailbox.LegacyExchangeDN)"}
		}
		
		foreach ($x500 in $SRCx500Addresses) {
			if ($DstMailUser.EmailAddresses -notcontains $x500) {
				write-host " adding x500 addresses $x500"
				Set-MailUser -Identity $DstUser.samAccountName -EmailAddresses @{add="$x500"}
			}
		}
		
		if ($DstMailUser.HiddenFromAddressListsEnabled) {
			write-host "HiddenFromAddressListsEnabled false"
			Set-MailUser -Identity $DstUser.samAccountName -HiddenFromAddressListsEnabled $false
		}
	}
}

#######################################################################################################################

. $IncFile_StdLogEndBlock
