#######################################################################################################################
# Verify-T2T-All-Source-Mailboxes.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
	[string]$SourceMailbox
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1
. $ScriptPath\include\include-functions-T2T.ps1

#######################################################################################################################

$LogFolder			= "t2t"
$LogFilePrefix		= "verify-meu-properties"
$LogFileFreq		= "Y"

$OutputFolder = "t2t-ujvrez\mailusers"
$OutputFilePrefix = "verify-meu-properties"

#######################################################################################################################

. $ScriptPath\include-Script-Start-Include.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Freq $LogFileFreq -Ext "log"
$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Freq "YMDHM" -Ext "csv"

[array]$Report = @()

$Src_PN_attr = "extension_93b54ce056df45bd8f5f398753fa17c0_employeeNumber"
$Dst_PN_attr = "employeeNumber"
$Dst_Mapping_attr = "extensionAttribute10"
$MigrationEndpoint = "UJVREZ_T2T_EXO_MIGRATION_ENDPOINT"

$b2bmailusers = "D:\data\t2t-ujvrez\ujvrez-guests-mailusers.csv"

$MbxFilter1 = "(alias -like '*')"
$MbxFilter2 = " -and ((RecipientTypeDetails -eq 'UserMailbox')"
$MbxFilter3 = " -or (RecipientTypeDetails -eq 'SharedMailbox')"
$MbxFilter4 = " -or (RecipientTypeDetails -eq 'RoomMailbox')"
$MbxFilter5 = " -or (RecipientTypeDetails -eq 'EquipmentMailbox'))"
$userMbxFilter = $MbxFilter1 + $MbxFilter2 + $MbxFilter3 + $MbxFilter4 + $MbxFilter5

#######################################################################################################################

. $IncFile_StdLogStartBlock


#$B2BMailUsers_DB = Import-CSVtoHashDB -Path $b2bmailusers -KeyName "PrimarySmtpAddress"

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
$Dst_AppReg_EXO_MGMT 			= $AppReg_EXO_MGMT
$Dst_AppReg_LOG_READER 			= $AppReg_LOG_READER

$Src_T2T_EXO_MIGRATION_GROUP 	= $UJVREZ_T2T_EXO_MIGRATION_GROUP

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

Request-MSALToken -AppRegName $Dst_AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriSelect = "id,userPrincipalName,onpremisesSamAccountName"
$UriFilter = "userType eq 'Member'"
$Uri = New-GraphUri -Resource $UriResource -Select $UriSelect -Filter $UriFilter -Top 999 -Version "v1.0"
[array]$DstAADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$Dst_AppReg_LOG_READER].AccessToken
$DstAADUsers_DB = @{}
foreach ($User in $DstAADUsers) {
	$DstAADUsers_DB.Add($User.userPrincipalName, $User)
}
Remove-Variable DstAADUsers

Connect-EXOService -AppRegName $Dst_AppReg_EXO_MGMT  -TTL 60 -ForceReconnect

if ($SourceMailbox) {
	$SrcMailboxes = $SrcMailboxes | Where-Object { $_.PrimarySmtpAddress -eq $SourceMailbox }
	write-host "SRC mailboxes after filtering by PrimarySmtpAddress $($SourceMailbox): $($SrcMailboxes.count)"
	if ($SrcMailboxes.count -eq 0) {
		Write-Host "No source mailbox found with PrimarySmtpAddress $($SourceMailbox). Exiting."
		Exit
	}
	else {
		if ($SrcMailboxes.count -gt 1) {
			Write-Host "Multiple source mailboxes found with PrimarySmtpAddress $($SourceMailbox). Please check the input and try again. Exiting."
			Exit
		}
	}
}

foreach ($SrcMailbox in $SrcMailboxes) {

	$SRCsmtpAddresses = $SRCx500Addresses = $SRCproxyAddresses = @()

	$SrcArchiveGuid = $null
	$ErrorFound = $false
	$DstUser = $DstMailUser = $DstAADUser = $null

	Request-MSALToken -AppRegName $Src_AppReg_LOG_READER -TTL 30
	Connect-EXOService -AppRegName $Dst_AppReg_EXO_MGMT  -TTL 60
	#Request-MSALToken -AppRegName $Dst_AppReg_LOG_READER -TTL 30

	write-host "$($SrcMailbox.PrimarySmtpAddress)" -ForegroundColor Green
	
	if ($SrcMailbox.EmailAddresses) {
        foreach ($EmailAddress in $SrcMailbox.EmailAddresses) {
            if ($EmailAddress -like "smtp:*" -and $EmailAddress -notlike "*onmicrosoft.com") {
                $SRCsmtpAddresses += $EmailAddress
            }
            if ($EmailAddress -like "X500:*") {
                $SRCx500Addresses += $EmailAddress
            }
        }
        [string]$SRCproxyAddresses = $SrcMailbox.EmailAddresses -join ";"
    }
    if ($SrcMailbox.ArchiveGuid -and $SrcMailbox.ArchiveGuid -ne "00000000-0000-0000-0000-000000000000") {
        $SrcArchiveGuid = $SrcMailbox.ArchiveGuid
    }
	
	$ReportObject = [PSCustomObject]@{
		VERIFY_RESULT = $null
		SRC_UPN = $SrcMailbox.UserPrincipalName
		SRC_PrimarySmtpAddress = $SrcMailbox.PrimarySmtpAddress
		SRC_ExchangeGuid = $SrcMailbox.ExchangeGuid
		SRC_RecipientTypeDetails = $SrcMailbox.RecipientTypeDetails
		SRC_ArchiveGuid = $SrcArchiveGuid
		SRC_ProxyAddresses = $SRCproxyAddresses
		SRC_LegacyExchangeDN = $SrcMailbox.LegacyExchangeDN
		SRC_AADUserId = $SrcMailbox.ExternalDirectoryObjectId
		DST_UPN = $null
		DST_SAM = $null
		DST_PrimarySmtpAddress = $null
		DST_ExternalEmailAddress = $null
		DST_ArchiveGuid = $null
		DST_ProxyAddresses = $null
		DST_AADUserId = $null
		MigrationTestResult = $null
	}

	$DstUser = Get-TargetT2TUser -SrcIdentity $SrcMailbox.ExternalDirectoryObjectId -SrcAccessToken $AuthDB[$Src_AppReg_LOG_READER].AccessToken -DstADCredential $ADCredential
	
	if ($DstUser) {
		Try {
			$DstMailUser = Get-MailUser -Identity $SrcMailbox.PrimarySmtpAddress -ErrorAction Stop
		}
		Catch {
			Try {
				$DstMailbox = Get-MailBox -Identity $DstUser.userPrincipalName -ErrorAction Stop
				Write-Host "$($DstUser.userPrincipalName) is type Mailbox" -ForegroundColor Cyan
				$ReportObject.VERIFY_RESULT = "ERR_IS_MBX"
			}
			catch {
				Write-Host "Failed to get MailUser for $($SrcMailbox.PrimarySmtpAddress): $($_.Exception.Message)"
				$ReportObject.VERIFY_RESULT = "ERR_NO_MEU"
			}
		}
	}
	else {
		Write-Log "No target user found for source mailbox $($SrcMailbox.UserPrincipalName) with id $($SrcMailbox.id). Skipping..." -MessageType "ERR"
		$ReportObject.VERIFY_RESULT = "ERR_NO_TARGET_USER"
	}

	if ($DstMailUser) {
		#$DstAADUser = $DstAADUsers_DB[$DstMailUser.UserPrincipalName]
		[string]$DstProxyAddresses = $DstMailUser.EmailAddresses -join ";"

		$ReportObject.DST_UPN = $DstMailUser.userPrincipalName
		$ReportObject.DST_SAM = $DstUser.samAccountName
		$ReportObject.DST_PrimarySmtpAddress = $DstMailUser.PrimarySmtpAddress
		$ReportObject.DST_ExternalEmailAddress = $DstMailUser.ExternalEmailAddress
		$ReportObject.DST_ArchiveGuid = $DstMailUser.ArchiveGuid
		$ReportObject.DST_ProxyAddresses = $DstProxyAddresses
		$ReportObject.DST_AADUserId = $DstMailUser.ExternalDirectoryObjectId

		if ($DstMailUser.ExchangeGuid -ne $SrcMailbox.ExchangeGuid) {
			Write-Log "$($SrcMailbox.PrimarySmtpAddress): ExchangeGuid mismatch. Expected: $($SrcMailbox.ExchangeGuid), Actual: $($DstMailUser.ExchangeGuid)" -MessageType "ERR"
			$ErrorFound = $true
		}
		
		if ($SrcArchiveGuid -and ($DstMailUser.ArchiveGuid -ne $SrcArchiveGuid)) {
			Write-Log "$($SrcMailbox.PrimarySmtpAddress): ArchiveGuid mismatch. Expected: $SrcArchiveGuid, Actual: $($DstMailUser.ArchiveGuid)" -MessageType "ERR"
			$ErrorFound = $true			
		}
		
		if ($DstMailUser.PrimarySmtpAddress -ne $DstUser.userPrincipalName) {
			Write-Log "$($SrcMailbox.PrimarySmtpAddress): PrimarySmtpAddress mismatch. Expected: $($DstUser.userPrincipalName), Actual: $($DstMailUser.PrimarySmtpAddress)" -MessageType "ERR"
			$ErrorFound = $true
		}
		
		if ($DstMailUser.ExternalEmailAddress -ne "SMTP:$($SrcMailbox.PrimarySmtpAddress)") {
			Write-Log "$($SrcMailbox.PrimarySmtpAddress): ExternalEmailAddress mismatch. Expected: SMTP:$($SrcMailbox.PrimarySmtpAddress), Actual: $($DstMailUser.ExternalEmailAddress)" -MessageType "ERR"
			$ErrorFound = $true
		}
		
		
		if ($DstMailUser.EmailAddresses -notcontains "smtp:$($DstUser.samAccountName)@$DstMailRoutingDomain") {
			Write-Log "$($SrcMailbox.PrimarySmtpAddress): Missing proxy address. Expected: smtp:$($DstUser.samAccountName)@$DstMailRoutingDomain" -MessageType "ERR"
			$ErrorFound = $true
		}
		

		if ($DstMailUser.EmailAddresses -notcontains "x500:$($SrcMailbox.LegacyExchangeDN)") {
			Write-Log "$($SrcMailbox.PrimarySmtpAddress): Missing LegacyExchangeDN in proxyAddresses. Expected: x500:$($SrcMailbox.LegacyExchangeDN)" -MessageType "ERR"
			$ErrorFound = $true
		}
		
		foreach ($x500 in $SRCx500Addresses) {
			if ($DstMailUser.EmailAddresses -notcontains $x500) {
				Write-Log "$($SrcMailbox.PrimarySmtpAddress): Missing x500 address. Expected: $x500" -MessageType "ERR"
				$ErrorFound = $true
			}
		}

		if (-not $ErrorFound) {
			$ReportObject.VERIFY_RESULT = "OK"
		}
		else {
			$ReportObject.VERIFY_RESULT	 = "ERR_VAL"
		}
	}
	
	#$Result = Test-MigrationServerAvailability -EndPoint $MigrationEndpoint -TestMailbox $SrcMailbox.PrimarySmtpAddress
	#write-host "Migration endpoint test result: $Result"
	$Report += $ReportObject
}

if (-not $SourceMailbox) {
	Export-Report -Report $Report -Path $OutputFile -SortProperty "SRC_PrimarySmtpAddress"
}


#######################################################################################################################

. $IncFile_StdLogEndBlock
