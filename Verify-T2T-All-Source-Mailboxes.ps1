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
$LogFilePrefix		= "verify-all-src-mailboxes"
$LogFileFreq		= "Y"

$OutputFolder = "t2t-ujvrez\srcmailboxes"
$OutputFilePrefix = "verify-all-src-mailboxes"

#######################################################################################################################

. $ScriptPath\include-Script-Start-Include.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Freq $LogFileFreq -Ext "log"
$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Freq "YMDHM" -Ext "csv"

[array]$Report = @()
[hashtable]$MigrationUsers_DB = @{}
[hashtable]$MailUsers_DB_per_mail = @{}

$BatchPrefix = "UJV"

$Src_PN_attr = "extension_93b54ce056df45bd8f5f398753fa17c0_employeeNumber"
$Dst_PN_attr = "employeeNumber"
$Dst_Mapping_attr = "extensionAttribute10"
$MigrationEndpoint = "UJVREZ_T2T_EXO_MIGRATION_ENDPOINT"

#CEZ_Lic_M365_T2T_DataMig
$DstT2TDataMigLicGroup = "85e25c8f-4289-4197-a135-e01822cf31c4"

$b2bmailusers = "D:\data\t2t-ujvrez\ujvrez-guests-mailusers.csv"
$DoNotMigrateList = "D:\data\t2t-ujvrez\DoNotMigrate.csv"

$MbxFilter1 = "(alias -like '*')"
$MbxFilter2 = " -and ((RecipientTypeDetails -eq 'UserMailbox')"
$MbxFilter3 = " -or (RecipientTypeDetails -eq 'SharedMailbox')"
$MbxFilter4 = " -or (RecipientTypeDetails -eq 'RoomMailbox')"
$MbxFilter5 = " -or (RecipientTypeDetails -eq 'EquipmentMailbox'))"
$userMbxFilter = $MbxFilter1 + $MbxFilter2 + $MbxFilter3 + $MbxFilter4 + $MbxFilter5

$Dst_AD_OU_list = @(
	"OU=CVREZ,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=RadioMedic,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=EGP,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=EngineeringPraha,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=iCVREZ,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",	
	"OU=NQ-Safe,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=UJVREZ,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=VZUP,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=UJV,OU=aplikacni,OU=uzivatele,DC=cezdata,DC=corp"
)
$Dst_AD_OU_list_App = @(
	"OU=UJV,OU=aplikacni,OU=uzivatele,DC=cezdata,DC=corp"
)

$DstADUserProperties = @(
	'displayName',
	'mail',
	'distinguishedName',
	'samAccountName',
	'userPrincipalName',
	'extensionAttribute10'
)

$DstADSearchBaseAll  = "OU=uzivatele,DC=cezdata,DC=corp"
$DstADSearchBaseSKC  = "OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp"

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

$DstGetADUserParams = @{
	Filter = "ObjectClass -eq 'user'"
	SearchBase = $DstADSearchBaseAll
	Properties = $DstADUserProperties
	Credential = $ADCredential
}

$DstMailRoutingDomain = "cezdata.mail.onmicrosoft.com"

$Src_AppReg_LOG_READER 			= $AppReg_UJVREZ_LOG_READER
$Src_AppReg_EXO_MGMT 			= $AppReg_UJVREZ_EXO_MGMT
$Dst_AppReg_EXO_MGMT 			= $AppReg_EXO_MGMT
$Dst_AppReg_LOG_READER 			= $AppReg_LOG_READER

$Src_T2T_EXO_MIGRATION_GROUP 	= $UJVREZ_T2T_EXO_MIGRATION_GROUP

$Src_AppReg_LOG_READER 			= $AppReg_UJVREZ_LOG_READER
$Src_AppReg_EXO_MGMT 			= $AppReg_UJVREZ_EXO_MGMT   
$Src_T2T_EXO_MIGRATION_GROUP 	= $UJVREZ_T2T_EXO_MIGRATION_GROUP

Request-MSALToken -AppRegName $Src_AppReg_LOG_READER -TTL 30
Connect-EXOService -AppRegName $Src_AppReg_EXO_MGMT  -TTL 120
$ExchangeSession = New-PSSession -Name "OnPremExchange" -ConfigurationName "Microsoft.Exchange" -ConnectionUri "http://cw00616exch3.cezdata.corp/PowerShell/" -Authentication Kerberos
Import-PSSession $ExchangeSession -DisableNameChecking -AllowClobber

#######################################################################################################################
# get DoNotMigrate list
#######################################################################################################################
$DoNotMigrate = Import-Csv -Path $DoNotMigrateList
Write-Host "DoNotMigrate list count: $($DoNotMigrate.count)" -ForegroundColor Yellow

#######################################################################################################################
# get T2T migration group members
#######################################################################################################################
$T2TMigrationGroupName = Get-GroupNameFromGraphById -AccessToken $AuthDB[$Src_AppReg_LOG_READER].AccessToken -Id $Src_T2T_EXO_MIGRATION_GROUP
if ($T2TMigrationGroupName) {
	Write-Host "SRC T2T migration group found: $T2TMigrationGroupName"
} else {
	Write-Host "SRC T2T migration group with ID $Src_T2T_EXO_MIGRATION_GROUP not found."
	Exit
}

write-host "SRC T2T migration group members: " -noNewline
$UriResource = "groups/$($Src_T2T_EXO_MIGRATION_GROUP)/members"
$Uri = New-GraphUri -Resource $UriResource -Version "v1.0"
[array]$T2TMigrationGroupMembers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$Src_AppReg_LOG_READER].AccessToken
write-host $T2TMigrationGroupMembers.count

#######################################################################################################################
# get SRC mailboxes
#######################################################################################################################
write-host "SRC mailboxes total: " -NoNewline
[array]$SrcMailboxes = Get-EXOMailbox -ResultSize Unlimited -PropertySets All
write-host $SrcMailboxes.count

#######################################################################################################################
# T2T Licensing group members
#######################################################################################################################

Request-MSALToken -AppRegName $Dst_AppReg_LOG_READER -TTL 30
write-host "DST T2T Licensing group members total: " -NoNewline
$UriResource = "groups/$($DstT2TDataMigLicGroup)/members"
$Uri = New-GraphUri -Resource $UriResource -Version "v1.0"
[array]$T2TDataMigLicGroupMembers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$Dst_AppReg_LOG_READER].AccessToken
write-host $T2TDataMigLicGroupMembers.count

#######################################################################################################################
# get DST AD users
#######################################################################################################################
write-host "DST AD users - reading from AD (long running operation)..." -NoNewline
$DstADUsers = Get-ADUser @DstGetADUserParams | Select-Object $DstADUserProperties
write-host "done ($($DstADUsers.count))"

#check duplicates in ext10 before proceeding
write-host "DST AD users - checking duplicate ext10..." -NoNewline
$DstADUsersExt10 = $DstADUsers | Where-Object { $_.extensionAttribute10 -ne $null -and $_.extensionAttribute10 -ne "" -and $_.extensionAttribute10 -notlike "_*" }
$duplicateUsers = $DstADUsersExt10 | Group-Object -Property "extensionAttribute10" | Where-Object { $_.Count -gt 1 }
foreach ($group in $duplicateUsers) {
	write-host "Duplicate ext10: $($group.Name) - Count: $($group.Count)"
	foreach ($user in $group.Group) {
		write-host "  User: $($user.displayName) - UPN: $($user.userPrincipalName)" -ForegroundColor Red
	}
	exit
}
write-host "done"

#filter out users with samAccountName starting with QT
$DstADUsers = $DstADUsers | Where-Object { $_.samAccountName -notlike 'QT*' }
write-host "DST AD users - (filtered out QT): $($DstADUsers.count)"

#filter out users with samAccountName starting with QK
$DstADUsers = $DstADUsers | Where-Object { $_.samAccountName -notlike 'QK*' }
write-host "DST AD users - (filtered out QK): $($DstADUsers.count)"

#OU property to each user object by parsing it from DistinguishedName, we will need it for filtering users by OU and for reporting
write-host "DST AD users - adding OU property..." -NoNewline
foreach ($user in $DstADUsers) {    
	$user | Add-Member -NotePropertyName 'OU' -NotePropertyValue ( $user.DistinguishedName -replace '^CN=[^,]+,' )
}
write-host "done"

#filter out only users from specific OUs
$DstADUsers = $DstADUsers | Where-Object { $_.OU -in $Dst_AD_OU_list }
write-host "DST AD users - filtered by OU: $($DstADUsers.count)"

Connect-EXOService -AppRegName $Dst_AppReg_EXO_MGMT  -TTL 60 -ForceReconnect

#######################################################################################################################
# get migration batches and users in migration
#######################################################################################################################
$MigrationBatches = Get-MigrationBatch
$MigrationBatches = $MigrationBatches | Where-Object { $_.Identity -like "$BatchPrefix*" }
Write-host "Found $($MigrationBatches.Count) migration batches with prefix '$BatchPrefix'"

foreach ($Batch in $MigrationBatches) {
	write-host ($Batch.Identity.ToString()).PadRight(25) -NoNewline
	$Result = Get-MigrationUser -BatchId $Batch.Identity
	write-host ": $($Result.Count)"
	$MigrationUsers += $Result
}
write-host "Found $($MigrationUsers.Count) migration users in batches with prefix '$BatchPrefix'"
write-host "Creating migration users DB..." -NoNewline
foreach ($MigrationUser in $MigrationUsers) {
	$MigrationUsers_DB.Add($MigrationUser.Identity, $MigrationUser)
}
write-host "done ($($MigrationUsers_DB.count))" 

#######################################################################################################################
# get DST mailusers
#######################################################################################################################
write-host "DST mailusers total: " -NoNewline
#$DstMailUsers = Get-MailUser -ResultSize Unlimited
write-host $DstMailUsers.count

write-host "Creating DST mailusers DB..." -NoNewline
foreach ($DstMailUser in $DstMailUsers) {
	$MailUsers_DB_per_mail.Add($DstMailUser.externalEmailAddress, $DstMailUser)
}
write-host "done ($($MailUsers_DB_per_mail.count))"

#######################################################################################################################
# main loop - iterate through source mailboxes in T2T migration group and verify properties of corresponding mailuser in target tenant
#######################################################################################################################
foreach ($SrcMailbox in $SrcMailboxes) {

	if ($SrcMailbox.PrimarySmtpAddress -like "DiscoverySearchMailbox*") {
		continue
	}

	$DstProxyAddresses = $SRCsmtpAddresses = $SRCx500Addresses = $SRCproxyAddresses = @()

	$SrcArchiveGuid = $null
	$ErrorFound = $false
	$MigrationUser = $DstMailUser = $DstMailbox = $DstADUser = $null

	Request-MSALToken -AppRegName $Src_AppReg_LOG_READER -TTL 30
	Connect-EXOService -AppRegName $Dst_AppReg_EXO_MGMT  -TTL 60
	#Request-MSALToken -AppRegName $Dst_AppReg_LOG_READER -TTL 30

	write-host "$($SrcMailbox.PrimarySmtpAddress) " -NoNewline
	
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
		MEU_VERIFY = $null
		MIGRATION_STATUS = $null
		T2T_SCOPE = $null
		T2T_LIC = $null
		SRC_UPN = $SrcMailbox.UserPrincipalName
		SRC_PrimarySmtpAddress = $SrcMailbox.PrimarySmtpAddress
		SRC_DisplayName = $SrcMailbox.DisplayName
		SRC_ExchangeGuid = $SrcMailbox.ExchangeGuid
		SRC_RecipientTypeDetails = $SrcMailbox.RecipientTypeDetails
		SRC_ArchiveGuid = $SrcArchiveGuid
		SRC_ProxyAddresses = $SRCproxyAddresses
		SRC_LegacyExchangeDN = $SrcMailbox.LegacyExchangeDN
		SRC_AADUserId = $SrcMailbox.ExternalDirectoryObjectId
		DST_UPN = $null
		DST_SAM = $null
		DST_DisplayName = $null
		DST_ext10 = $null
		DST_PrimarySmtpAddress = $null
		DST_ExternalEmailAddress = $null
		DST_ArchiveGuid = $null
		DST_ProxyAddresses = $null
		DST_AADUserId = $null
		MigrationTestResult = $null
	}

	if ($T2TMigrationGroupMembers.id -contains $SrcMailbox.ExternalDirectoryObjectId) {
		$ReportObject.T2T_SCOPE = $true
	}
	else {
		$ReportObject.T2T_SCOPE = $false
	}

	if ($DoNotMigrate.Mail -contains $SrcMailbox.PrimarySmtpAddress) {
		$ReportObject.MIGRATION_STATUS = "DO_NOT_MIGRATE"
		$Report += $ReportObject
		continue
	}

	$DstADUser = Get-TargetT2TUser -SrcIdentity $SrcMailbox.ExternalDirectoryObjectId -SrcAccessToken $AuthDB[$Src_AppReg_LOG_READER].AccessToken -DstADCredential $ADCredential
	
	if ($DstADUser) {
		$ReportObject.DST_SAM = $DstADUser.samAccountName
		$ReportObject.DST_UPN = $DstADUser.userPrincipalName
		$ReportObject.DST_ext10 = $DstADUser.extensionAttribute10
		$ReportObject.DST_DisplayName = $DstADUser.DisplayName
	}
	else {
		Write-Host
		Write-Log "No target user found for source mailbox $($SrcMailbox.UserPrincipalName) with id $($SrcMailbox.id)" -MessageType "ERR"
		$ReportObject.MEU_VERIFY = "ERR_NO_TARGET_USER"
	}

	Try {
		$DstMailUser = Get-MailUser -Identity $SrcMailbox.PrimarySmtpAddress -ErrorAction Stop
		if ($DstMailUser.RecipientTypeDetails -eq "GuestMailUser") {
			$ReportObject.MEU_VERIFY = "ERR_IS_GUEST"
		}
	}
	Catch {
		if ($DstADUser) {
			Try {
				$DstMailbox = Get-MailBox -Identity $DstADUser.userPrincipalName -ErrorAction Stop
				Write-Host "$($DstADUser.userPrincipalName) is type Mailbox" -ForegroundColor Cyan
				$ReportObject.MEU_VERIFY = "ERR_IS_MBX"
			}
			catch {
				Write-Host "Failed to get MEU or mailbox for $($SrcMailbox.PrimarySmtpAddress)" -ForegroundColor Red
				$ReportObject.MEU_VERIFY = "ERR_NO_MEU"
			}
		}
		else {
			$ReportObject.MEU_VERIFY = "ERR_NO_ADUSER"
		}
	}

	if ($DstMailUser) {
		 if ($T2TDataMigLicGroupMembers.id -contains $DstMailUser.ExternalDirectoryObjectId) {
			$ReportObject.T2T_LIC = $true
		}
		else {
			$ReportObject.T2T_LIC = $false
		}
	}

	if ($DstMailUser -and $DstMailUser.RecipientTypeDetails -eq "MailUser") {
		if ($MigrationUsers_DB.ContainsKey($DstMailUser.userPrincipalName)) {
			$MigrationUser = $MigrationUsers_DB[$DstMailUser.userPrincipalName]
			$ReportObject.MIGRATION_STATUS = $MigrationUser.Status
		}
		else {
			$ReportObject.MIGRATION_STATUS = "NOT_IN_MIGRATION"
		}
	}

	if ($DstMailUser -and $DstMailUser.RecipientTypeDetails -eq "MailUser") {
		#$DstAADUser = $DstAADUsers_DB[$DstMailUser.UserPrincipalName]
		[string]$DstProxyAddresses = $DstMailUser.EmailAddresses -join ";"

		$ReportObject.DST_UPN = $DstMailUser.userPrincipalName
		
		$ReportObject.DST_PrimarySmtpAddress = $DstMailUser.PrimarySmtpAddress
		$ReportObject.DST_ExternalEmailAddress = $DstMailUser.ExternalEmailAddress
		$ReportObject.DST_ArchiveGuid = $DstMailUser.ArchiveGuid
		$ReportObject.DST_ProxyAddresses = $DstProxyAddresses
		$ReportObject.DST_AADUserId = $DstMailUser.ExternalDirectoryObjectId

		if ($DstMailUser.ExchangeGuid -ne $SrcMailbox.ExchangeGuid) {
			Write-Host
			Write-Log "$($SrcMailbox.PrimarySmtpAddress): ExchangeGuid mismatch. Expected: $($SrcMailbox.ExchangeGuid), Actual: $($DstMailUser.ExchangeGuid)" -MessageType "ERR"
			$ErrorFound = $true
		}
		
		if ($SrcArchiveGuid -and ($DstMailUser.ArchiveGuid -ne $SrcArchiveGuid)) {
			Write-Host
			Write-Log "$($SrcMailbox.PrimarySmtpAddress): ArchiveGuid mismatch. Expected: $SrcArchiveGuid, Actual: $($DstMailUser.ArchiveGuid)" -MessageType "ERR"
			$ErrorFound = $true			
		}
		
		if ($DstMailUser.PrimarySmtpAddress -ne $DstADUser.userPrincipalName) {
			Write-Host
			Write-Log "$($SrcMailbox.PrimarySmtpAddress): PrimarySmtpAddress mismatch. Expected: $($DstADUser.userPrincipalName), Actual: $($DstMailUser.PrimarySmtpAddress)" -MessageType "ERR"
			$ErrorFound = $true
		}
		
		if ($DstMailUser.ExternalEmailAddress -ne "SMTP:$($SrcMailbox.PrimarySmtpAddress)") {
			Write-Host
			Write-Log "$($SrcMailbox.PrimarySmtpAddress): ExternalEmailAddress mismatch. Expected: SMTP:$($SrcMailbox.PrimarySmtpAddress), Actual: $($DstMailUser.ExternalEmailAddress)" -MessageType "ERR"
			$ErrorFound = $true
		}
		
		
		if ($DstMailUser.EmailAddresses -notcontains "smtp:$($DstADUser.samAccountName)@$DstMailRoutingDomain") {
			Write-Host
			Write-Log "$($SrcMailbox.PrimarySmtpAddress): Missing proxy address. Expected: smtp:$($DstADUser.samAccountName)@$DstMailRoutingDomain" -MessageType "ERR"
			$ErrorFound = $true
		}

		if ($DstMailUser.EmailAddresses -notcontains "x500:$($SrcMailbox.LegacyExchangeDN)") {
			Write-Host
			Write-Log "$($SrcMailbox.PrimarySmtpAddress): Missing LegacyExchangeDN in proxyAddresses. Expected: x500:$($SrcMailbox.LegacyExchangeDN)" -MessageType "ERR"
			$ErrorFound = $true
		}
		
		foreach ($x500 in $SRCx500Addresses) {
			if ($DstMailUser.EmailAddresses -notcontains $x500) {
				Write-Host
				Write-Log "$($SrcMailbox.PrimarySmtpAddress): Missing x500 address. Expected: $x500" -MessageType "ERR"
				$ErrorFound = $true
			}
		}

		if (-not $ErrorFound) {
			$ReportObject.MEU_VERIFY = "OK"
			Write-Host " - OK ($($ReportObject.MIGRATION_STATUS))" -ForegroundColor Green
		}
		else {
			$ReportObject.MEU_VERIFY = "ERR_VAL"
		}
	}
	$Report += $ReportObject
}

Export-Report -Report $Report -Path $OutputFile -SortProperty "SRC_PrimarySmtpAddress"



#######################################################################################################################

. $IncFile_StdLogEndBlock
