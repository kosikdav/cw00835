#######################################################################################################################
# Get-AAD-Users-Reports.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder					= "exports"
$LogFilePrefix				= "aad-users-report"

$OutputFolder				= "aad-users\reports"
$OutputFolderCopilot		= "copilot\reports"
$OutputFilePrefix			= "aad-users"
$OutputFileSuffixDeletedUsers	= "deleted"
$OutputFileSuffixCopilotLic	= "copilot-lic"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputFile 	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Ext "csv"
$OutputFileDeletedUsers = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixDeletedUsers -Ext "csv"
$OutputFileCopilotLic = New-OutputFile -RootFolder $ROF -Folder $OutputFolderCopilot -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixCopilotLic -Ext "csv"

$ADCredentialPath = $aadauthmobmgmt_cred

[array]$UserListReport = @()
[array]$DeletedUserListReport = @()
[array]$CopilotLicenseReport = @()
[hashtable]$SIA_DB = @{}
[hashtable]$ADUser_DB = @{}

$now = Get-Date

#######################################################################################################################

. $IncFile_StdLogBeginBlock

##############################################################################
# deleted users
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "directory/deletedItems/microsoft.graph.user"
$UriSelect = "id,userPrincipalName,deletedDateTime,department,companyName,displayName,mail,onPremisesSamAccountName,onPremisesUserPrincipalName"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
[array]$deletedUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON

foreach ($user in $deletedUsers) {
    if ($user.userPrincipalName) {
        $deletedUserObject = [pscustomobject]@{
			id = $user.id;
            userPrincipalName = $user.UserPrincipalName;
            DisplayName = $user.displayName;
            Department = $user.department;
            Company = $user.companyName;
            Mail = $user.mail;
            SamAccountName = $user.onPremisesSamAccountName
			onPremisesUserPrincipalName = $user.onPremisesUserPrincipalName;
			deletedDateTime = $user.deletedDateTime;
            daysSinceDeleted = (New-TimeSpan -Start $User.deletedDateTime -End $now).Days;
        }
		$DeletedUserListReport += $deletedUserObject
    }
}
Export-Report -Text "AAD deleted users report" -Report $DeletedUserListReport -Path $OutputFileDeletedUsers -SortProperty "userPrincipalName"

$UserLicensing_DB = Import-CSVtoHashDB -Path $DBFileUsersMemLic -KeyName "id"

if ($TenantShortName -eq "CEZDATA") {
	$ADProperties = @("userPrincipalName","msDS-cloudExtensionAttribute1","msExchExtensionAttribute29","msExchExtensionAttribute40","cEZIntuneMFAAuthMobile")
	$ADFilter = "(sAMAccountName -notlike `"qh*`") -and (msExchExtensionAttribute29 -like `"*`") -and (msExchExtensionAttribute40 -like `"*`")"
	if ($interactiveRun) {
		[array]$ADUsers = Get-ADUser -Filter $ADFilter -Properties $ADProperties
	}
	else {
		Write-Log "AD credential file: $($ADCredentialPath)"
		$ADCredential = Import-Clixml -Path $ADCredentialPath
		[array]$ADUsers = Get-ADUser -Credential $ADCredential -Filter $ADFilter -Properties $ADProperties
	}
	Write-Log "AD users: $(Get-Count -Object $ADUsers)"
	foreach ($ADUser in $ADUsers) {
		$UserObject = [pscustomobject]@{
			msDScloudExtensionAttribute1	= $ADUser.'msDS-cloudExtensionAttribute1';
			msExchExtensionAttribute29		= $ADUser.msExchExtensionAttribute29;
			msExchExtensionAttribute40		= $ADUser.msExchExtensionAttribute40;
			cEZIntuneMFAAuthMobile			= $ADUser.cEZIntuneMFAAuthMobile
		}
		$ADUser_DB.Add($ADUser.userPrincipalName, $UserObject)
	}
}

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriSelect1 = "id,UserPrincipalName,DisplayName,UserType,AccountEnabled,mail,mailNickname,companyName,department,JobTitle,mobilePhone,officeLocation,preferredLanguage"
$UriSelect2 = "CreatedDateTime,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesSamAccountName,onPremisesDistinguishedName,onPremisesImmutableId"
$UriSelect = $UriSelect1 , $UriSelect2 -join ","
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Select $UriSelect
[array]$Users = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "users" -ProgressDots

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriSelect = "id,signInActivity"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 99 -Select $UriSelect
[array]$UsersSIA = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "users (signInActivity)" -ProgressDots
$UsersSIA | ForEach-Object {$SIA_DB.Add($_.id, $_.signInActivity)}

Write-Log "Total users: $($Users.Count) SIA users: $($UsersSIA.Count)"

ForEach ($User in $Users) {
	Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
	$UserLicensingRecord = $null
	$Mail = $MailDomain = $mobilePhone = $mobilePhoneAuth = $ODfBUrl = [string]::Empty
	$AADPremLicense = $CopilotLicense = $EXOLicense = $SPOLicense = $TMSLicense = $IntuneLicense = $PwrAutLicense = $PwrAppLicense = [string]::Empty
	$AADPremLicenseNeeded = $false
	$LastSignInDateTime = $LastSignInDateTime_NI = "never"
	$DaysSinceLastSignIn = $DaysSinceLastSignIn_NI = "n/a"
	$GroupMemberCount = 0
	$UserDrive = $null
	$onPremisesImmutableId = $onPremisesGUID = [string]::Empty

	if ($User.Mail) {
		$Mail = $User.Mail.ToLower()
		$MailDomain = $Mail.Split("@")[1]
	}

	if ($SIA_DB.Contains(($User.id))) {
		$SIA = $SIA_DB.Item($User.id)
		if (-not (($null -eq $SIA.LastSignInDateTime) -or ($SIA.LastSignInDateTime -eq '1/1/0001 1:00:00 AM'))) {
			$LastSignInDateTime	= [DateTime]$SIA.LastSignInDateTime
			$DaysSinceLastSignIn = (New-TimeSpan -Start $SIA.LastSignInDateTime -End $Today).Days
		}
		if (-not (($null -eq $SIA.LastNonInteractiveSignInDateTime) -or ($SIA.LastNonInteractiveSignInDateTime -eq '1/1/0001 1:00:00 AM'))) {
			$LastSignInDateTime_NI = [DateTime]$SIA.LastNonInteractiveSignInDateTime
			$DaysSinceLastSignIn_NI = (New-TimeSpan -Start $SIA.LastNonInteractiveSignInDateTime -End $Today).Days
		}	
	}

	if ($AADUserReportGroupMemberCount) {
		$UriResource = "users/$($User.id)/memberOf/`$count"
		$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
		try {
			$GroupMemberCount = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ConsistencyLevel "eventual"
		}
		catch {
			write-host "Get-GraphOutputREST failed for group member count of $($User.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Red
			$GroupMemberCount = "n/a"
		}
	}
	
	if ($User.UserType -eq "Member") {
		if ($UserLicensing_DB.ContainsKey($User.id)) {
			$UserLicensingRecord = $UserLicensing_DB[$User.id]
		}
	}

	if ($User.CreatedDateTime) {
		$CreatedDate = (([DateTime]$User.CreatedDateTime).ToUniversalTime()).ToString("yyyy-MM-dd")
		$DaysSinceCreated = (New-TimeSpan -Start $createdDate -End $currentDate).Days
	}
	else {
		$DaysSinceCreated = "n/a"
	}

	if ($User.mobilePhone) {
		$mobilePhone = "Tel:" + $User.mobilePhone
	}
	if ($ADUser_DB[$User.UserPrincipalName].cEZIntuneMFAAuthMobile) {
		$mobilePhoneAuth = "Tel:" + $ADUser_DB[$User.UserPrincipalName].cEZIntuneMFAAuthMobile
	}
	if ($User.onPremisesImmutableId) {
		$onPremisesImmutableId = $User.onPremisesImmutableId.Trim()
		$onPremisesGUID = ([Guid]([Convert]::FromBase64String($onPremisesImmutableId))).Guid
	}

	if ($User.userType -eq "Member") {
		$UriResource = "users/$($User.id)/drive"
		$UriSelect = "driveType,owner,webUrl"
		$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
		Try {
			$UserDrive = Invoke-RestMethod -Uri $Uri -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -ContentType $ContentTypeJSON
		}
		Catch {
			if ($_.Exception.Message -like "*(404) Not Found*") {
				$ODfBUrl = [string]::Empty
			}
			else {
				Write-Host "Get-GraphOutputREST failed for drive of $($User.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Red
			}
		}
		if ($UserDrive) {
			$ODfBUrl = $UserDrive.webUrl.Trim("/Documents")
		}
	}

	$UserObject = [pscustomobject]@{
		UserId						= $User.id
		UserPrincipalName 			= $User.UserPrincipalName
		UPNDomain					= $User.UserPrincipalName.Split("@")[1]
		DisplayName 				= $User.DisplayName;
		UserType 					= $User.UserType
		Enabled						= $User.AccountEnabled
		Mail 						= $Mail
		MailDomain	 				= $MailDomain
		MailNickname				= $User.mailNickname
		CompanyName					= $User.companyName
		Department					= $User.department
		JobTitle					= $User.JobTitle
		OfficeLocation				= $User.officeLocation
		MobilePhone 				= $mobilePhone
		preferredLanguage			= $User.preferredLanguage
		ODfBUrl						= $ODfBUrl
		LastSignIn					= $LastSignInDateTime
		DaysSinceLastSignIn			= $DaysSinceLastSignIn
		LastSignIn_NI				= $LastSignInDateTime_NI
		DaysSinceLastSignIn_NI		= $DaysSinceLastSignIn_NI
		CreatedDateTime 			= $User.CreatedDateTime;
		DaysSinceCreated			= $DaysSinceCreated
		GroupMemberCount 			= $GroupMemberCount
		SyncEnabled 				= $User.onPremisesSyncEnabled
		LastSyncTime 				= $User.onPremisesLastSyncDateTime
		onPremisesSamAccountName	= $User.onPremisesSamAccountName
		onPremisesDN				= $User.onPremisesDistinguishedName
		onPremisesImmutableId		= $onPremisesImmutableId
		onPremisesGUID				= $onPremisesGUID
		E3							= $UserLicensingRecord.M365E3SKU
		E3UAT						= $UserLicensingRecord.M365E3UATSKU
		F3							= $UserLicensingRecord.M365F3SKU
		E5							= $UserLicensingRecord.M365E5SKU
		E5SEC						= $UserLicensingRecord.M365E5SecSKU
		Copilot						= $UserLicensingRecord.M365CopilotSKU
		EXOLicense					= $UserLicensingRecord.EXOLicense
		SPOLicense					= $UserLicensingRecord.SPOLicense
		TMSLicense 					= $UserLicensingRecord.TMSLicense
		IntuneLicense				= $UserLicensingRecord.IntuneLicense
		PwrAutLicense 				= $UserLicensingRecord.PwrAutLicense
		PwrAppLicense 				= $UserLicensingRecord.PwrAppLicense
		CopilotLicense				= $UserLicensingRecord.CopilotLicense
		AADPremLicense				= $UserLicensingRecord.AADPremLicense
		AADPremLicenseNeeded		= $UserLicensingRecord.AADPremLicenseNeeded
	}
	
	if ($TenantShortName -eq "CEZDATA") {
		Add-Member $UserObject -NotePropertyName "mobilePhoneAuth" -NotePropertyValue $mobilePhoneAuth
		Add-Member $UserObject -NotePropertyName "msExchExtensionAttribute29" -NotePropertyValue $ADUser_DB[$User.UserPrincipalName].msExchExtensionAttribute29
		Add-Member $UserObject -NotePropertyName "msExchExtensionAttribute40" -NotePropertyValue $ADUser_DB[$User.UserPrincipalName].msExchExtensionAttribute40
	}

	if ($AADUserReportTNR) {
		add-member $UserObject -NotePropertyName "DepartmentTNR" -NotePropertyValue $ADUser_DB[$User.UserPrincipalName].msDScloudExtensionAttribute1
	}

	$UserListReport += $UserObject
	
	if ($UserLicensingRecord.CopilotLicense) {
		$UserMinObject = [pscustomobject]@{
			UserId						= $User.id
			UserPrincipalName 			= $User.UserPrincipalName
			UPNDomain					= $User.UserPrincipalName.Split("@")[1]
			DisplayName 				= $User.DisplayName
			UserType 					= $User.UserType
			Enabled						= $User.AccountEnabled
			Mail 						= $Mail
			MailDomain	 				= $MailDomain
			CompanyName					= $User.companyName
			Department					= $User.department
			JobTitle					= $User.JobTitle
			OfficeLocation				= $User.officeLocation
			onPremisesSamAccountName	= $User.onPremisesSamAccountName
			onPremisesDN				= $User.onPremisesDistinguishedName
			EXOLicense					= $EXOLicense
			SPOLicense					= $SPOLicense
			TMSLicense 					= $TMSLicense
			IntuneLicense				= $IntuneLicense
			PwrAutLicense 				= $PwrAutLicense
			PwrAppLicense 				= $PwrAppLicense
			CopilotLicense				= $CopilotLicense
			AADPremLicense				= $AADPremLicense
			AADPremLicenseNeeded		= $AADPremLicenseNeeded
		}
		if ($AADUserReportTNR) {
			add-member $UserMinObject -NotePropertyName "DepartmentTNR" -NotePropertyValue $ADUser_DB[$User.UserPrincipalName].msDScloudExtensionAttribute1
		}
		$CopilotLicenseReport += $UserMinObject
	}
}

Export-Report -Text "AAD users report" -Report $UserListReport -Path $OutputFile -SortProperty "UserPrincipalName"
Export-Report -Text "AAD users report (DB folder)" -Report $UserListReport -Path $DBFileUsers -SortProperty "UserPrincipalName"
Export-Report -Text "AAD users report (Copilot lic)" -Report $CopilotLicenseReport -Path $OutputFileCopilotLic -SortProperty "UserPrincipalName"

#######################################################################################################################

. $IncFile_StdLogStartBlock
