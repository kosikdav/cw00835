#######################################################################################################################
# Get-AAD-Guests-Reports
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder				= "exports"
$LogFilePrefix			= "aad-guests-reports"

$OutputFolder			= "aad-guests\reports"
$OutputFolderT2T		= "aad-guests\reports\t2t"
$OutputFilePrefix		= "aad-guests"
$OutputFilePrefixT2T	= "aad-guests-T2T"

$InactivityLimit	= 365
$AuditLogRetention	= 180

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile 		= New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$OutputFile 	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Ext "csv"
$OutputFileT2T 	= New-OutputFile -RootFolder $ROF -Folder $OutputFolderT2T -Prefix $OutputFilePrefixT2T -Ext "csv"

##################################################################################################

. $IncFile_StdLogStartBlock

[System.Collections.ArrayList]$GuestListReport = @()
[System.Collections.ArrayList]$GuestListReportT2T = @()

$AADEXTTenant_DB = Import-CSVtoHashDB -Path $DBFileExtAADTenants -KeyName "domain"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

$UriResource = "users"
$UriFilter = "UserType+eq+'Guest'"
$UriSelect1 = "id,userPrincipalName,accountEnabled,createdDateTime,displayName,employeeType,employeeHireDate,userType,companyName,signInActivity,mail,creationType"
$UriSelect2 = "externalUserState,externalUserStateChangeDateTime,otherMails,proxyAddresses,identities,onPremisesExtensionAttributes,showInAddressList"

$UriSelect = $UriSelect1, $UriSelect2 -join ","
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 99 -Filter $UriFilter -Select $UriSelect
$Guests = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -text "AAD Guests" -ProgressDots
write-host $Guests.count
ForEach ($Guest in $Guests) {
	Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
	write-host "$($Guest.displayName) " -ForegroundColor White -nonewline
	$proxyAddresses	= $otherMails = $currentAADTenant = $ext15 = $LastAuditTrace = $null
	$ExtTenantInboundSync = $T2TSync = $issuerMSA = $issuerExtAAD = $issuerMail = $false
	$ExtTenantId = $ExtTenantDisplayName = $ExtTenantDefaultDomain = $lsi = $lsiNI = $MSAGuestForAADOrg = $null
	$LastSignInDateTime = $LastSignInDateTime_NI = $LastSignInDateTime_success = "never"
	$DaysSinceLastSignIn = $DaysSinceLastSignIn_NI = $DaysSinceLastSignIn_success = $DaysSinceLAT = -1
	$strUPNExtMail = $strUPNExtMailDomain = "OK"
	$GroupMemberCount = -1
	
	$upn = $Guest.UserPrincipalName
	$ext15 = $Guest.onPremisesExtensionAttributes.extensionAttribute15

	if ($ext15 -and $ext15.StartsWith("XTSync_")) {
		$T2TSync = $true
		$UriResource = "users/$($Guest.id)"
		$UriSelect1 = "id,streetAddress,city,state,postalCode,country,department,employeeId,jobTitle"
		$UriSelect2 = "manager,physicalDeliveryOfficeName,preferredLanguage,telephoneNumber"
		$UriSelect3 = "officeLocation,streetAddress,state,city,country,mobilePhone,businessPhones"
		$UriSelect = $UriSelect1, $UriSelect2, $UriSelect3 -join ","
		$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
		$GuestT2T  = Invoke-RestMethod -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON -ErrorAction Stop
		$UriResource = "users/$($Guest.id)/manager"
		$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
		Try {
			$GuestT2TManager  = Invoke-RestMethod -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON -ErrorAction Stop
		}
		Catch {
			$GuestT2TManager = $null
		}
	}

	### mail, mailDomain ###
	$Mail = $Guest.Mail.ToLower()
	$mailDomain	= $Mail.Split("@")[1]

	### proxyAddresses ###
	$proxyAddresses = $Guest.proxyAddresses
	if ($Guest.proxyAddresses.Count -ge 0) { 
		$strProxyAddresses = ($proxyAddresses | ConvertTo-Json -Compress | Out-String).Trim("`t`n`r")
	}
	switch ($proxyAddresses.Count) {
		0	{ $strProxyAddresses = "none" }
		1 	{ if ($proxyAddresses[0] -eq ("SMTP:"+$Mail)) { $strProxyAddresses = "OK" } }
		default { if ($proxyAddresses.Contains("SMTP:"+$mail)) { $strProxyAddresses = "OK:" + $strProxyAddresses } }
	}
	
	### otherMails ###
	$otherMails = $Guest.otherMails
	if ($Guest.otherMails.Count -ge 0) { 
		$strOtherMails = ($otherMails | ConvertTo-Json -Compress | Out-String).Trim("`t`n`r")
	}
	switch ($OtherMails.Count) {
		0 	{ $strOtherMails = "none" }
		1 	{ if ($otherMails[0] -eq $Mail) { $strOtherMails = "OK"} }
		default { if ($otherMails.Contains($mail)) { $strOtherMails = "OK:" + $strOtherMails } }
	}

	### EXTAAD ###
	foreach ($Identity in $Guest.Identities) {
		if (($Identity.Issuer -eq "MicrosoftAccount") -and ($mailDomain -ne "microsoft.com")) {
			$issuerMSA = $true
		}
		if ($Identity.Issuer -eq "ExternalAzureAD") {
			$issuerExtAAD = $true
		}
		if ($Identity.Issuer -eq "mail") {
			$issuerMail = $true
		}
	} 
	if ($AADEXTTenant_DB.ContainsKey($mailDomain)) {
		$currentAADTenant = $AADEXTTenant_DB.Item($mailDomain)
		$ExtTenantId 			= $currentAADTenant.tenantId
		$ExtTenantDisplayName 	= $currentAADTenant.displayName
		$ExtTenantDefaultDomain = $currentAADTenant.defaultDomainName
		$ExtTenantInboundSync 	= $currentAADTenant.InboundSyncAllowed
	}
	
	if ($T2TSync) {
		write-host "($($ExtTenantDisplayName))" -ForegroundColor Green
	}
	else {
		write-host "($($ExtTenantDisplayName))" -ForegroundColor Gray
	}

	if ($currentAADTenant) {
		if  ($issuerMSA) {
			$MSAGuestForAADOrg = $true
		}
		Else {
			$MSAGuestForAADOrg = $false
		}
	}

	### UPNExtMail, UPNExtMailDomain ###
	$UPNExtMail = GetMailFromGuestUPN ($upn)
	$UPNExtMailDomain = ($UPNExtMail.Split("@")[1]).ToLower()
	if ($UPNExtMail -ne $Mail) { $strUPNExtMail = GetMailFromGuestUPN ($upn) }
	if ($UPNExtMailDomain -ne $mailDomain) { $strUPNExtMailDomain = $UPNExtMailDomain }

	### last signin ###
	$LAT = $Guest.employeeHireDate
	$lsi = $Guest.SignInActivity.LastSignInDateTime
	$lsiNI = $Guest.SignInActivity.LastNonInteractiveSignInDateTime
	$lsi_success = $Guest.SignInActivity.lastSuccessfulSignInDateTime
	$CreatedDate = (([DateTime]$Guest.CreatedDateTime).ToUniversalTime()).ToString("yyyy-MM-dd")
	$DaysSinceCreated = (New-TimeSpan -Start $createdDate -End $currentDate).Days
	$DaysSinceLastSignIn 	= $DaysSinceCreated
	$DaysSinceLastSignIn_NI = $DaysSinceCreated
	$DaysSinceLastSignIn_success = $DaysSinceCreated
	$DaysSinceLAT = $DaysSinceCreated
	if (-not (($null -eq $lsi) -or ($lsi -eq "0001-01-01T00:00:00Z"))) {
		$LastSignInDateTime	= [DateTime]$lsi
		$DaysSinceLastSignIn = (New-TimeSpan -Start $LastSignInDateTime.ToString("yyyy-MM-dd") -End $CurrentDate).Days
	}
	if (-not(($null -eq $lsiNI) -or ($lsiNI -eq "0001-01-01T00:00:00Z"))) {
		$LastSignInDateTime_NI	= [DateTime]$lsiNI
		$DaysSinceLastSignIn_NI = (New-TimeSpan -Start $LastSignInDateTime_NI.ToString("yyyy-MM-dd") -End $CurrentDate).Days
	}
	if (-not(($null -eq $lsi_success) -or ($lsi_success -eq "0001-01-01T00:00:00Z"))) {
		$LastSignInDateTime_success = [DateTime]$lsi_success
		$DaysSinceLastSignIn_success = (New-TimeSpan -Start $LastSignInDateTime_success.ToString("yyyy-MM-dd") -End $CurrentDate).Days
	}
	if (-not(($null -eq $LAT) -or ($LAT -eq "0001-01-01T00:00:00Z"))) {
		$LAT = [DateTime]$LAT
		$DaysSinceLAT = (New-TimeSpan -Start $LAT.ToString("yyyy-MM-dd") -End $CurrentDate).Days
	}
	### group membership count ###
	$UriResource = "users/$($Guest.id)/memberOf/`$Count"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
	Try {
		$GroupMemberCount = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ConsistencyLevel "eventual"
	}
	Catch {
		Write-Log "Critical Error: $($_.Exception.Message)" -MessageType Error
	}	
	
	if ($Guest.employeeHireDate) {
		$LastAuditTrace = [datetime]::Parse($Guest.employeeHireDate)
	}
	$GuestObject = [pscustomobject]@{
		UserId					= $Guest.id
		Mail 					= $Mail
		UserPrincipalName 		= $upn
		DisplayName 			= $Guest.DisplayName
		UserType 				= $Guest.UserType
		CreatedBy				= $Guest.employeeType
		CompanyName 			= $Guest.CompanyName
		Enabled					= $Guest.AccountEnabled
		T2TSync					= $T2TSync
		MailDomain 				= $mailDomain
		MSAGuestForAADOrg		= $MSAGuestForAADOrg
		ExtTenantId				= $ExtTenantId
		ExtTenantDisplayName	= $ExtTenantDisplayName
		ExtTenantDefaultDomain	= $ExtTenantDefaultDomain
		ExtTenantInboundSync	= $ExtTenantInboundSync
		UPNExtMail				= $strUPNExtMail
		UPNExtMailDomain		= $strUPNExtMailDomain
		otherMails				= $strOtherMails
		proxyAddresses			= $strProxyAddresses
		CreatedDateTime 		= [datetime]::Parse($Guest.CreatedDateTime)
		LastAuditTrace			= $LastAuditTrace
		LastSignIn				= $LastSignInDateTime
		LastSignIn_NI			= $LastSignInDateTime_NI
		LastSignIn_success		= $LastSignInDateTime_success
		DaysSinceLastSignIn		= $DaysSinceLastSignIn
		DaysSinceLastSignIn_NI 	= $DaysSinceLastSignIn_NI
		DaysSinceLastSignIn_success = $DaysSinceLastSignIn_success
		DaysSinceLAT			= $DaysSinceLAT
		DaysSinceCreated		= $DaysSinceCreated
		GroupMemberCount 		= $GroupMemberCount
		CreationType 			= $Guest.CreationType
		ExtUsrState 			= $Guest.ExternalUserState
		ExtUsrStateChangeDT		= $Guest.ExternalUserStateChangeDateTime
		issuerMSA				= $issuerMSA
		issuerExtAAD			= $issuerExtAAD
		issuerMail 				= $issuerMail
		showInAddressList		= $Guest.showInAddressList
		extAttr15				= $Guest.onPremisesExtensionAttributes.extensionAttribute15
	}
	$GuestListReport += $GuestObject

	if ($T2TSync) {
		$GuestObjectT2T = [pscustomobject]@{
			UserId					= $Guest.id
			Mail 					= $Mail
			UserPrincipalName 		= $upn
			DisplayName 			= $Guest.DisplayName
			UserType 				= $Guest.UserType
			Enabled					= $Guest.AccountEnabled
			CompanyName 			= $Guest.CompanyName
			Department				= $GuestT2T.Department
			JobTitle				= $GuestT2T.JobTitle
			Manager					= $GuestT2TManager.DisplayName
			ManagerMail				= $GuestT2TManager.Mail
			StreetAddress			= $GuestT2T.StreetAddress
			City					= $GuestT2T.City
			State					= $GuestT2T.State
			PostalCode				= $GuestT2T.PostalCode
			Country					= $GuestT2T.Country
			EmployeeId				= $GuestT2T.EmployeeId
			PreferredLanguage		= $GuestT2T.PreferredLanguage
			TelephoneNumber			= $GuestT2T.TelephoneNumber
			OfficeLocation			= $GuestT2T.OfficeLocation
			MobilePhone				= $GuestT2T.MobilePhone
			BusinessPhones			= $GuestT2T.BusinessPhones -join ";"

			MailDomain 				= $mailDomain
			ExtTenantId				= $ExtTenantId
			ExtTenantDisplayName	= $ExtTenantDisplayName
			ExtTenantDefaultDomain	= $ExtTenantDefaultDomain
			ExtTenantInboundSync	= $ExtTenantInboundSync
			UPNExtMail				= $strUPNExtMail
			UPNExtMailDomain		= $strUPNExtMailDomain
			otherMails				= $strOtherMails
			proxyAddresses			= $strProxyAddresses
			CreatedDateTime 		= [datetime]::Parse($Guest.CreatedDateTime)
			LastAuditTrace			= $LastAuditTrace
			LastSignIn				= $LastSignInDateTime
			LastSignIn_NI			= $LastSignInDateTime_NI
			DaysSinceLastSignIn		= $DaysSinceLastSignIn
			DaysSinceLastSignIn_NI 	= $DaysSinceLastSignIn_NI
			DaysSinceCreated		= $DaysSinceCreated
			GroupMemberCount 		= $GroupMemberCount
			CreationType 			= $Guest.CreationType
			ExtUsrStateChangeDT		= $Guest.ExternalUserStateChangeDateTime
			showInAddressList		= $Guest.showInAddressList
			extAttr15				= $Guest.onPremisesExtensionAttributes.extensionAttribute15
		}
		$GuestListReportT2T += $GuestObjectT2T
	}
}

Export-Report -Text "AAD guests report" -Report $GuestListReport -SortProperty "UserPrincipalName" -Path $OutputFile
Export-Report -Text "AAD guests report (T2T)" -Report $GuestListReportT2T -SortProperty "UserPrincipalName" -Path $OutputFileT2T

#######################################################################################################################

. $IncFile_StdLogEndBlock
