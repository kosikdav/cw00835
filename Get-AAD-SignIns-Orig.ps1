#######################################################################################################################
# Get-AAD-SignIns
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "aad-signins"
$LogFilePrefix		= "aad-signins"

$LogFolderGAT		= "aad-guests-audit-trace"
$LogFilePrefixGAT	= "aad-guests-audit-trace"
$LogFileSuffixGAT	= "signins"

$OutputFolder		= "aad-signins"
$OutputFilePrefix	= "aad-signins"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$LogFileGAT = New-OutputFile -RootFolder $RLF -Folder $LogFolderGAT -Prefix $LogFilePrefixGAT -Suffix $LogFileSuffixGAT -Ext "log"

$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -FileDateYesterday -Ext "csv"

[array]$SignInReport = @()
[hashtable]$guestAuditRecords_DB = @{}
[hashtable]$AADGuest_DB = @{}

$start = $strYesterdayUTCStart
$end = $strYesterdayUTCEnd

#######################################################################################################################

. $IncFile_StdLogStartBlock

Write-Log "Getting all sign-in events from  $($start) to: $($end)"

$AADEXTTenant_DB = Import-CSVtoHashDB -Path $DBFileExtAADTenants -KeyName "tenantId"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriFilter = "userType+eq+'Guest'"
$UriSelect = "id,userPrincipalName"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Filter $UriFilter -Select $UriSelect
$Guests = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -text "AAD Guests" -ProgressDots
foreach ($Guest in $Guests) {
	$AADGuest_DB[$Guest.id] = $Guest.userPrincipalName
}

$UriResource = "auditLogs/signins"
$UriFilter = "createdDateTime+ge+$($start)+and+createdDateTime+le+$($end)"
$Uri = New-GraphUri -Version "beta" -Resource $UriResource -Top 99 -Filter $UriFilter
$SignIns = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ProgressDots -Text "AAD user signins"

ForEach ($SignIn in $Signins) {
	$appliedConditionalAccessPolicies = $MFADetail = $networkLocationDetails = $signInEventTypes = $null
	$sessionLifetimePolicies = $resourceTenantName = $resourceTenantDomainName = $null
	$UPN_CEZDATA = $null

	if ($Signin.mfaDetail.authDetail) {
		$MFADetail = ($Signin.mfaDetail.authDetail).Replace("+","Tel:+")
	}
	
	if ($Signin.appliedConditionalAccessPolicies) {
		ForEach ($Policy in $Signin.appliedConditionalAccessPolicies) {
			if (($Policy.result -eq "success") -or ($Policy.result -eq "failure")) {
				$appliedConditionalAccessPolicies = $appliedConditionalAccessPolicies + $Policy.id + ":" + $Policy.result +";" 
			}
		}
		if ($appliedConditionalAccessPolicies) {
			$appliedConditionalAccessPolicies = $appliedConditionalAccessPolicies.Trim(";")
		}
	}
	
	if ($SignIn.networkLocationDetails) {
		ForEach ($Detail in $SignIn.networkLocationDetails) {
			foreach ($networkName in $Detail.networkNames) {
				$networkLocationDetails = $networkLocationDetails + $networkName +";"
			}
		}
		if  ($networkLocationDetails) {
			$networkLocationDetails = $networkLocationDetails.Trim(";") 
		}
	}
	
	if ($Signin.signInEventTypes) {
		ForEach ($type in $Signin.signInEventTypes) {
			$signInEventTypes = $signInEventTypes + $type.ToLower() + ";"
		}
		if  ($signInEventTypes) {
			$signInEventTypes = $signInEventTypes.Trim(";") 
		}
	}

	if ($Signin.sessionLifetimePolicies) {
		ForEach ($policy in $Signin.sessionLifetimePolicies) {
			if ($policy.expirationRequirement -eq "signInFrequencyPeriodicReauthentication") {
				$sessionLifetimePolicies = 	"signInFrequencyPeriodicReauthentication"
			}
		}
	}

	$location = $Signin.location.city + ";" + $Signin.location.state
	$country = $Signin.location.countryOrRegion
	if ($Signin.status.errorCode -eq 0) {
		$status = "success"
	}
	else {
		$status = "failure"
	}
	$failureReason = $Signin.status.failureReason
	if ($failureReason -eq "Other.") {
		$failureReason = $null
	}

	if ($Signin.resourceTenantId -eq $TenantId) {
		$resourceTenantName = "CEZDATA"
	}
	$requestDirection = "inbound"
	if ($SignIn.homeTenantId -eq $TenantId) {
		$homeTenantName = "CEZDATA" 
		if ($SignIn.homeTenantId -eq $Signin.resourceTenantId) {
			$requestDirection = "internal"
		}
		else {
			$requestDirection = "outbound"
			Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
			$UriResource = "tenantRelationships/findTenantInformationByTenantId"
			$UriEqualsParam = "tenantId='$($Signin.resourceTenantId)'"
			$Uri = New-GraphUri -Version "beta" -Resource $UriResource -EqualsParam $UriEqualsParam
			Try {
				$Query = Invoke-RestMethod -Headers @{ Authorization = "Bearer $accessToken" } -Uri $Uri -Method "GET" -ContentType "application/json" -ErrorAction Stop
				$resourceTenantName = $Query.displayName
				$resourceTenantDomainName = $Query.defaultDomainName
			}
			Catch {
				#nothing
			}
		}
	}
	else {
		$homeTenantName = ($AADEXTTenant_DB[$SignIn.homeTenantId]).DisplayName
	}
	
	$homeTenantId = $SignIn.homeTenantId
	if ($homeTenantId -eq $TenantId) {
		$homeTenantId = "cezdata_id"
	}
	
	$resourceTenantId = $SignIn.resourceTenantId
	if ($resourceTenantId -eq $TenantId) {
		$resourceTenantId = "cezdata_id"
	}

	if ($AADGuest_DB.Contains($SignIn.userId)) {
		$UPN_CEZDATA = $AADGuest_DB[$SignIn.userId]
	}
	$userType = $Signin.userType
	if (($requestDirection -eq "outbound") -and ($Sigin.homeTenantId -eq $TenantId) -and ($null -eq $Signin.userType)) {
		$userType = "member"
	}
	$SignInReport += [PSCustomObject]@{
		Id								= $Signin.id
		CreatedDateTime         		= $Signin.createdDateTime
		CorrelationId                	= $Signin.correlationId
		OriginalRequestId            	= $Signin.originalRequestId

		#user & app
		UserId                     		= $Signin.userId
		UserDisplayName             	= $Signin.userDisplayName
		UserPrincipalName           	= $Signin.userPrincipalName
		UPN_CEZDATA						= $UPN_CEZDATA
		UPNDomain						= $Signin.userPrincipalName.Split("@")[1]
		userType						= $userType
		Status 							= $status
		requestDirection				= $requestDirection
		CAStatus						= $Signin.conditionalAccessStatus
		crossTenantAccessType			= $Signin.crossTenantAccessType
		IsInteractive                	= $Signin.isInteractive
		#ServicePrincipalName         	= $Signin.servicePrincipalName
		ServicePrincipalId           	= $Signin.servicePrincipalId
		ClientAppUsed                	= $Signin.clientAppUsed
		homeTenantId					= $SignIn.homeTenantId
		homeTenantName					= $homeTenantName

		#app
		AppId                    		= $Signin.appId
		AppDisplayName               	= $Signin.appDisplayName
		
		#ip address and location
		IpAddress                    	= $Signin.ipAddress
		ipAddressFromResourceProvider	= $Signin.ipAddressFromResourceProvider
		networkLocationDetails			= $networkLocationDetails
		location						= $location
		country							= $country

		#risk
		RiskDetail                   	= $Signin.riskDetail
		RiskLevelAggregated          	= $Signin.riskLevelAggregated
		RiskLevelDuringSignIn       	= $Signin.riskLevelDuringSignIn
		RiskState                    	= $Signin.riskState

		#resource
		ResourceDisplayName          	= $Signin.resourceDisplayName
		ResourceId                   	= $Signin.resourceId
		resourceTenantId				= $resourceTenantId
		resourceTenantName				= $resourceTenantName
		resourceTenantDomainName		= $resourceTenantDomainName
		resourceServicePrincipalId		= $Signin.resourceServicePrincipalId
		isTenantRestricted				= $Signin.isTenantRestricted

		#authentication
		appliedCAPolicies 				= $appliedConditionalAccessPolicies
		ErrorCode                    	= $Signin.status.errorCode
		FailureReason                	= $failureReason
		AdditionalDetails            	= $Signin.status.additionalDetails
		MFAMethod						= $Signin.mfaDetail.authMethod
		MFADetail						= $MFADetail
		#tokenIssuerName 				= $Signin.tokenIssuerName
		tokenIssuerType 				= $Signin.tokenIssuerType
		clientCredentialType			= $Signin.clientCredentialType
		authenticationRequirement		= $Signin.authenticationRequirement
		signInIdentifier				= $Signin.signInIdentifier
		signInIdentifierType			= $Signin.signInIdentifierType
		signInEventTypes				= $signInEventTypes
		#federatedCredentialId			= $Signin.federatedCredentialId
		#svcPrincipalCredKeyId 			= $Signin.servicePrincipalCredentialKeyId
		#svcPrincipalCredThumbprint 	= $Signin.servicePrincipalCredentialThumbprint
		uniqueTokenIdentifier 			= $Signin.uniqueTokenIdentifier
		incomingTokenType 				= $Signin.incomingTokenType
		authProtocol 					= $Signin.authenticationProtocol
		#authContextClassReferences 	= $authContextClassReferences
		#authProcessingDetails 			= $Signin.authenticationProcessingDetails
		#authDetails 					= $Signin.authenticationDetails
		#authRequirementPolicies 		= $Signin.authenticationRequirementPolicies
		#authAppDeviceDetails 			= $Signin.authenticationAppDeviceDetails
		#authAppPolicyEvaluationDetails = $Signin.authenticationAppPolicyEvaluationDetails
		sessionLifetimePolicies 		= $sessionLifetimePolicies
		
		#device
		DeviceId                     	= $Signin.deviceDetail.deviceId
		DeviceDisplayName            	= $Signin.deviceDetail.displayName
		DeviceOS						= $Signin.deviceDetail.operatingSystem
		DeviceBrowser                	= $Signin.deviceDetail.browser
		DeviceisCompliant            	= $Signin.deviceDetail.isCompliant
		DeviceisManaged              	= $Signin.deviceDetail.isManaged
		DevicetrustType              	= $Signin.deviceDetail.trustType

		#other
		ProcessingTimeInMilliseconds	= $Signin.processingTimeInMilliseconds
		UserAgent                    	= $Signin.userAgent
		#autonomousSystemNumber			= $Signin.autonomousSystemNumber
		#privateLinkDetails				= $Signin.privateLinkDetails
		#flaggedForReview				= $Signin.flaggedForReview
	}
	if (($Signin.userType -eq "Guest") -and ($requestDirection -eq "inbound") -and ($GuestAzureApps -contains $Signin.appId)) {
		#write-host "$($Signin.userPrincipalName) $($Signin.appDisplayName)"
		#write-host $Signin -ForegroundColor DarkGray
		#write-host "--------------------------------------------" -ForegroundColor White
		Update-GuestAuditRecordDB -Id $Signin.userId -DateTime $Signin.createdDateTime -hashtableDB $guestAuditRecords_DB
	}
}

#write guest audit records to Entra attribute "employeeHireDate"
Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
Write-GuestAuditRecordDBToEntra -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -hashtableDB $guestAuditRecords_DB -EntraAttribute "employeeHireDate" -LogFile $LogFileGAT

Export-Report -Text "AAD users signin report" -Report $SignInReport -SortProperty "createdDateTime" -Path $OutputFile

#######################################################################################################################

. $IncFile_StdLogEndBlock