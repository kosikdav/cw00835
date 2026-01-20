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

$OutputFolder		= "aad-signins"
$OutputFilePrefix	= "aad-signins"
$OutputFileSuffixNI	= "guests-ni"


#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$LogFileGAT = New-OutputFile -RootFolder $RLF -Folder $LogFolderGAT -Prefix $LogFilePrefixGAT -Ext "log"

$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -FileDateYesterday -Ext "csv"
$OutputFileNI = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixNI -FileDateYesterday -Ext "csv"

[array]$SignInReport = @()
[array]$SignInReportNI = @()
[hashtable]$AADGuest_DB = @{}
[array]$SignIns = @()
[array]$SignInsNI = @()
$start = $strYesterdayUTCStart
$end = $strYesterdayUTCEnd
$timeFrameLength = 15

function Get-SigninReportObject {
	param(
		[Parameter(Mandatory = $true)]$SignIn,
		[Parameter(Mandatory = $true)][string]$AccessToken
	)
	$appliedConditionalAccessPolicies = [string]::Empty
	$MFADetail = $networkLocationDetails = $signInEventTypes = $null
	$sessionLifetimePolicies = $resourceTenantName = $resourceTenantDomainName = $null
	$UPN_CEZDATA = $null
	$requestDirection = "unknown"

	if ($Signin.mfaDetail.authDetail) {
		$MFADetail = ($Signin.mfaDetail.authDetail).Replace("+","Tel:+")
	}
	
	if ($Signin.appliedConditionalAccessPolicies) {
		ForEach ($Policy in $Signin.appliedConditionalAccessPolicies) {
			if (($Policy.result -eq "success") -or ($Policy.result -eq "failure")) {
				try{
					$appliedConditionalAccessPolicies = $appliedConditionalAccessPolicies + $Policy.id + ":" + $Policy.result +";" 
				}
				catch {
					# In case of an error, we just skip this policy
					write-host $Policy -ForegroundColor Red
				}
			}
		}
		$appliedConditionalAccessPolicies = $appliedConditionalAccessPolicies.Trim(";")
	}
	
	if ($SignIn.networkLocationDetails) {
		ForEach ($Detail in $SignIn.networkLocationDetails) {
			foreach ($networkName in $Detail.networkNames) {
				$networkLocationDetails = $networkLocationDetails + $networkName +";"
			}
		}
		$networkLocationDetails = $networkLocationDetails.Trim(";") 
	}
	
	if ($Signin.signInEventTypes) {
		$signInEventTypes = $Signin.signInEventType -join ";"
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
		$failureReason = [string]::Empty
	}

	if ($Signin.resourceTenantId -eq $TenantId) {
		$resourceTenantName = "CEZDATA"
	}
	
	if ($SignIn.homeTenantId -eq $TenantId) {
		#request from home tenant
		$homeTenantName = "CEZDATA" 
		if ($SignIn.homeTenantId -eq $Signin.resourceTenantId) {
			$requestDirection = "internal"
		}
		else {
			$requestDirection = "outbound"
		}
		if ($AADEXTTenant_DB.ContainsKey($Signin.resourceTenantId)) {
			# If the tenant is in the DB, we use the information from the DB
			$resourceTenantName = $AADEXTTenant_DB[$Signin.resourceTenantId].displayName
			$resourceTenantDomainName = $AADEXTTenant_DB[$Signin.resourceTenantId].defaultDomainName
		}
		else {
			# If the tenant is not in the DB, we try to get the tenant information from the Graph API
			$UriResource = "tenantRelationships/findTenantInformationByTenantId"
			$UriEqualsParam = "tenantId='$($Signin.resourceTenantId)'"
			$Uri = New-GraphUri -Version "beta" -Resource $UriResource -EqualsParam $UriEqualsParam
			Try {
				$Query = Invoke-RestMethod -Headers @{ Authorization = "Bearer $AccessToken" } -Uri $Uri -Method "GET" -ContentType "application/json" -ErrorAction Stop
				$resourceTenantName = $Query.displayName
				$resourceTenantDomainName = $Query.defaultDomainName
			}
			Catch {
				#nothing
			}
		}
	}
	else {
		#request from external tenant - inbound request
		$requestDirection = "inbound"
		try {
			$homeTenantName = ($AADEXTTenant_DB[$SignIn.homeTenantId]).displayName
		}
		Catch {
			$homeTenantName = [string]::Empty
		}
	}
	
	$homeTenantId = $SignIn.homeTenantId
	if ($homeTenantId -eq $TenantId) {
		$homeTenantId = "b233f9e1"
	}
	
	$resourceTenantId = $SignIn.resourceTenantId
	if ($resourceTenantId -eq $TenantId) {
		$resourceTenantId = "b233f9e1"
	}

	if ($SignIn.userId -and ($AADGuest_DB.Contains($SignIn.userId))) {
		$UPN_CEZDATA = $AADGuest_DB[$SignIn.userId]
	}
	$userType = $Signin.userType
	if (($requestDirection -eq "outbound") -and ($Sigin.homeTenantId -eq $TenantId) -and ($null -eq $Signin.userType)) {
		$userType = "member"
	}
	try {
		$ReportObject = [PSCustomObject]@{
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
		homeTenantId					= $homeTenantId
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
	} 
	catch {
		# In case of an error, we just skip this sign-in
		write-host $Signin -ForegroundColor Red
		$ReportObject = $null
	}
	Return $ReportObject
}

#######################################################################################################################

. $IncFile_StdLogStartBlock

Write-Log "Getting all sign-in events from  $($start) to: $($end)"
Write-Log "Time frame length: $($timeFrameLength) minutes"

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

[datetime]$currentStart = $start
[datetime]$currentEnd = $currentStart.AddMinutes($timeFrameLength)
$StopwatchGraph =  [system.diagnostics.stopwatch]::StartNew()

do {
	Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 10
	[hashtable]$guestAuditRecords_DB = @{}
	[hashtable]$guestAuditRecordsNI_DB = @{}
	$strCurrentStart = $currentStart.ToString("yyyy-MM-ddTHH:mm:ssZ")
	$strCurrentEnd = $currentEnd.ToString("yyyy-MM-ddTHH:mm:ssZ")
	write-host "$($strCurrentStart) - $($strCurrentEnd) " -ForegroundColor DarkGray -NoNewline
	
	####################################################################
	# Process interactive sign-ins
	$UriResource = "auditLogs/signins"
	$UriFilter = "createdDateTime+ge+$($strCurrentStart)+and+createdDateTime+le+$($strCurrentEnd)"
	$Uri = New-GraphUri -Version "beta" -Resource $UriResource -Top 999 -Filter $UriFilter
	
	[array]$SignIns = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
	write-host "[" -ForegroundColor Cyan -NoNewline
	ForEach ($SignIn in $SignIns) {
		$SignInReport += Get-SigninReportObject -Signin $SignIn -accessToken $AuthDB[$AppReg_LOG_READER].AccessToken
		$inboundRequest = $false
		if ($Signin.resourceTenantId -eq $TenantId -and ($SignIn.homeTenantId -ne $Signin.resourceTenantId)) {
			$inboundRequest = $true
		}
		if ($Signin.userId -and ($Signin.userType -eq "Guest") -and $inboundRequest -and ($GuestAzureApps -contains $Signin.appId)) {
			Update-GuestAuditRecordDB -Id $Signin.userId -DateTime $Signin.createdDateTime -hashtableDB $guestAuditRecords_DB
		}
	}
	write-host "]" -ForegroundColor Cyan -NoNewline
	####################################################################
	# Process non-interactive sign-ins
	$UriResource = "auditLogs/signins"
	$UriFilter = "createdDateTime+ge+$($strCurrentStart)+and+createdDateTime+le+$($strCurrentEnd)+and+userType+eq+'Guest'+and+signInEventTypes/any(t: t eq 'nonInteractiveUser')"
	$Uri = New-GraphUri -Version "beta" -Resource $UriResource -Top 999 -Filter $UriFilter
	[array]$SignInsNI = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
	write-host "[" -ForegroundColor Green -NoNewline
	foreach ($Signin in $SignInsNI) {
		$SignInReportNI += Get-SigninReportObject -Signin $SignIn -accessToken $AuthDB[$AppReg_LOG_READER].AccessToken		
		if ($Signin.userId -and ($Signin.resourceTenantId -eq $TenantId) -and ($Signin.status.errorCode -eq 0)) {
			#Update-GuestAuditRecordDB -Id $Signin.userId -DateTime $Signin.createdDateTime -hashtableDB $guestAuditRecordsNI_DB
		}
	}
	write-host "] " -ForegroundColor Green -NoNewline
	
	$currentStart = $currentEnd
	$currentEnd = $currentStart.AddMinutes($timeFrameLength)
	if ($currentEnd -gt $end) {
		$currentEnd = $end
	}

	#######################################################################
	
	write-host "I:$($SignIns.Count.ToString('0000000')) ($($SignInReport.Count.ToString('0000000'))) NI:$($SignInsNI.Count.ToString('0000000')) ($($SignInReportNI.Count.ToString('0000000'))) elapsed:$($StopwatchGraph.Elapsed.Seconds)" -ForegroundColor White

	if (($guestAuditRecords_DB.Count -gt 0) -or ($guestAuditRecords_DB.Count -gt 0)) {
		Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
		Write-GuestAuditRecordDBToEntra -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -hashtableDB $guestAuditRecords_DB -EntraAttribute "employeeHireDate" -AuditType "AzureLogin" -LogFile $LogFileGAT
		Write-GuestAuditRecordDBToEntra -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -hashtableDB $guestAuditRecordsNI_DB -EntraAttribute "employeeHireDate" -AuditType "NILoginSuccess" -LogFile $LogFileGAT
	}
	
	Clear-Variable SignIns
	Clear-Variable SignInsNI
	Clear-Variable guestAuditRecords_DB
	Clear-Variable guestAuditRecordsNI_DB
	
} Until ($currentStart -eq $end)

#write guest audit records to Entra attribute "employeeHireDate"

Export-Report -Text "AAD users signin report" -Report $SignInReport -SortProperty "createdDateTime" -Path $OutputFile
Export-Report -Text "AAD users signin report NI" -Report $SignInReportNI -SortProperty "createdDateTime" -Path $OutputFileNI

#######################################################################################################################

. $IncFile_StdLogEndBlock