#######################################################################################################################
# Get-AAD-Groups-Audit
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder 					= "exports"
$LogFilePrefix				= "aad-apps-reports-"

$OutputFolder 				= "aad-apps\reports"
$OutputFilePrefix			= "aad-apps"

$OutputFileSuffixAppList	= "app-list"
$OutputFileSuffixSPList		= "sp-list"
$OutputFileSuffixPermApp	= "sp-perm-app"
$OutputFileSuffixCredApp	= "app-credentials"
$OutputFileSuffixCredSP		= "sp-credentials"

$OutputFileNameCredAppLtd	= "app-creds.csv"
$OutputFileNameCredSPLtd	= "sp-creds.csv"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputFileAppList 		= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixAppList -Ext "csv"
$OutputFileSPList 		= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixSPList -Ext "csv"
$OutputFileSPPermApp 	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixPermApp -Ext "csv"
$OutputFileCredApp 		= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixCredApp -Ext "csv"
$OutputFileCredSP		= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixCredSP -Ext "csv"

if ($TenantShortName -eq "CEZDATA") {
	$OutputFileCredAppLtd	= [System.IO.Path]::Combine($AADCredsFolder,$OutputFileNameCredAppLtd)
	$OutputFileCredSPLtd	= [System.IO.Path]::Combine($AADCredsFolder,$OutputFileNameCredSPLtd)
}

[array]$ReportAppList = @()
[array]$ReportSPList = @()
[array]$ReportSPPermApp = @()
[array]$ReportCredApp = @()
[array]$ReportCredSP = @()
[array]$ReportCredAppLtd = @()
[array]$ReportCredSPLtd = @()

[hashtable]$GraphPermissions_DB = @{}
$regexMail = '[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'

function Get-CredentialReportObject {
	[CmdletBinding()]
    param (
        [Parameter(Mandatory)]$Application,
		[Parameter(Mandatory)]$Credential,
		[Parameter(Mandatory)][string][ValidateSet("Password","Key")]$CredentialType,
		[switch]$LimitedDetails
    )
	# main function body ##################################
	$notesExtractedMails = $null
	if ($Application.notes) {
		$notesExtractedMails = [regex]::Matches($Application.notes, $regexMail) -join ";"
	}

	if ($CredentialType -eq "Password") {
		$Hint = [char]34+$Credential.hint+[char]34
		$customKeyIdentifier = $null
	}
	else {
		$Hint = $null
		$customKeyIdentifier = $Credential.customKeyIdentifier
	}
	if ($LimitedDetails) {
		$object = [pscustomobject]@{
			Id				= $Application.id; 
			AppId			= $Application.appId; 
			CredentialType 	= $CredentialType;
			keyId 			= $Credential.keyId;
			ExpiresOn 		= [DateTime]::Parse($Credential.endDateTime);
			DaysLeft 		= (New-TimeSpan -Start $script:CurrentDate -End $Credential.endDateTime).Days
		}
	}
	Else {
		$object = [pscustomobject]@{
			Id				= $Application.id; 
			AppId			= $Application.appId; 
			DisplayName		= $Application.displayName;
			CredentialType 	= $CredentialType;
			keyId 			= $Credential.keyId;
			Hint 			= $Hint;
			customKeyIdentifier	= $customKeyIdentifier;
			certType 		= $Credential.type;
			certUsage 		= $Credential.usage;
			CreatedOn 		= [DateTime]::Parse($Credential.startDateTime);
			ExpiresOn 		= [DateTime]::Parse($Credential.endDateTime);
			DaysLeft 		= (New-TimeSpan -Start $script:CurrentDate -End $Credential.endDateTime).Days
			Contacts		= $notesExtractedMails
		}
	}
	return $object
}

##################################################################################################

. $IncFile_StdLogBeginBlock

Write-Log "Getting AAD apps report as of: $(Get-Date -Format "yyyy-MM-dd HH:MM:ss")"
Write-Log "AAD Apps list ouput file:  $($OutputFileAppList)"
Write-Log "AAD SP list ouput file:  $($OutputFileSPList)"
Write-Log "AAD SP application perms:  $($OutputFileSPPermApp)"
Write-Log "AAD Apps credentials:  $($OutputFileCredApp)"
Write-Log "AAD SP credentials:  $($OutputFileCredSP)"

if ($TenantShortName -eq "CEZDATA") {
	Write-Log "AAD Apps credentials (limited details):  $($OutputFileCredAppLtd)"
	Write-Log "AAD SP credentials (limited details):  $($OutputFileCredSPLtd)"
}


$GraphPermissions_DB = Import-CSVtoHashDB -Path $DBFileAADPermissions -KeyName "id"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "applications"
$Uri = New-GraphUri -Version "beta" -Resource $UriResource
[array]$AADApplications = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "AAD applications" -ProgressDots

foreach ($Application in $AADApplications) {
	$notesExtractedMails = $null
	if ($Application.notes) {
		$notesExtractedMails = [regex]::Matches($Application.notes, $regexMail) -join ";"
		write-host $notesExtractedMails
	}
	
	if ($Application.passwordCredentials) {
        foreach ($Credential in $Application.passwordCredentials) {
			$ReportCredApp += Get-CredentialReportObject -Application $Application -Credential $Credential -CredentialType "Password"
			$ReportCredAppLtd += Get-CredentialReportObject -Application $Application -Credential $Credential -CredentialType "Password" -LimitedDetails
		}
	}

	if ($Application.keyCredentials) {
        foreach ($Credential in $Application.keyCredentials) {
			$ReportCredApp += Get-CredentialReportObject -Application $Application -Credential $Credential -CredentialType "Key"
			$ReportCredAppLtd += Get-CredentialReportObject -Application $Application -Credential $Credential -CredentialType "Key" -LimitedDetails
		}
	}
	
	$ReportAppList += [pscustomobject]@{
		Id				= $Application.id
		AppId			= $Application.appId
		DisplayName		= $Application.displayName
		CreatedDateTime = $Application.createdDateTime
		DeletedDateTime = $Application.deletedDateTime
		groupMembershipClaims = $Application.groupMembershipClaims
		isDeviceOnlyAuthSupported = $Application.isDeviceOnlyAuthSupported
		isFallbackPublicClient = $Application.isFallbackPublicClient
		notes = $Application.notes
		notesExtractedMails = $notesExtractedMails
		oauth2RequiredPostResponse = $Application.oauth2RequiredPostResponse
		samlMetadataUrl = $Application.samlMetadataUrl
		PublisherDomain	= $Application.PublisherDomain
		SignInAudience 	= $Application.signInAudience
		RequiredResourceAccess = $Application.requiredResourceAccess.resourceAppId -join ";"
		Web_HomepageUrl = $Application.web.homepageURL
		Web_LogoutUrl = $Application.web.logoutURL
	}
}

Export-Report -Text "AAD apps list" -Report $ReportAppList -Path $OutputFileAppList -SortProperty "displayName"
Export-Report -Text "AAD apps credentials" -Report $ReportCredApp -Path $OutputFileCredApp -SortProperty "displayName"
if ($TenantShortName -eq "CEZDATA") {
	Export-Report -Text "AAD apps credentials (limited details)" -Report $ReportCredAppLtd -Path $OutputFileCredAppLtd -SortProperty "displayName"
}


##################################################################################################
##################################################################################################

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "servicePrincipals"
$Uri = New-GraphUri -Version "beta" -Resource $UriResource
$AADServicePrincipals = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "AAD service principals" -ProgressDots
ForEach ($ServicePrincipal in $AADServicePrincipals) {
	Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

	switch ($ServicePrincipal.appOwnerOrganizationId) {
		"f8cdef31-a31e-4b4a-93e4-5f571e91255a" {$appOwnerOrganizationId = "Microsoft Services Tenant"}
		"b233f9e1-5599-4693-9cef-38858fe25406" {$appOwnerOrganizationId = "cezdata.onmicrosoft.com"}
		default {$appOwnerOrganizationId = $ServicePrincipal.appOwnerOrganizationId}
	}
	
	##############################################

	if ($ServicePrincipal.passwordCredentials) {
        foreach ($Credential in $ServicePrincipal.passwordCredentials) {
			$ReportCredSP += Get-CredentialReportObject -Application $ServicePrincipal -Credential $Credential -CredentialType "Password"
			$ReportCredSPLtd += Get-CredentialReportObject -Application $ServicePrincipal -Credential $Credential -CredentialType "Password" -LimitedDetails
		}
	}

	if ($ServicePrincipal.keyCredentials) {
        foreach ($Credential in $ServicePrincipal.keyCredentials) {
			$ReportCredSP += Get-CredentialReportObject -Application $ServicePrincipal -Credential $Credential -CredentialType "Key"			
			$ReportCredSPLtd += Get-CredentialReportObject -Application $ServicePrincipal -Credential $Credential -CredentialType "Key" -LimitedDetails
		}
	}

	##############################################

	$resourceSpecificApplicationPermissions = $null
	if ($null -ne $ServicePrincipal.resourceSpecificApplicationPermissions) {
		foreach ($resourceSpecificApplicationPermission in $ServicePrincipal.resourceSpecificApplicationPermissions) {
			if ($null -eq $resourceSpecificApplicationPermission) {
				$add = $resourceSpecificApplicationPermission.value
			} 
			else {
				$add = "|" + $resourceSpecificApplicationPermission.value
			}
			$resourceSpecificApplicationPermissions += $add
		}
	}
	
	$alternativeNames = $null
	if ($null -ne $ServicePrincipal.alternativeNames) {
		foreach ($alternativeName in $ServicePrincipal.alternativeNames) {
			if ($null -eq $alternativeNames) {
				$add = $alternativeName
			}
			else {
				$add = "|" + $alternativeName
			}
			$alternativeNames += $add
		}
	}
	$ReportSPList += [pscustomobject]@{
		Id										= $ServicePrincipal.id; 
		Enabled									= $ServicePrincipal.accountEnabled; 
		AppId									= $ServicePrincipal.appId; 
		DisplayName								= $ServicePrincipal.displayName;
		AppDisplayName							= $ServicePrincipal.appDisplayName;
		alternativeNames						= $alternativeNames;
		CreatedDateTime							= $ServicePrincipal.createdDateTime;
		Desription								= $ServicePrincipal.appDescription;
		PublisherName							= $ServicePrincipal.publisherName;
		appOwnerOrganizationId 					= $appOwnerOrganizationId;
		appRoleAssignmentRequired				= $ServicePrincipal.appRoleAssignmentRequired; 
		disabledByMicrosoftStatus				= $ServicePrincipal.disabledByMicrosoftStatus;
		isAuthorizationServiceEnabled			= $ServicePrincipal.isAuthorizationServiceEnabled; 
		isManagementRestricted					= $ServicePrincipal.isManagementRestricted;
		preferredSingleSignOnMode				= $ServicePrincipal.preferredSingleSignOnMode;
		servicePrincipalType					= $ServicePrincipal.servicePrincipalType; 
		signInAudience							= $ServicePrincipal.signInAudience;
		resourceSpecificApplicationPermissions 	= $resourceSpecificApplicationPermissions;
		deviceManagementAppType					= $ServicePrincipal.deviceManagementAppType  
	}
	
	##############################################

	$UriResource = "servicePrincipals/$($ServicePrincipal.id)/approleassignments"
	$Uri = New-GraphUri -Version "beta" -Resource $UriResource
	$SPApplicationPermissions = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
	foreach ($SPApplicationPermission in $SPApplicationPermissions) {
		$principal, $principaldisplayName, $appRoleName = $null
		if ($SPApplicationPermission.principalId) {
			$principal = Get-GraphUserById -id $SPApplicationPermission.principalId -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
			$principaldisplayName = $principal.displayName
			if (($null -eq $principal) -and $GraphError -and ($GraphErrorCode -eq 404)) {
				$principal = Get-GraphServicePrincipalById -id $SPApplicationPermission.principalId -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
				$principaldisplayName = $principal.appdisplayName
			}
			if ($SPApplicationPermission.appRoleId) {
				if ($GraphPermissions_DB.ContainsKey($SPApplicationPermission.appRoleId)) {
					$appRoleName = $GraphPermissions_DB[$SPApplicationPermission.appRoleId].Value
					#Write-Host $appRoleName
				}
			}
			
			$ReportSPPermApp += [pscustomobject]@{
				Id					= $ServicePrincipal.id; 
				AppId				= $ServicePrincipal.appId; 
				DisplayName			= $ServicePrincipal.displayName;
				AppDisplayName		= $ServicePrincipal.appDisplayName;
				PublisherName		= $ServicePrincipal.publisherName;
				creationTimeStamp	= $SPApplicationPermission.creationTimeStamp;
				principalId			= $SPApplicationPermission.principalId;
				principalUPN		= $principal.userPrincipalName;
				principalName		= $principaldisplayName;
				principalAppId		= $principal.appId;
				principalType		= $SPApplicationPermission.principalType;
				appRoleId			= $SPApplicationPermission.appRoleId;
				appRoleName			= $appRoleName;
				resourceId			= $SPApplicationPermission.resourceId;
				resourceName		= $SPApplicationPermission.resourceDisplayName
			}
		}
	}
}

#######################################################################################################################

Export-Report -Text "AAD SP list" -Report $ReportSPList -Path $OutputFileSPList -SortProperty "displayName"
Export-Report -Text "AAD SP application permissions" -Report $ReportSPPermApp -Path $OutputFileSPPermApp -SortProperty "displayName"
Export-Report -Text "AAD SP credentials" -Report $ReportCredSP -Path $OutputFileCredSP -SortProperty "displayName"
if ($TenantShortName -eq "CEZDATA") {
	Export-Report -Text "AAD SP credentials (limited details)" -Report $ReportCredSPLtd -Path $OutputFileCredSPLtd -SortProperty "displayName"
}


. $IncFile_StdLogEndBlock
