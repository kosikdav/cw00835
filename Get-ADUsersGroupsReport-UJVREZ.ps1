#######################################################################################################################
# Get-ADUsersGroupsReport-UJVREZ.ps1
#######################################################################################################################

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

<#
$AppName = "CEZ_UJV_T2T_MIGRATION_STORAGE"
$CertYears = 10
$CertPassword = "12345678"
$StartDate = (Get-Date).AddDays(-1)
$EndDate = (Get-Date).AddYears($CertYears)
$Password = ConvertTo-SecureString $CertPassword -AsPlainText -Force

Create-SelfSignedCertificate.ps1 -CommonName $AppName -StartDate $StartDate -EndDate $EndDate -Password $Password -Force
#>

$ADCredentialPathUsr = "c:\cred\qp_aad_grp_mgmt\qp_aad_grp_mgmt_qskosikdav.cred"
$ADCredentialPathSys = "c:\cred\qp_aad_grp_mgmt\qp_aad_grp_mgmt.cred"

$LogFile = "d:\logs\$ScriptName.log"

$AppName = "CEZ_UJV_T2T_MIGRATION_STORAGE"
$StorageAccount = "cezt2tstore"
$Container = "ujvrez"

$ClientId   = 'c0b8e48f-aacc-4db1-bcae-8a7341ce436d'
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"

$ADUserSearchBase  = "DC=ad,DC=ujv,DC=cz"
$ADGroupSearchBase  = "DC=ad,DC=ujv,DC=cz"

$ADUserProperties = @(
	'displayName',
	'mail',
	'distinguishedName',
	'samAccountName',
	'userPrincipalName',
	'objectGUID',
	'employeeId',
	'employeeNumber'
)

$ADGroupProperties = @(
	"DistinguishedName",
	"GroupCategory",
	"GroupScope",
	"Name",
	"mail",
	"ObjectGUID",
	"SamAccountName",
	"SID"
)

#######################################################################################################################
# variable initialization
#######################################################################################################################

$interactiveRun = [Environment]::UserInteractive

[array]$ReportUsr = @()
[array]$ReportGrpLst = @()
[array]$ReportGrpMem = @()

$UploadParams = @{
	StorageAccount = $StorageAccount
	Container = $container
}

if ($InteractiveRun) {
	$ADCredentialPath = $ADCredentialPathUsr
}
else {
	$ADCredentialPath = $ADCredentialPathSys
}

#######################################################################################################################
# function definitions
#######################################################################################################################

function Write-Log {
	[CmdletBinding(SupportsShouldProcess = $true)]
	param (
		[Parameter(Position=0)][string]$String,
		[string][ValidateSet("Info","Warning","Warn","Error","Err")]$MessageType = "Info"
	)
	# main function body ##################################
	$File = $script:LogFile
	$TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
	switch ($MessageType) {
		"Info"		{$LineType = "INFO"}
		"Warning" 	{$LineType = "WARN"}
		"Error" 	{$LineType = "ERR"}
		Default 	{$LineType = "INFO"}
	}
	$LinePrefix = $TimeStamp + " [" + ($LineType.PadRight(4," ")).ToUpper() + "] "
	Add-Content $File -Value ($LinePrefix + $String)
	Write-Host $String
}

function Request-MSALToken {
	param (
		[parameter(Mandatory = $true)]$Certificate,
		[parameter(Mandatory = $true)][string]$ClientId,
		[parameter(Mandatory = $true)][string]$TenantId,
		[int]$TTL = 20,
		[string]$Authority = "login.microsoftonline.com",
		[string]$Scope = "https://graph.microsoft.com/.default",
		[string]$Resource
	)
	# main function body ##################################

		[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
			
		$AuthorityURI = "https://$($Authority)/$($TenantId)"
		if ($Authority -eq "login.microsoftonline.com") {
			$tokenEndpoint = "$($AuthorityURI)/oauth2/v2.0/token"
		}
		if ($Authority -eq "login.windows.net") {
			$tokenEndpoint = "$($AuthorityURI)/oauth2/token"
		}
		
		$CertificateBase64Hash = [System.Convert]::ToBase64String($Certificate.GetCertHash())  
		$StartDate = (Get-Date "1970-01-01T00:00:00Z").ToUniversalTime()  
		$JWTExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End (Get-Date).ToUniversalTime().AddMinutes(60)).TotalSeconds  
		$JWTExpiration = [math]::Round($JWTExpirationTimeSpan,0)  
		$NotBeforeExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End ((Get-Date).ToUniversalTime())).TotalSeconds  
		$NotBefore = [math]::Round($NotBeforeExpirationTimeSpan,0)  
		$JWTHeader = @{  
			alg = "RS256"  
			typ = "JWT"  
			x5t = $CertificateBase64Hash -replace '\+','-' -replace '/','_' -replace '='  
		}  

		$JWTPayLoad = @{  
			aud = "$($AuthorityURI)/oauth2/token"  
			exp = $JWTExpiration  
			iss = $ClientId  
			jti = [guid]::NewGuid()  
			nbf = $NotBefore  
			sub = $ClientId  
		}  
		
		$JWTHeaderToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTHeader | ConvertTo-Json))  
		$EncodedHeader = [System.Convert]::ToBase64String($JWTHeaderToByte)  
		$JWTPayLoadToByte =  [System.Text.Encoding]::UTF8.GetBytes(($JWTPayload | ConvertTo-Json))  
		$EncodedPayload = [System.Convert]::ToBase64String($JWTPayLoadToByte)  
		$JWT = $EncodedHeader + "." + $EncodedPayload  
		$PrivateKey = ([System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate))  
		$RSAPadding = [Security.Cryptography.RSASignaturePadding]::Pkcs1  
		$HashAlgorithm = [Security.Cryptography.HashAlgorithmName]::SHA256  
		$Signature = [Convert]::ToBase64String($PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($JWT),$HashAlgorithm,$RSAPadding)) -replace '\+','-' -replace '/','_' -replace '='  
		$JWT = $JWT + "." + $Signature
		
		$body = @{
			client_id = $ClientId
			client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
			client_assertion = $JWT
			grant_type = "client_credentials"
		}
		if ($Resource) {
			$body.Add("resource",$Resource)
		} else {
			$body.Add("scope",$scope)
		}
		Try {
			$Token = Invoke-RestMethod -Uri $tokenEndpoint -Method "POST" -Body $body
		}
		Catch {
			Write-Log -String $_.Exception.Message -MessageType Error
		}
		return $Token.access_token
}


function Invoke-AzureBlobUpload {
	[CmdletBinding()]
	param (
		[string]$StorageAccount,
		[string]$Container,
		[string]$BlobName,
		[string]$Content,
		[string]$AccessToken
	)
	# main function body ##################################
	$Uri = "https://$StorageAccount.blob.core.windows.net/$Container/$BlobName"
	$Headers  = @{
		Authorization = "Bearer $AccessToken"
		"x-ms-version" = "2020-04-08"
		"x-ms-blob-type" = "BlockBlob"
		"Content-Type" = "text/csv"
	}
	try {
		Invoke-RestMethod -Uri $Uri -Method "PUT" -Headers $Headers -Body $Content
		Write-Log "Upload successful: $BlobName"
	}
	catch {
		Write-Log "Failed to upload blob: $_" -MessageType "ERR"
	}
}

#######################################################################################################################
# main script logic
#######################################################################################################################

Write-Log "--------------------------------------------------------------"
Write-Log "Script start"
Write-Log "Script file: $($ScriptPath)\$($ScriptName)"
If ($InteractiveRun) {
    Write-Log "Running interactively"
}
Else {
    Write-Log "Running non-interactively"
}
Write-Log "Log file: $($LogFile)"
Write-Log "ADCredentialPath: $($ADCredentialPath)"

$ADCredential = Import-Clixml -Path $ADCredentialPath

#read UJVREZ AD users
$GetADUserParams = @{
	Filter = '*'
	SearchBase = $ADUserSearchBase
	Properties = $ADUserProperties
	Credential = $ADCredential
}
$ADUsers = Get-ADUser @GetADUserParams | Select-Object $ADUserProperties
Write-Log "AD users: $($ADUsers.count)"

#read UJVREZ AD groups
$GetADGroupParams = @{
	Filter = '*'
	SearchBase = $ADGroupSearchBase
	Properties = $ADGroupProperties
	Credential = $ADCredential
}
$ADGroups = Get-ADGroup @GetADGroupParams | Select-Object $ADGroupProperties
Write-Log "AD groups: $($ADGroups.count)"

#process users
foreach ($user in $ADUsers) {
	$userObject = [PSCustomObject]@{
		userPrincipalName = $user.userPrincipalName
		mail = $user.mail
		displayName = $user.displayName
		samAccountName = $user.samAccountName
		distinguishedName = $user.distinguishedName
		objectGUID = $user.objectGUID
		employeeNumber = $user.employeeNumber
		employeeId = $user.employeeId
	}
	$ReportUsr += $userObject
}

#process groups and group members
foreach ($group in $ADGroups) {
	write-host "." -NoNewline
	$groupObject = [PSCustomObject]@{
		Name = $group.Name
		SamAccountName = $group.SamAccountName
		DistinguishedName = $group.DistinguishedName
		GroupCategory = $group.GroupCategory
		GroupScope = $group.GroupScope
		ObjectGUID = $group.ObjectGUID
		SID = $group.SID
	}
	$ReportGrpLst += $groupObject

	#group members - users
	$adUserParams = @{
		Filter     = "MemberOf -eq '$($group.DistinguishedName)'"
		Properties = $ADUserProperties
		Credential = $ADCredential
	}
	$GroupMembersUsr = Get-ADUser @adUserParams
	foreach ($member in $GroupMembersUsr) {
		$memberObject = [PSCustomObject]@{
			GroupSamAccountName = $group.SamAccountName
			GroupObjectGUID = $group.ObjectGUID
			MemberSamAccountName = $member.SamAccountName
			MemberObjectGUID = $member.objectGUID
			type = "user"
		}
		$ReportGrpMem += $memberObject
	}

	#group members - groups
	$adGroupParams = @{
		Filter     = "MemberOf -eq '$($group.DistinguishedName)'"
		Properties = $ADGroupProperties
		Credential = $ADCredential
	}
	$GroupMembersGrp = Get-ADGroup @adGroupParams
	foreach ($member in $GroupMembersGrp) {
		$memberObject = [PSCustomObject]@{
			GroupSamAccountName = $group.SamAccountName
			GroupObjectGUID = $group.ObjectGUID
			MemberSamAccountName = $member.SamAccountName
			MemberObjectGUID = $member.objectGUID
			type = "group"
		}
		$ReportGrpMem += $memberObject
	}
}
write-host
Write-Log "AD group memberships: $($ReportGrpMem.count)"

#get MSAL access token from Entra ID
$Certificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
$AccessToken = Request-MSALtoken -TenantId $TenantId -ClientId $ClientId -Certificate $Certificate -Scope "https://storage.azure.com/.default"
$UploadParams.AccessToken = $AccessToken

#create CSV content
$csvStringUsr = $ReportUsr | ConvertTo-Csv -NoTypeInformation -Delimiter ',' | Out-String
$csvStringGrpLst = $ReportGrpLst | ConvertTo-Csv -NoTypeInformation -Delimiter ',' | Out-String
$csvStringGrpMem = $ReportGrpMem | ConvertTo-Csv -NoTypeInformation -Delimiter ',' | Out-String

#upload CSV content to Azure Blob Storage
$Date = (Get-Date).ToString("yyMMdd")

$UploadParams.BlobName = "ADUsr_$Date.csv"
$UploadParams.Content = $csvStringUsr
Invoke-AzureBlobUpload @UploadParams

$UploadParams.BlobName = "ADGrpLst_$Date.csv"
$UploadParams.Content = $csvStringGrpLst
Invoke-AzureBlobUpload @UploadParams

$UploadParams.BlobName = "ADGrpMem_$Date.csv"
$UploadParams.Content = $csvStringGrpMem
Invoke-AzureBlobUpload @UploadParams

Write-Log "Script finish"
