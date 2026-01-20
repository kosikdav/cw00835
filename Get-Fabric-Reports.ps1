#######################################################################################################################
# Get-Fabric-Reports
#######################################################################################################################


function Request-MSALToken {
	param (
		[int]$TTL = 20,
		$Certificate,
		[string]$ClientId,
		[string]$TenantId,
		[string]$Authority = "login.microsoftonline.com",
		[string]$Scope = "https://graph.microsoft.com/.default",
		[string]$Resource,
		[switch]$Silent,
		[switch]$Force
	)
	# main function body ##################################

	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

	$AuthorityURI = "https://$($Authority)/$($tenantId)"
	if ($Authority -eq "login.microsoftonline.com") {
		$tokenEndpoint = "$($AuthorityURI)/oauth2/v2.0/token"
	}
	if ($Authority -eq "login.windows.net") {
		$tokenEndpoint = "$($AuthorityURI)/oauth2/token"
	}
	
	$CertificateBase64Hash = [System.Convert]::ToBase64String($Certificate.GetCertHash())  
	$StartDate = (Get-Date "1970-01-01T00:00:00Z" ).ToUniversalTime()  
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
		client_id = $clientId
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
		$ExpiresOn = (Get-Date).AddSeconds($Token.expires_in)
		if (-not $Silent) {
			Write-Host "MSAL access token for $($TenantShortName) (app $($AppRegName)) $($Operation) - expires: $($ExpiresOn), TTL $($TTL)" -ForegroundColor DarkGray
		}
		return $Token.access_token		
	}
	Catch {
		Write-Host -String $_.Exception.Message -MessageType Error -ForceOnScreen -ForegroundColor "Red"
	}
}

#######################################################################################################################

$ClientId   = "151c09c7-73bc-4958-9819-173b7c07d0f6"
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$Thumbprint = "06fb98321befd4e9494aeea598eebc5b9fa24129"

$Certificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"

$AccessToken = Request-MSALToken -Certificate $Certificate -ClientId $ClientId -TenantId $TenantId -TTL 30 -Scope "https://analysis.windows.net/powerbi/api/.default"
$Headers = @{Authorization = "Bearer $($AccessToken)"}

$Uri = "https://api.powerbi.com/v1.0/myorg/admin/groups?`$filter=type+eq+'Workspace'&`$top=5000"
write-host $Uri

$Result = Invoke-RestMethod -Headers $Headers -Uri $Uri -ContentType "application/json"
$FabricGroups = $Result.value

foreach ($Group in $FabricGroups) {
	if (-not ($Group.Type.StartsWith("Personal"))) {
		write-host $Group.Name
		$Counter++
		#write-host $Group -ForegroundColor DarkGray
	}
}
write-host "Total Groups: $Counter" -ForegroundColor Cyan
write-host "Including personal: $($FabricGroups.Count)" -ForegroundColor Cyan	

$workspaceIds = @(
    '7edd39bd-8a19-4135-bb99-91657cc23e88'
)

$body = @{
	workspaces = $workspaceIds
}

$bodyJson = $body | ConvertTo-Json
$Result = $null
#$uri = "https://api.powerbi.com/v1.0/myorg/admin/workspaces/getInfo?lineage=True&datasourceDetails=True&datasetSchema=True&datasetExpressions=True&getArtifactUsers=True"
$uri = "https://api.powerbi.com/v1.0/myorg/admin/workspaces/getInfo"
#$Result = Invoke-RestMethod -Headers $Headers -Uri $Uri -Body $bodyJson -ContentType "application/json" -Method "Post"
$Result = Invoke-WebRequest -Headers $Headers -Uri $Uri -ContentType "application/json" -Method "POST" -Body $bodyJson
$FabricInfo = $Result | ConvertFrom-Json

write-host $FabricInfo -ForegroundColor Green


