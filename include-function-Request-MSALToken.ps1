########################################################################################
# REQUEST-MSALTOKEN
########################################################################################
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
		
		#$client_assertion = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($Certificate))

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
			Write-Log -String $_.Exception.Message -MessageType Error -ForceOnScreen -ForegroundColor "Red"
		}
		return $Token.access_token
}
