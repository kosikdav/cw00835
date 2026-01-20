#######################################################################################################################
# Get-ADO-Users
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder 			= "exports"
$LogFilePrefix		= "ado-users"

$OutputFolder 		= "ado\reports"
$OutputFilePrefix	= "add-users"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$OutputFile = New-OutputFile -RootFolder $REF -Folder $OutputFolder -Prefix $OutputFilePrefix -FileDateYesterday -Ext "csv"

##################################################################################################

. $IncFile_StdLogBeginBlock
$incFile = [System.IO.Path]::Combine($incFolder,"include-appreg-$($AppReg_LOG_READER).ps1")
. $incFile

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "projects"
$Organization = "cez-icts"
$Uri = New-ADOUri -Version "7.1" -Resource $UriResource -Organization $Organization

write-host $Uri

# Define parameters
$resource = "499b84ac-1321-427f-aa17-267ca6975798"

# Load the certificate
#$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $certPath, $certPassword

$cert = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"
$certificate = $cert

# Create the JWT header
$JWTheader = @{
    alg = "RS256";
    typ = "JWT";
    x5t = [System.Convert]::ToBase64String($cert.GetCertHash())
}

# Create the JWT payload
$StartDate = (Get-Date "1970-01-01T00:00:00Z" ).ToUniversalTime()  
$JWTExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End (Get-Date).ToUniversalTime().AddMinutes(60)).TotalSeconds  
$JWTExpiration = [math]::Round($JWTExpirationTimeSpan,0) 

$now = [DateTimeOffset]::Now.ToUnixTimeSeconds()
#$expiry = (Get-Date).AddMinutes(60).ToUniversalTime().ToUnixTimeSeconds()
$JWTpayload = @{
    aud = "https://login.microsoftonline.com/$tenantId/oauth2/token";
    iss = $clientId;
    sub = $clientId;
    jti = [Guid]::NewGuid().ToString();
    nbf = $now;
    exp = $JWTExpiration;
    resource = $resource
}

# Create the JWT token

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

# Request the access token
$response = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/token" -Method Post -Body @{
    grant_type = "client_credentials";
    client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer";
    client_assertion = $jwt;
    resource = $resource
}

# Output the access token
$response.access_token

$users = Get-GraphOutputREST -Uri $Uri -AccessToken $response.access_token -ContentType $ContentTypeJSON
write-host $users

#######################################################################################################################

. $IncFile_StdLogEndBlock
