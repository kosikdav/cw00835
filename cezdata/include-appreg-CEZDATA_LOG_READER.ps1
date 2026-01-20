###############################
# App name CEZ_LOG_READER
###############################
$AppName    = "CEZ_LOG_READER"
$ClientId   = 'd57d8058-3905-4894-be94-fc25429c1579'
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName = "cezdata.onmicrosoft.com"
$Thumbprint = '3c99134b869043425c5b07df961e94e5239e9127'
$CertficateThumbprint = $Thumbprint
$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"
$Certificate = $ClientCertificate
$ApplicationId = $ClientId


<#
$appName = "CEZ_LOG_READER"
$certYears = 5
$certPassword = "xxxxxxxxxxxx"
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
