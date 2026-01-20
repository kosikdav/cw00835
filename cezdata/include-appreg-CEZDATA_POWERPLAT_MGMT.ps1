###############################
#App name CEZ_POWERPLAT_MGMT
###############################
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName = "cezdata.onmicrosoft.com"
$ClientId   = "133eb721-64ad-4e3b-bdac-2e649b269bb0"
$Thumbprint = "1beec4a3caf63ed9d648780d1e0d639dfb357811"
$PwrEndPoint  = "prod"
$CertficateThumbprint = $Thumbprint
$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"
$Certificate = $ClientCertificate

<#
$appName = "CEZ_POWERPLAT_MGMT"
$certYears = 2
$certPassword = "123456"
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force

#>
