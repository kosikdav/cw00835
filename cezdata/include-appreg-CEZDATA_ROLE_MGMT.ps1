###############################
#App name CEZ_ROLE_MGMT
###############################
$AppName    = "CEZ_ROLE_MGMT"
$ClientId   = "6f6da64c-0a56-42cb-83d5-640282fd88b4"
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName = "cezdata.onmicrosoft.com"
$Thumbprint = "83a42a10fe0a3fcd952b476cf98ef95dc9e93c55"
$CertficateThumbprint = $Thumbprint
$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"
$Certificate = $ClientCertificate
$ApplicationId = $ClientId

<#
$appName = "CEZ_ROLE_MGMT"
$certYears = 1
$certPassword = ""
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
