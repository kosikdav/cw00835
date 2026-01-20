###############################
# App name CEZ_AAD_USR_MGMT
###############################
$AppName = "CEZ_AAD_USR_MGMT"
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName = "cezdata.onmicrosoft.com"
$ClientId = "3c8784e5-5a20-4cf6-8db5-f5eb98e3c1b7"
$Thumbprint = "04e326fdf889eeafe1c074a10dd3fc5e0e9bf5a4"
$CertficateThumbprint = $Thumbprint
$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"
$Certificate = $ClientCertificate

<#
$appName = "CEZ_AAD_USR_MGMT"
$certYears = 5
$certPassword = "xxx"
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
