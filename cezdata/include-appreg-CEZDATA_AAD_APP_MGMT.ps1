###############################
# App name CEZ_AAD_APP_MGMT
###############################
$AppName = "CEZ_AAD_APP_MGMT"
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName = "cezdata.onmicrosoft.com"
$ClientId = "57761131-1c63-4959-85e2-8f7ffc7ed9e3"
$Thumbprint = "86c38c864525f6836be50b12f038cc91ceef4dcb"
$CertficateThumbprint = $Thumbprint
$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"
$Certificate = $ClientCertificate

<#
$appName = "CEZ_AAD_APP_MGMT"
$certYears = 5
$certPassword = ""
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
