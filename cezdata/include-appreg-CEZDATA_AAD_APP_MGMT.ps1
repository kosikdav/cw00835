###############################
# App name CEZ_AAD_APP_MGMT
###############################
$AppName = "CEZ_AAD_APP_MGMT"
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName = "cezdata.onmicrosoft.com"
$ClientId = "57761131-1c63-4959-85e2-8f7ffc7ed9e3"
$ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
$Certificate = $ClientCertificate
$Thumbprint = $ClientCertificate.Thumbprint
$CertficateThumbprint = $Thumbprint
$ApplicationId = $ClientId
write-host $Certificate -ForegroundColor Green

<#
$appName = "CEZ_AAD_APP_MGMT"
$certYears = 5
$certPassword = ""
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
