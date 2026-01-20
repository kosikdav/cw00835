###############################
# App name CEZ_SPO_CPR_DOWNLOAD
###############################
$AppName    = "CEZ_SPO_CPR_DOWNLOAD"
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName = "cezdata.onmicrosoft.com"
$ClientId   = "bcaea0f6-1377-4d54-9988-aa15a79d040f"
$Thumbprint = "b03da5d347298551a034f1c1c27649f6c1af905f"
$CertificateThumbprint = $Thumbprint
$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"
$Certificate = $ClientCertificate

<#
$appName = "CEZ_SPO_CPR_DOWNLOAD"
$certYears = 1
$certPassword = "xxx"
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
