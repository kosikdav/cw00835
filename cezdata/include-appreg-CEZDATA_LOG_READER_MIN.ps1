###############################
# App name CEZ_LOG_READER_MIN
###############################
$AppName    = "CEZ_LOG_READER_MIN"
$ClientId   = '91088521-ce70-4213-89aa-73b40b40f4cf'
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName = "cezdata.onmicrosoft.com"
$Thumbprint = 'd7ad4c2167fa7160c5cf54a8d62af28e40e60a9f'
$CertficateThumbprint = $Thumbprint
$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"
$Certificate = $ClientCertificate
$ApplicationId = $ClientId

<#
$appName = "CEZ_LOG_READER_MIN"
$certYears = 5
$certPassword = "xxxxxxxxx"
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
