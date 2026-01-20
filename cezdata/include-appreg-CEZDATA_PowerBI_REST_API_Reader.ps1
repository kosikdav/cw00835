###############################
# App name CEZ_PowerBI_REST_API_Reader
###############################
$AppName    = "CEZ_PowerBI_REST_API_Reader"
$ClientId   = "151c09c7-73bc-4958-9819-173b7c07d0f6"
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName = "cezdata.onmicrosoft.com"
$Thumbprint = "06fb98321befd4e9494aeea598eebc5b9fa24129"
$CertficateThumbprint = $Thumbprint
$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"
$Certificate = $ClientCertificate
$ApplicationId = $ClientId

<#
$appName = "CEZ_PowerBI_REST_API_Reader"
$certYears = 5
$certPassword = "skjsuehwirvi3niu"
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>

