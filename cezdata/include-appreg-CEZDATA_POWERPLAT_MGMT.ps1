###############################
#App name CEZ_POWERPLAT_MGMT
###############################
$appName = "CEZ_POWERPLAT_MGMT"
$ClientId   = "133eb721-64ad-4e3b-bdac-2e649b269bb0"
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName = "cezdata.onmicrosoft.com"
$TenantShortName = "CEZDATA"
$PwrEndPoint  = "prod"

$ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
$Certificate = $ClientCertificate
$Thumbprint = $ClientCertificate.Thumbprint
$CertficateThumbprint = $Thumbprint
$ApplicationId = $ClientId

<#
$appName = "CEZ_POWERPLAT_MGMT"
$certYears = 2
$certPassword = "123456"
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
