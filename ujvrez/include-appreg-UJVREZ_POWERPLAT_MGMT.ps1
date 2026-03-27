###############################
#App name CEZ_POWERPLAT_MGMT
###############################
$appName = "UJVREZ_POWERPLAT_MGMT"
$ClientId   = "e40d92cb-7021-4683-a32c-5576a8896540"
$TenantId   = "56b31968-ca9e-4cc3-9257-477c3699b885"
$TenantName = "ujvrez.onmicrosoft.com"
$TenantShortName = "UJVREZ"
$PwrEndPoint  = "prod"

$ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
$Certificate = $ClientCertificate
$Thumbprint = $ClientCertificate.Thumbprint
$CertficateThumbprint = $Thumbprint
$ApplicationId = $ClientId

<#
$appName = "UJVREZ_POWERPLAT_MGMT"
$certYears = 5
$certPassword = "123456"
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
