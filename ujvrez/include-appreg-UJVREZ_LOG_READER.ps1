###############################
# App name UJV_LOG_READER
###############################
$AppName    = "UJV_LOG_READER"
$ClientId   = '71fa9a7c-9048-4561-92ed-5c852dbdf48d'
$TenantId   = "56b31968-ca9e-4cc3-9257-477c3699b885"
$TenantName = "ujvrez.onmicrosoft.com"

$ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
$Certificate = $ClientCertificate
$Thumbprint = $ClientCertificate.Thumbprint
$CertficateThumbprint = $Thumbprint
$ApplicationId = $ClientId

<#
$appName = "UJV_LOG_READER"
$certYears = 5
$certPassword = ""
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
