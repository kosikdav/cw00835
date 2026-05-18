###############################
# App name UJVREZ_LOG_READER
###############################
$AppName    = "UJV_LOG_READER"
$ClientId   = '71fa9a7c-9048-4561-92ed-5c852dbdf48d'
$TenantId   = "56b31968-ca9e-4cc3-9257-477c3699b885"
$TenantName = "ujvrez.onmicrosoft.com"
$TenantShortName = "UJVREZ"

$ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
$Certificate = $ClientCertificate
$Thumbprint = $ClientCertificate.Thumbprint
$CertficateThumbprint = $Thumbprint
$ApplicationId = $ClientId

<#
$appName = "UJVREZ_LOG_READER"
$CertYears = 10
$CertPassword = "123456789"
$Password = ConvertTo-SecureString $CertPassword -AsPlainText -Force
$StartDate = (Get-Date).AddDays(-1)
$EndDate = (Get-Date).AddYears($CertYears)
Create-SelfSignedCertificate.ps1 -CommonName $AppName -StartDate $StartDate -EndDate $EndDate -Password $Password -Force
Write-Host $env:computername.toupper() -ForegroundColor Green  
#>
