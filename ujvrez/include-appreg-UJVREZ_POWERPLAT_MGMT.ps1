###############################
#App name CEZ_POWERPLAT_MGMT
###############################
$AppName = "UJVREZ_POWERPLAT_MGMT"
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
$AppName = "UJVREZ_POWERPLAT_MGMT"
$CertYears = 10
$CertPassword = "123456789"
$Password = ConvertTo-SecureString $CertPassword -AsPlainText -Force
$StartDate = (Get-Date).AddDays(-1)
$EndDate = (Get-Date).AddYears($CertYears)
Create-SelfSignedCertificate.ps1 -CommonName $AppName -StartDate $StartDate -EndDate $EndDate -Password $Password -Force
Write-Host $env:computername.ToUpper() -ForegroundColor Green  
#>
