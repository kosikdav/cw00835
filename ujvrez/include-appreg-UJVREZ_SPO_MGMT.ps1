###############################
# App name UJVERZ_SPO_MGMT
###############################
$AppName    = "UJVREZ_SPO_MGMT"
$ClientId   = "8d2cbe22-bb09-45c8-b2ab-d8cd1bf40ef2"
$TenantId   = "56b31968-ca9e-4cc3-9257-477c3699b885"
$ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
$Certificate = $ClientCertificate
$Thumbprint = $ClientCertificate.Thumbprint
$CertficateThumbprint = $Thumbprint
$ApplicationId = $ClientId

<#
$appName = "UJVREZ_SPO_MGMT"
$CertYears = 10
$CertPassword = "123456789"
$Password = ConvertTo-SecureString $CertPassword -AsPlainText -Force
$StartDate = (Get-Date).AddDays(-1)
$EndDate = (Get-Date).AddYears($CertYears)
Create-SelfSignedCertificate.ps1 -CommonName $AppName -StartDate $StartDate -EndDate $EndDate -Password $Password -Force
Write-Host $env:computername.toupper() -ForegroundColor Green

#>
