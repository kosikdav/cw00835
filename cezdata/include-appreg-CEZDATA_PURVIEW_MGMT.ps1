###############################
#App name CEZDATA_PURVIEW_MGMT
###############################
$AppName    = "CEZ_PURVIEW_MGMT"
$ClientId   = "6795ed13-df56-486b-9855-124c90954316"
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName = "cezdata.onmicrosoft.com"
$TenantShortName = "CEZDATA"

$ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
$Certificate = $ClientCertificate
$Thumbprint = $ClientCertificate.Thumbprint
$CertficateThumbprint = $Thumbprint
$ApplicationId = $ClientId

<#
$appName = "CEZ_PURVIEW_MGMT"
$CertYears = 10
$CertPassword = "123456789"
$Password = ConvertTo-SecureString $CertPassword -AsPlainText -Force
$StartDate = (Get-Date).AddDays(-1)
$EndDate = (Get-Date).AddYears($CertYears)
Create-SelfSignedCertificate.ps1 -CommonName $AppName -StartDate $StartDate -EndDate $EndDate -Password $Password -Force
Write-Host $env:computername.toupper() -ForegroundColor Green
#>
