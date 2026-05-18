###############################
# App name CEZ_PowerBI_REST_API_Reader
###############################
$AppName    = "CEZ_PowerBI_REST_API_Reader"
$ClientId   = "151c09c7-73bc-4958-9819-173b7c07d0f6"
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName = "cezdata.onmicrosoft.com"
$ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
$Certificate = $ClientCertificate
$Thumbprint = $ClientCertificate.Thumbprint
$CertficateThumbprint = $Thumbprint
$ApplicationId = $ClientId

<#
$appName = "CEZ_PowerBI_REST_API_Reader"
$CertYears = 10
$CertPassword = "123456789"
$Password = ConvertTo-SecureString $CertPassword -AsPlainText -Force
$StartDate = (Get-Date).AddDays(-1)
$EndDate = (Get-Date).AddYears($CertYears)
Create-SelfSignedCertificate.ps1 -CommonName $AppName -StartDate $StartDate -EndDate $EndDate -Password $Password -Force
Write-Host $env:computername.toupper() -ForegroundColor Green  
#>

