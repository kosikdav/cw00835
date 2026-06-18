###############################
#App name CEZ_TEAMS_MGMT
###############################
$AppName    = "CEZ_TEAMS_MGMT"
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName = "cezdata.onmicrosoft.com"
$ClientId   = "d51e8332-64e4-4478-985c-43b3e60a99e7"
$Thumbprint = "605d5b70995f2f645afaeca13bb1a87fae2b414f"

$ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
if (-not $ClientCertificate) {
    $AppName = $AppName.Replace("CEZ_", "CEZDATA_")
    $ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
}

$Certificate = $ClientCertificate
$Thumbprint = $ClientCertificate.Thumbprint
$CertficateThumbprint = $Thumbprint
$ApplicationId = $ClientId

<#
$appName = "CEZ_TEAMS_MGMT"
$CertYears = 10
$CertPassword = "123456789"
$Password = ConvertTo-SecureString $CertPassword -AsPlainText -Force
$StartDate = (Get-Date).AddDays(-1)
$EndDate = (Get-Date).AddYears($CertYears)
Create-SelfSignedCertificate.ps1 -CommonName $AppName -StartDate $StartDate -EndDate $EndDate -Password $Password -Force
Write-Host $env:computername.toupper() -ForegroundColor Green  
#>
