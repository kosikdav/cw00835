###############################
#App name CEZDATA_EXO_MGMT
###############################
$AppName    = "CEZDATA_AIRPLUS_MGMT"
$ClientId   = "779e006e-420e-4042-a9aa-5ce1c1c181c2"
$TenantId   = "3687fd79-edff-4560-9dda-317079330262"
$TenantName = "airpluscz1.onmicrosoft.com"
$TenantShortName = "AIRPLUSCZ1"

$ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
if (-not $ClientCertificate) {
    $AppName = $AppName.Replace("CEZ_", "CEZDATA_")
    $ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
}
write-host "Thumbprint: $($ClientCertificate.Thumbprint)"
$Certificate = $ClientCertificate
$Thumbprint = $ClientCertificate.Thumbprint
$CertficateThumbprint = $Thumbprint
$ApplicationId = $ClientId

<#
$appName = "CEZDATA_EXO_MBX_MGMT"
$CertYears = 10
$CertPassword = "123456789"
$Password = ConvertTo-SecureString $CertPassword -AsPlainText -Force
$StartDate = (Get-Date).AddDays(-1)
$EndDate = (Get-Date).AddYears($CertYears)
Create-SelfSignedCertificate.ps1 -CommonName $AppName -StartDate $StartDate -EndDate $EndDate -Password $Password -Force
Write-Host $env:computername.toupper() -ForegroundColor Green
#>
