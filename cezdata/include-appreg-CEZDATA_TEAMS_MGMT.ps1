###############################
#App name CEZ_TEAMS_MGMT
###############################
$AppName    = "CEZ_TEAMS_MGMT"
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName = "cezdata.onmicrosoft.com"
$ClientId   = "d51e8332-64e4-4478-985c-43b3e60a99e7"
$Thumbprint = "605d5b70995f2f645afaeca13bb1a87fae2b414f"

$CertficateThumbprint = $Thumbprint
$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"
$Certificate = $ClientCertificate

<#
$appName = "CEZ_TEAMS_MGMT"
$certYears = 2
$certPassword = "21-!d12f5ygye1aF#$&#"
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
