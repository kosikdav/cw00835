###############################
# App name CEZ_ICTS_Fabric_Mgmt
###############################
$AppName    = "CEZ_ICTS_Fabric_Mgmt"
$ClientId   = '99734d0b-2375-472a-8ecb-65af0d6923c5'
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName = "cezdata.onmicrosoft.com"
$ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
$Certificate = $ClientCertificate
$Thumbprint = $ClientCertificate.Thumbprint
$CertficateThumbprint = $Thumbprint
$ApplicationId = $ClientId

<#
$appName = "CEZ_ICTS_Fabric_Mgmt"
$certYears = 1
$certPassword = ""
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
