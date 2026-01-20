###############################
# App name CEZ_ICTS_Fabric_Mgmt
###############################
$AppName    = "CEZ_ICTS_Fabric_Mgmt"
$ClientId   = '99734d0b-2375-472a-8ecb-65af0d6923c5'
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName = "cezdata.onmicrosoft.com"
$Thumbprint = 'f1eb1a07186a940e378494d33eb46a297634e46d'
$CertficateThumbprint = $Thumbprint
$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"
$Certificate = $ClientCertificate
$ApplicationId = $ClientId

<#
$appName = "CEZ_ICTS_Fabric_Mgmt"
$certYears = 1
$certPassword = ""
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
