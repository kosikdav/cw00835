###############################
#App name CEZ_EXO_MBX_MGMT
###############################
$AppName    = "CEZ_EXO_MBX_MGMT"
$ClientId   = "1f4e0528-c8a5-40da-babc-47e994e14454"
$Thumbprint = "95091439108119bc724317a7ed6e0bdc52fd379a"
$CertficateThumbprint = $Thumbprint
$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"
$Certificate = $ClientCertificate

<#
$appName = "CEZ_EXO_MBX_MGMT"
$certYears = 1
$certPassword = "xxxxxxxxxxxxxxxxxxxx"
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
