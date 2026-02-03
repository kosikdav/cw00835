###############################
#App name CEZ_EXO_MBX_MGMT
###############################
$AppName    = "CEZ_EXO_MBX_MGMT"
$ClientId   = "1f4e0528-c8a5-40da-babc-47e994e14454"
$ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
$Certificate = $ClientCertificate
$Thumbprint = $ClientCertificate.Thumbprint
$CertficateThumbprint = $Thumbprint
$ApplicationId = $ClientId
write-host $Certificate -ForegroundColor Green

<#
$appName = "CEZ_EXO_MBX_MGMT"
$certYears = 1
$certPassword = "xxxxxxxxxxxxxxxxxxxx"
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
