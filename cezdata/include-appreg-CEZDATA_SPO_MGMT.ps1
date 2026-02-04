###############################
# App name CEZ_SPO_MGMT
###############################
$AppName    = "CEZ_SPO_MGMT"
$ClientId   = "15e80ec1-456a-4aa5-8e7a-1cfb0422a393"
$Thumbprint = "935ba43436e8e127b767f9782b43c323c9f400ab"
$ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
$Certificate = $ClientCertificate
$Thumbprint = $ClientCertificate.Thumbprint
$CertficateThumbprint = $Thumbprint
$ApplicationId = $ClientId

<#
$appName = "CEZ_SPO_MGMT"
$certYears = 10
$certPassword = ""
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
