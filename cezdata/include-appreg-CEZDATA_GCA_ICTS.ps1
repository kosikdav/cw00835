###############################
# App name CEZ_Copilot_Graph_Connector_Agent_ICTS
###############################
$AppName    = "CEZ_Copilot_Graph_Connector_Agent_ICTS"
$ClientId   = '67f4a989-db76-42eb-a90d-ef039d5ea710'
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName = "cezdata.onmicrosoft.com"
$TenantShortName = "CEZDATA"

$ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
$Certificate = $ClientCertificate
$Thumbprint = $ClientCertificate.Thumbprint
$CertficateThumbprint = $Thumbprint
$ApplicationId = $ClientId

<#
$appName = "CEZ_Copilot_Graph_Connector_Agent_ICTS"
$certYears = 10
$certPassword = "P@ssw0rd"
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
