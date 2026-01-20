###############################
# App name CEX_REPORT_READER
###############################
$AppName    = "CEX_REPORT_READER"
$TenantId   = "e85940c5-a691-4da1-922a-ca217be4685c"
$TenantName = "cezext.onmicrosoft.com"
$ClientId   = "e9d8cd22-6da7-4922-aad7-457da98575f6"
$Thumbprint = "c6efe92b6015d6d7d0ac6aa1101cd4633357f6da"
$CertficateThumbprint = $Thumbprint
$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"
$Certificate = $ClientCertificate

<#
$appName = "CEX_REPORT_READER"
$certYears = 1
$certPassword = "xxxxxxxxxxxxxxxxxxxx"
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
