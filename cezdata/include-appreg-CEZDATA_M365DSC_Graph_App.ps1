###############################
# App name CEZ_M365DSC_Graph_App
###############################
$AppName    = "CEZ_M365DSC_Graph_App"
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName = "cezdata.onmicrosoft.com"
$ClientId   = "e1be2293-4cf4-4a22-aa67-ebeececc6f1d"
$Thumbprint = "517fb00c0e53c91be15bcf0ce91ca359481bfda1"
$CertificateThumbprint = $Thumbprint
$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"
$Certificate = $ClientCertificate

<#
$AppName    = "CEZ_M365DSC_Graph_App"
$certYears = 2
$certPassword = "eijuqehiquhxxx"
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
