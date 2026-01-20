###############################
# App name UJV_LOG_READER
###############################
$AppName    = "UJV_LOG_READER"
$ClientId   = '71fa9a7c-9048-4561-92ed-5c852dbdf48d'
$TenantId   = "56b31968-ca9e-4cc3-9257-477c3699b885"
$TenantName = "ujvrez.onmicrosoft.com"
$Thumbprint = 'f4c9f4c9849d88a9bbc5a79abb9886bfe8176648'
$CertficateThumbprint = $Thumbprint
$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"
$Certificate = $ClientCertificate
$ApplicationId = $ClientId

<#
$appName = "UJV_LOG_READER"
$certYears = 5
$certPassword = "xx23xvsfgfgxxxxxx"
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
