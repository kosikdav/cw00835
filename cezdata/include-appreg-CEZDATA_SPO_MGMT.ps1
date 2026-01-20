###############################
# App name CEZ_SPO_MGMT
###############################
$script:AppName    = "CEZ_SPO_MGMT"
$script:ClientId   = "15e80ec1-456a-4aa5-8e7a-1cfb0422a393"
$script:Thumbprint = "935ba43436e8e127b767f9782b43c323c9f400ab"
$script:CertificateThumbprint = $script:Thumbprint
$script:ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($script:Thumbprint)"
$script:Certificate = $script:ClientCertificate

<#
$appName = "CEZ_SPO_MGMT"
$certYears = 10
$certPassword = ""
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>
