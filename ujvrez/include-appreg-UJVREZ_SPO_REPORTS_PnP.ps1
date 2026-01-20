###############################
# App name UJV_SPO_REPORTS_PnP
###############################
$AppName            = "UJV_SPO_REPORTS_PnP"
$TenantId           = "56b31968-ca9e-4cc3-9257-477c3699b885"
$TenantName         = "ujvrez.onmicrosoft.com"
$TenantShortName    = "UJVREZ"
$RootSPOURL         = "https://ujvrez.sharepoint.com/sites"
$PnPURL             = "https://ujvrez.sharepoint.com"
$PnPTenant          = "ujvrez.onmicrosoft.com"
$TenantAdminURL     = "https://ujvrez-admin.sharepoint.com"
$ClientId           = "22b439c7-c86f-45e6-a0dc-7eac35ba8e80"

$Thumbprint = "75fbf4c9743cdcafba53bce12863bae8323f8f88"
$CertificateThumbprint = $Thumbprint
$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"
$Certificate = $ClientCertificate

<#
$Password = "P@ssw0rd"
$SecPassword = ConvertTo-SecureString -String $Password -AsPlainText -Force

$Props = @{
    Outpfx              = "UJV_SPO_REPORTS_PnP.pfx" 
    ValidYears          = 15
    CertificatePassword = $SecPassword 
    CommonName          = "UJV_SPO_REPORTS_PnP" 
    Country             = "CZ" 
    State               = "Prague"
    Locality            = "cw00835po365log"
}

$Cert = New-PnPAzureCertificate @Props
#>

