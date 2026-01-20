###############################
# App name CEZ_SPO_REPORTS_PnP
###############################
$AppName            = "CEZ_SPO_REPORTS_PnP"
$TenantId           = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName         = "cezdata.onmicrosoft.com"
$TenantShortName    = "CEZDATA"
$RootSPOURL         = "https://cezdata.sharepoint.com/sites"
$PnPURL             = "https://cezdata.sharepoint.com"
$PnPTenant          = "cezdata.onmicrosoft.com"
$TenantAdminURL     = "https://cezdata-admin.sharepoint.com"
$ClientId           = "4488584d-a49a-42f3-83d1-aa19ae7e787f"

$Thumbprint = "8028a32475ab839f7c7eb5ced246880ef9afb2bf"
$CertificateThumbprint = $Thumbprint
$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"
$Certificate = $ClientCertificate

<#
$Password = "P@ssw0rd"
$SecPassword = ConvertTo-SecureString -String $Password -AsPlainText -Force

$Props = @{
    Outpfx              = "CEZ_SPO_REPORTS_PnP.pfx" 
    ValidYears          = 15
    CertificatePassword = $SecPassword 
    CommonName          = "CEZ_SPO_REPORTS_PnP" 
    Country             = "CZ" 
    State               = "Prague"
    Locality            = "cw00835po365log"
}

$Cert = New-PnPAzureCertificate @Props
#>

