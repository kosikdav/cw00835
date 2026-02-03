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
$CertSubject        = "L=CEZDATA, S=CZ, C=CZ, CN=CEZ_SPO_REPORTS_PnP"

$ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq $CertSubject }
$Certificate = $ClientCertificate
$Thumbprint = $ClientCertificate.Thumbprint
$CertficateThumbprint = $Thumbprint
$ApplicationId = $ClientId

<#
$Password = "P@ssw0rd"
$SecPassword = ConvertTo-SecureString -String $Password -AsPlainText -Force

$Props = @{
    Outpfx              = "CEZ_SPO_REPORTS_PnP.pfx" 
    ValidYears          = 15
    CertificatePassword = $SecPassword 
    CommonName          = "CEZ_SPO_REPORTS_PnP" 
    Country             = "CZ" 
    State               = "CZ"
    Locality            = "CEZDATA"
}

$Cert = New-PnPAzureCertificate @Props
#>

