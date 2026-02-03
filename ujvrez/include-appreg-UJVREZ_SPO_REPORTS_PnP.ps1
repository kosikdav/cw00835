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
$CertSubject        = "L=CEZDATA, S=CZ, C=CZ, CN=UJV_SPO_REPORTS_PnP"

$ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq $CertSubject }
$Certificate = $ClientCertificate
$Thumbprint = $ClientCertificate.Thumbprint
$CertficateThumbprint = $Thumbprint
$ApplicationId = $ClientId

<#
$Password = "P@ssw0rd"
$SecPassword = ConvertTo-SecureString -String $Password -AsPlainText -Force

$Props = @{
    Outpfx              = "UJV_SPO_REPORTS_PnP.pfx" 
    ValidYears          = 15
    CertificatePassword = $SecPassword 
    CommonName          = "UJV_SPO_REPORTS_PnP" 
    Country             = "CZ" 
    State               = "CZ"
    Locality            = "CEZDATA"
}

$Cert = New-PnPAzureCertificate @Props
#>

