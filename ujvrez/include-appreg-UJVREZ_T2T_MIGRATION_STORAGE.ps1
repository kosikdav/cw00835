###############################
# App name CEZ_UJV_T2T_MIGRATION_STORAGE
###############################
$AppName            = "CEZ_UJV_T2T_MIGRATION_STORAGE"
$TenantId           = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName         = "cezdata.onmicrosoft.com"
$TenantShortName    = "CEZDATA"
$ClientId           = "c0b8e48f-aacc-4db1-bcae-8a7341ce436d"

$ClientCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
$Certificate = $ClientCertificate
$Thumbprint = $ClientCertificate.Thumbprint
$CertficateThumbprint = $Thumbprint
$ApplicationId = $ClientId
