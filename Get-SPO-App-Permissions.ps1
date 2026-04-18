#######################################################################################################################
# Set-SPO-App-Permissions
#######################################################################################################################
param(
    [parameter(Mandatory = $true)][string]$Url
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1
. $ScriptPath\include-Script-Start-Include.ps1

#######################################################################################################################

$SiteId = [string]::Empty
$RelSiteUrl = $RootUrl = $null
if ($Url.StartsWith("https://cezdata.sharepoint.com")) {
    $RelSiteUrl = $Url.Replace("https://cezdata.sharepoint.com","")
    $RootUrl = "cezdata.sharepoint.com"
}
else {
    if ($Url.StartsWith("https://cezdata-my.sharepoint.com")) {
        $RelSiteUrl = $Url.Replace("https://cezdata-my.sharepoint.com","")
        $RootUrl = "cezdata-my.sharepoint.com"
    }
}

if (-not $RelSiteUrl) {
    write-host "Url must start with https://cezdata.sharepoint.com or https://cezdata-my.sharepoint.com" -ForegroundColor Red
    Exit
}

Request-MSALToken -AppRegName $AppReg_SPO_MGMT -TTL 30

$UriResource = "sites/$($RootUrl):$RelSiteUrl"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
$Site = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_SPO_MGMT].AccessToken -ContentType $ContentTypeJSON

if (-not $Site) {
    write-host "Site not found: $Url" -ForegroundColor Red
    Exit
}

$SiteId = $Site.id
write-host "Site:        $Url ($($Site.name))" -ForegroundColor Yellow
write-host "Site id:     $SiteId" -ForegroundColor Gray

#existing permissions
$UriResource = "sites/$SiteId/permissions"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
[array]$ExistingPermissions = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_SPO_MGMT].AccessToken -ContentType $ContentTypeJSON
if ($ExistingPermissions.Count -gt 0) {
    foreach ($Permission in $ExistingPermissions) {
        write-host $string_divider
        write-host "Roles: $($Permission.roles -join ", ")" -ForegroundColor Green
        write-host "Application: $($Permission.grantedToIdentitiesV2.application.displayName)" -ForegroundColor Yellow
        write-host "Application id: $($Permission.grantedToIdentitiesV2.application.id)"
    }
    write-host $string_divider
}
