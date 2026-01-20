#######################################################################################################################
# Set-SPO-App-Permissions
#######################################################################################################################
param(
    [parameter(Mandatory = $true)][string]$Url,
	[ValidateSet("read","write","fullcontrol","manage")]$Role="read",
    [switch]$Delete,
    [switch]$Force
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1
. $ScriptPath\include-Script-StdIncBlock.ps1
. $ScriptPath\include-Functions-Common.ps1

#######################################################################################################################

$SiteId = [string]::Empty

Request-MSALToken -AppRegName $AppReg_SPO_MGMT -TTL 30
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

#read all sites from SPO and check if site exists
$UriResource = "sites"
$UriSelect = "id,webUrl,name"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect -Top 999
[array]$Sites = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_SPO_MGMT].AccessToken -ContentType $ContentTypeJSON -Text "Getting SPO sites" -ProgressDots
foreach ($Site in $Sites) {
	if ($Site.webUrl -eq $Url) {
        $SiteId = $Site.id
        write-host "Site:        $Url ($($Site.name))" -ForegroundColor Yellow
        write-host "Site id:     $SiteId" -ForegroundColor Gray
    }
}
if ($SiteId -eq [string]::Empty) {
	write-host "Site $Url not found" -ForegroundColor Red
	Exit
}
#existing permissions
$UriResource = "sites/$SiteId/permissions"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
[array]$ExistingPermissions = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_SPO_MGMT].AccessToken -ContentType $ContentTypeJSON
if ($ExistingPermissions.Count -gt 0) {
    foreach ($Permission in $ExistingPermissions) {
        write-host "$($Permission.grantedToIdentitiesV2.application.id) $($Permission.grantedToIdentitiesV2.application.displayName)" -ForegroundColor DarkYellow
        write-host $Permission.grantedToIdentities
    }
}
