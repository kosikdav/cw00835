#######################################################################################################################
# Get-SPO-Permissions
#######################################################################################################################
param(
    [parameter(Mandatory = $true)][string]$Url,
	[parameter(Mandatory = $true)][string][ValidateSet("read","write")]$Role,
	[parameter(Mandatory = $true)][string]$ApplicationId
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1
. $ScriptPath\include-Script-StdIncBlock.ps1

#######################################################################################################################

Request-MSALToken -AppRegName $AppReg_SPO_MGMT -TTL 30
$UriResource = "sites"
$UriSelect = "id,webUrl"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect -Top 999
$Sites = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_SPO_MGMT].AccessToken -ContentType $ContentTypeJSON -Text "Getting SPO sites" -ProgressDots
foreach ($Site in $Sites) {
	$Site_DB.Add($Site.webUrl,$Site.id)
}

if $Site_DB.ContainsKey($Url) {
	$SiteId = $Site_DB[$Url]
} else {
	write-host "Site $Url not found" -ForegroundColor Red
	Exit
}

$UriResource = "applications(appId='$ApplicationId')"
$UriSelect = "id,appId,displayName"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
$ApplicationByAppId = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_SPO_MGMT].AccessToken -ContentType $ContentTypeJSON -Text "Getting SPO sites" -ProgressDots
$AppName = $ApplicationByAppId.displayName
if ($null -eq $ApplicationByAppId) {
	$UriResource = "applications/$ApplicationId"
	$UriSelect = "id,appId,displayName"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
	$ApplicationByObjId = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_SPO_MGMT].AccessToken -ContentType $ContentTypeJSON -Text "Getting SPO sites" -ProgressDots
	$AppName = $ApplicationByObjId.displayName
	if ($null -eq $ApplicationByObjId) {
		write-host "Application with appId=$($ApplicationId) not found" -ForegroundColor Red
		Exit
	}
}

#$siteIds = ("cezdata.sharepoint.com,5c65230e-a1d7-4bc9-9c85-93579072ba99,57cfe4b0-074f-4946-8288-a1fb91262de4","cezdata.sharepoint.com,a8ad11a3-4136-4f0d-8d7d-3ed7868cb790,3e61b528-6cfd-416d-bedb-1bcadaff94c8")

$TimeStamp = Get-Date -Format "yyyyMMdd-HHmmss"

$GraphBody = @{
	"roles" = $Role
	"grantedToIdentitiesV2" = @(
		@{
			"application" = @{
				"id" = $ApplicationId
				"displayName" = $AppName + "_" + $Role + "_" + $TimeStamp
			}
		}
	)
}

Try {
	$ResponsePOST = Invoke-WebRequest -Headers $AuthDB[$AppReg_SPO_MGMT].AuthHeaders -Uri $Uri -Body $GraphBody -Method "POST" -ContentType $ContentTypeJSON
}
Catch {
	$errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
	Write-Log "$($UPN): Error configuring MFA phone: $($IDMAuthPhone) ($($operation)) - $($errObj.error.code)" -MessageType "ERROR" -ForceOnScreen
  }