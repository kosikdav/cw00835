$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

. $ScriptPath\include-Var-Define.ps1
. $ScriptPath\include-Var-Init.ps1
. $ScriptPath\include-Functions-Common.ps1

$DownProp = "@microsoft.graph.downloadUrl"

Request-MSALToken -AppRegName "CEZ_SPO_CPR_DOWNLOAD" -TTL 30


$UriResource = "sites/cezdata.sharepoint.com:/sites/mdm"

function Get-SPOFolderREST {
	param (
		[Parameter(Mandatory=$true)][string]$Uri,
		[Parameter(Mandatory=$true)][string]$AccessToken
	)
	Do {
		$Response = Invoke-RestMethod -Uri $NextLink -Headers @{Authorization = "Bearer $AccessToken"} -Method Get -ContentType $ContentType
		$Result += $Response.value
		$NextLink = $
	} while ($NextLink)
}

$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
$Site = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB["CEZ_SPO_CPR_DOWNLOAD"].AccessToken -ContentType $ContentTypeJSON
write-host $Site.webUrl -ForegroundColor Yellow
$TargetSiteId = $Site.id
write-host $TargetSiteId -ForegroundColor Yellow

$UriResource = "sites/$($TargetSiteId)/drives"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
$Drives = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB["CEZ_SPO_CPR_DOWNLOAD"].AccessToken -ContentType $ContentTypeJSON

foreach ($Drive in $Drives) {
	write-host $Drive.name -ForegroundColor Yellow -NoNewline
	write-host " ($($Drive.driveType))"
	$UriResource = "drives/$($Drive.id)/root/children"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
	$RootItems = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB["CEZ_SPO_CPR_DOWNLOAD"].AccessToken -ContentType $ContentTypeJSON
	
	foreach ($Item in $RootItems) {
		If ($Item.folder) {
			write-host "Folder: $($Item.name)" -ForegroundColor Magenta
		}
		if ($Item.file) {
			$Outfile = "d:\data\_downtest\" + $Item.name
			write-host "File: $($Item.name) => " -ForegroundColor Cyan -NoNewline
			write-host "$($Outfile) ($($Item.size) bytes)"
			Start-BitsTransfer -Source $Item.$($DownProp) -Destination $Outfile
		}
	}
}
