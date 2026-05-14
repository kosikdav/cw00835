#######################################################################################################################
# Get-ADUsersGroupsReport-UJVREZ.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

. $ScriptPath\include-Script-Start-Generic.ps1
. $ScriptPath\include-Script-Start-Include.ps1

$LogFolder			= "t2t\azure-storage"
$LogFilePrefix		= "file-download"

$OutputFolder		= $ROF + "\" + "ad-reports"

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$StorageAccount = "cezt2tstore"
$Container = "ujvrez"

#######################################################################################################################
# variable initialization
#######################################################################################################################

$DownloadParams = @{
	StorageAccount = $StorageAccount
	Container = $container
}

#######################################################################################################################
# function definitions
#######################################################################################################################

function Get-AzureBlobList {
	[CmdletBinding()]
	param (
		[string]$StorageAccount,
		[string]$Container,
		[string]$AccessToken
	)
	# main function body ##################################
	$Uri = "https://$storageAccount.blob.core.windows.net/$container`?restype=container&comp=list"
	$headers = @{
		Authorization   = "Bearer $accessToken"
		"x-ms-version"  = "2020-04-08"
	}
	try {
		$raw = Invoke-RestMethod -Uri $Uri -Method "GET" -Headers $Headers
		[xml]$xml = $raw.substring(3)
		$AvailableBlobs = $xml.EnumerationResults.Blobs.Blob
	}
	catch {
		Write-Log "Failed to list blobs: $_" -MessageType "ERR"
	}
	return $AvailableBlobs
}

#######################################################################################################################
# main script logic
#######################################################################################################################

Write-Log "--------------------------------------------------------------"
Write-Log "Script start"
Write-Log "Script file: $($ScriptPath)\$($ScriptName)"
If ($InteractiveRun) {
    Write-Log "Running interactively"
}
Else {
    Write-Log "Running non-interactively"
}
Write-Log "Log file: $($LogFile)"

#get MSAL access token from Entra ID
Request-MSALtoken -AppRegName $AppReg_T2T_Migration_Storage -Scope "https://storage.azure.com/.default" -TTL 30

$DownloadParams.AccessToken = $AuthDB[$AppReg_T2T_Migration_Storage].AccessToken
$AvailableBlobs = Get-AzureBlobList @DownloadParams

$Headers    = @{
	Authorization  = "Bearer $($AuthDB[$AppReg_T2T_Migration_Storage].AccessToken)"
	"x-ms-version" = "2020-04-08"
}

foreach ($Blob in $AvailableBlobs) {
	$BlobName = $Blob.Name
	$BlobUri = "https://$StorageAccount.blob.core.windows.net/$Container/$BlobName"
	$LocalFilePath = Join-Path -Path $OutputFolder -ChildPath $BlobName
	if (Test-Path -Path $LocalFilePath) {
		$localFileModifiedDate = (Get-Item -Path $LocalFilePath).LastWriteTime
		$blobModifiedDate = [datetime]$Blob.Properties."Last-Modified"
		if ($localFileModifiedDate -ge $blobModifiedDate) {
			continue
		}
	}
	try {
		Invoke-RestMethod -Uri $BlobUri -Headers $Headers -Method "GET" -OutFile $LocalFilePath
		Write-Log "Successfully downloaded: $LocalFilePath"
	}
	catch {
		Write-Log "Failed to download blob: $_" -MessageType "ERR"
	}
}
Write-Log "Script finish"

