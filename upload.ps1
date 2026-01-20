$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

. $ScriptPath\include-Var-Define.ps1
. $ScriptPath\include-Var-Init.ps1
. $ScriptPath\include-Functions-Common.ps1
. $ScriptPath\include-Functions-SPO-Graph.ps1

##############################################################################################################
# Main
$ProgressPreference = 'SilentlyContinue'

$LocalFolder = "d:\data\_downtest"  # Local folder path to upload
$UNC = "\\cezdata.corp\sdp\Public\PrehledVelicin\EVD\EST"

#$TargetSiteId = "cezdata.sharepoint.com,953dff82-356d-4d0e-89ab-3388f601b5ec,6aeca208-fdc9-4774-8ca0-699260404c20"
$TargetSiteId = "cezdata.sharepoint.com,f39805be-1db5-4458-ae7d-010c0cb2773c,fa4e1e61-1297-47ce-acf9-2b85240f3342"

$SPOLibraryName = "Dokumenty"
$SPOFolder = "_upload"

Request-MSALToken -AppRegName "CEZ_SPO_CPR_DOWNLOAD" -TTL 30

$rootFolderName = Split-Path $LocalFolder -Leaf
write-host "Local folder: $LocalFolder"
write-host "Root local folder: $rootFolderName"
write-host "SharePoint library name: $SPOLibraryName"
write-host "SharePoint folder: $SPOFolder"
write-host "Site id: $TargetSiteId"
$UriResource = "sites/$($TargetSiteId)/drives"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
$Drives = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB["CEZ_SPO_CPR_DOWNLOAD"].AccessToken -ContentType $ContentTypeJSON
foreach ($Drive in $Drives) {
    if (($Drive.driveType -eq "documentLibrary") -and ($SPOLibraryName -eq $Drive.name)) {
        write-host "Found SharePoint library: $($Drive.name)"
        $driveId = $Drive.id
    }
}
exit

write-host "###############################################################################################################"

$FunctionParams = @{
    accessToken = $AuthDB["CEZ_SPO_CPR_DOWNLOAD"].AccessToken
    currentLocalPath = $LocalFolder
    SPOSiteId = $TargetSiteId
    SPODriveId = $driveId
    SPOLibraryName = $SPOLibraryName 
    SPOCurrentPath = $SPOFolder
    conflictBehavior = "Replace"
}

Sync-FolderToSharePoint @FunctionParams
