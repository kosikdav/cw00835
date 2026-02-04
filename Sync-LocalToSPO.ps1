<#
.SYNOPSIS
    Syncs a local folder structure to SharePoint Online using Microsoft Graph API.

.DESCRIPTION
    A robocopy-like utility for SharePoint Online that synchronizes local files and folders
    to a SPO document library. Supports Mirror mode (upload + delete orphans) and Copy mode
    (upload only). Uses the archive attribute to detect modified files.

.PARAMETER LocalPath
    The local folder path or UNC path to sync from.

.PARAMETER SPOSiteUrl
    The SharePoint site URL in format: "tenant.sharepoint.com:/sites/sitename"
    or the site ID in format: "tenant.sharepoint.com,guid,guid"

.PARAMETER SPOLibraryName
    The name of the document library (e.g., "Documents", "Shared Documents").

.PARAMETER SPOTargetFolder
    Optional target folder path within the document library.

.PARAMETER AppRegName
    The name of the Entra app registration for authentication.

.PARAMETER FileShareCredential
    PSCredential object for accessing network shares (UNC paths).

.PARAMETER SyncMode
    Mirror - Upload new/modified files and delete files not present locally.
    Copy - Upload new/modified files only, no deletions.

.PARAMETER UploadMode
    Replace - Overwrite existing files.
    Version - Create new versions of existing files (relies on SPO versioning settings).

.PARAMETER ResetArchiveAttribute
    If specified, clears the archive attribute on files after successful upload.

.PARAMETER LargeFileThreshold
    File size threshold in bytes for chunked uploads. Default: 262144000 (250MB).

.PARAMETER ChunkSizeMB
    Chunk size in MB for large file uploads. Default: 10.

.PARAMETER MaxRetries
    Maximum retry attempts for failed operations. Default: 3.

.PARAMETER WhatIf
    Preview mode - shows what would be done without making changes.

.EXAMPLE
    # Mirror sync from local folder to SPO
    .\Sync-LocalToSPO.ps1 -LocalPath "D:\exports" `
        -SPOSiteUrl "contoso.sharepoint.com:/sites/archive" `
        -SPOLibraryName "Documents" `
        -SPOTargetFolder "Backups/2024" `
        -AppRegName "CONTOSO_SPO_SYNC" `
        -SyncMode Mirror `
        -UploadMode Replace `
        -ResetArchiveAttribute

.EXAMPLE
    # Copy mode from file share with credentials
    .\Sync-LocalToSPO.ps1 -LocalPath "\\fileserver\share\data" `
        -SPOSiteUrl "contoso.sharepoint.com:/sites/reports" `
        -SPOLibraryName "Shared Documents" `
        -AppRegName "CONTOSO_SPO_SYNC" `
        -FileShareCredential $cred `
        -SyncMode Copy `
        -UploadMode Version

.EXAMPLE
    # Preview what would be synced (WhatIf)
    .\Sync-LocalToSPO.ps1 -LocalPath "D:\data" `
        -SPOSiteUrl "contoso.sharepoint.com:/sites/mysite" `
        -SPOLibraryName "Documents" `
        -AppRegName "CONTOSO_SPO_SYNC" `
        -SyncMode Mirror `
        -UploadMode Replace `
        -WhatIf
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$LocalPath,

    [Parameter(Mandatory=$true)]
    [string]$SPOSiteUrl,

    [Parameter(Mandatory=$true)]
    [string]$SPOLibraryName,

    [string]$SPOTargetFolder = "",

    [Parameter(Mandatory=$true)]
    [string]$AppRegName,

    [PSCredential]$FileShareCredential,

    [Parameter(Mandatory=$true)]
    [ValidateSet("Mirror","Copy")]
    [string]$SyncMode,

    [Parameter(Mandatory=$true)]
    [ValidateSet("Replace","Version")]
    [string]$UploadMode,

    [switch]$ResetArchiveAttribute,

    [int64]$LargeFileThreshold = 262144000,

    [int]$ChunkSizeMB = 10,

    [int]$MaxRetries = 3,

    [switch]$WhatIf,

    [Alias("Definitions","IniFile")]
    [string]$VariableDefinitionFile
)

#######################################################################################################################
# INITIALIZATION
#######################################################################################################################

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

# Include standard startup block
. $ScriptPath\include-Script-StdStartBlock.ps1

# Script-specific configuration
$LogFolder = "spo-sync"
$LogFilePrefix = "sync-local-to-spo"

# Include standard include block (loads variables and functions)
. $ScriptPath\include-Script-StdIncBlock.ps1

# Include SPO Graph functions
. $ScriptPath\include-functions-SPO-Graph.ps1

# Create log file
$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

# Include standard log start block
. $IncFile_StdLogStartBlock

#######################################################################################################################
# MAIN SCRIPT
#######################################################################################################################

$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$NetworkShareConnected = $false

Write-Log "========================================================================"
Write-Log "SharePoint Online Sync"
Write-Log "========================================================================"
Write-Log "Local Path:      $LocalPath"
Write-Log "SPO Site:        $SPOSiteUrl"
Write-Log "SPO Library:     $SPOLibraryName"
Write-Log "SPO Target:      $(if ([string]::IsNullOrEmpty($SPOTargetFolder)) { '(root)' } else { $SPOTargetFolder })"
Write-Log "Sync Mode:       $SyncMode"
Write-Log "Upload Mode:     $UploadMode"
Write-Log "Archive Attr:    $(if ($ResetArchiveAttribute) { 'Reset after upload' } else { 'No change' })"
Write-Log "Large File:      > $([math]::Round($LargeFileThreshold/1MB, 0)) MB uses chunked upload"
Write-Log "WhatIf Mode:     $WhatIf"
Write-Log "========================================================================"

try {
    #######################################################################################################################
    # CONNECT TO NETWORK SHARE (if UNC path and credentials provided)
    #######################################################################################################################

    if ($LocalPath.StartsWith("\\") -and $FileShareCredential) {
        Write-Log "Connecting to network share: $LocalPath"
        $NetworkShareConnected = Connect-NetworkShareWithCredential -SharePath $LocalPath -Credential $FileShareCredential
        if (-not $NetworkShareConnected) {
            Write-Log "Failed to connect to network share. Aborting." -MessageType Error
            exit 1
        }
    }

    # Validate local path exists
    if (-not (Test-Path -Path $LocalPath)) {
        Write-Log "Local path does not exist: $LocalPath" -MessageType Error
        exit 1
    }

    #######################################################################################################################
    # AUTHENTICATE TO GRAPH API
    #######################################################################################################################

    Write-Log "Authenticating to Microsoft Graph..."
    Request-MSALToken -AppRegName $AppRegName -TTL 30

    if (-not $AuthDB[$AppRegName]) {
        Write-Log "Failed to acquire access token. Aborting." -MessageType Error
        exit 1
    }

    $AccessToken = $AuthDB[$AppRegName].AccessToken
    Write-Log "Authentication successful."

    #######################################################################################################################
    # RESOLVE SPO SITE AND DRIVE
    #######################################################################################################################

    Write-Log "Resolving SharePoint site..."
    $SPOSiteId = Get-SPOSiteId -AccessToken $AccessToken -SiteUrl $SPOSiteUrl

    if (-not $SPOSiteId) {
        Write-Log "Failed to resolve SharePoint site. Aborting." -MessageType Error
        exit 1
    }

    Write-Log "Resolving document library '$SPOLibraryName'..."
    $SPODriveId = Get-SPODriveId -AccessToken $AccessToken -SPOSiteId $SPOSiteId -LibraryName $SPOLibraryName

    if (-not $SPODriveId) {
        Write-Log "Failed to resolve document library. Aborting." -MessageType Error
        exit 1
    }

    #######################################################################################################################
    # CREATE TARGET FOLDER STRUCTURE (if needed)
    #######################################################################################################################

    if (-not [string]::IsNullOrEmpty($SPOTargetFolder)) {
        Write-Log "Ensuring target folder structure exists..."
        $SPOTargetFolder = $SPOTargetFolder.Replace("\", "/").TrimStart("/").TrimEnd("/")
        $FolderParts = $SPOTargetFolder -split "/"
        $CurrentPath = ""

        foreach ($Part in $FolderParts) {
            $ParentPath = $CurrentPath
            $CurrentPath = if ([string]::IsNullOrEmpty($CurrentPath)) { $Part } else { "$CurrentPath/$Part" }

            $FolderExists = Test-SPOItemExists -AccessToken $AccessToken -SPODriveId $SPODriveId -ItemPath $CurrentPath

            if (-not $FolderExists) {
                if ($WhatIf) {
                    Write-Log "[WhatIf] Would create folder: $CurrentPath"
                }
                else {
                    Write-Log "Creating folder: $CurrentPath"
                    Create-SharePointFolder -accessToken $AccessToken -SPODriveId $SPODriveId `
                        -parentItemId $ParentPath -folderName $Part -conflictBehavior "Fail"
                }
            }
        }
    }

    #######################################################################################################################
    # PERFORM SYNC
    #######################################################################################################################

    Write-Log "Starting synchronization..."
    Write-Log "------------------------------------------------------------------------"

    $SyncParams = @{
        AccessToken = $AccessToken
        LocalPath = $LocalPath
        SPOSiteId = $SPOSiteId
        SPODriveId = $SPODriveId
        SPOLibraryName = $SPOLibraryName
        SPOCurrentPath = $SPOTargetFolder
        SyncMode = $SyncMode
        UploadMode = $UploadMode
        ResetArchiveAttribute = $ResetArchiveAttribute
        LargeFileThreshold = $LargeFileThreshold
        ChunkSizeMB = $ChunkSizeMB
        MaxRetries = $MaxRetries
        WhatIf = $WhatIf
    }

    $Stats = Sync-FolderToSharePointEx @SyncParams

    #######################################################################################################################
    # SUMMARY
    #######################################################################################################################

    Write-Log "------------------------------------------------------------------------"
    Write-Log "Synchronization complete."
    Write-Log "========================================================================"
    Write-Log "SUMMARY"
    Write-Log "========================================================================"
    Write-Log "Files uploaded:    $($Stats.FilesUploaded)"
    Write-Log "Files skipped:     $($Stats.FilesSkipped)"
    Write-Log "Files deleted:     $($Stats.FilesDeleted)"
    Write-Log "Folders created:   $($Stats.FoldersCreated)"
    Write-Log "Folders deleted:   $($Stats.FoldersDeleted)"
    Write-Log "Errors:            $($Stats.Errors)"
    Write-Log "Bytes uploaded:    $([math]::Round($Stats.BytesUploaded/1MB, 2)) MB"
    Write-Log "========================================================================"

    if ($Stats.Errors -gt 0) {
        Write-Log "Sync completed with $($Stats.Errors) error(s)." -MessageType Warning
    }
}
catch {
    Write-Log "Unexpected error: $($_.Exception.Message)" -MessageType Error
    Write-Log $_.ScriptStackTrace -MessageType Error
}
finally {
    # Disconnect network share if connected
    if ($NetworkShareConnected -and $LocalPath.StartsWith("\\")) {
        Write-Log "Disconnecting from network share..."
        Disconnect-NetworkShare -SharePath $LocalPath | Out-Null
    }
}

#######################################################################################################################
# END
#######################################################################################################################

$Stopwatch.Stop()
Write-Log "Execution time: $($Stopwatch.Elapsed.ToString('hh\:mm\:ss\.fff'))"

. $IncFile_StdLogEndBlock
