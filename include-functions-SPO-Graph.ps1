#######################################################################################################################
#######################################################################################################################
# INCLUDE-FUNCTIONS-SPO-GRAPH
#######################################################################################################################
#######################################################################################################################
#
#
#
########################################################################################
# Create-SharePointFolder
########################################################################################

function Create-SharePointFolder {
    param(
        [string]$accessToken,
        [string]$SPOSiteId,
        [string]$SPODriveId,
        [string]$parentItemId,
        [string]$folderName,
        [string][ValidateSet("Rename","Fail","Replace")]$conflictBehavior = "Rename"
    )
    
    $FolderExists = $false 
    
    if (($null -eq $parentItemId) -or ($parentItemId -eq [string]::Empty) -or ($parentItemId -eq "root")) {
        $UriResource = "drives/$($SPODriveId)/items/root/children"
        $FolderExists = $true
    }
    else {
        $parentItemId = $parentItemId.Replace("\","/")
        $UriResource = "drives/$($SPODriveId)/items/root:/$($parentItemId):/children"
        $UriResourceTestEx = "drives/$($SPODriveId)/items/root:/$($parentItemId)/$($folderName)"
        $Uri = New-GraphUri -Version "v1.0" -Resource $UriResourceTestEx
        Try {
            $Result = Invoke-RestMethod -Method "GET" -Uri $Uri -Headers $Headers -ContentType $ContentTypeJSON | Out-Null
            $FolderExists = $true
        }
        Catch {
            $FolderExists = $false
        }
    }
    
    if (-not $FolderExists) {
        $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
        write-host "Creating folder $($parentItemId)/$($folderName)"
        $body = @{
            "name" = $folderName
            "folder" = @{}
            "@microsoft.graph.conflictBehavior" = $conflictBehavior
        } | ConvertTo-Json
    
        $headers = @{
            Authorization = "Bearer $accessToken"
        }
    
        Try {
            Invoke-RestMethod -Method "POST" -Uri $Uri -Headers $Headers -Body $Body -ContentType $ContentTypeJSON | Out-Null
        }   
        Catch {
            Write-Host "Error creating folder: $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
}

########################################################################################
# Upload-FileToSharePoint
########################################################################################
function Upload-FileToSharePoint {
    param(
        [string]$AccessToken,
        [string]$ADCredential,
        [string]$LocalFilePath,
        [string]$SPOSiteId,
        [string]$SPODriveId,
        [string]$SPOLibraryName,
        [string]$SPOFolder,
        [string][ValidateSet("Rename","Fail","Replace")]$conflictBehavior = "Rename",
        [switch]$ResetArchiveAttribute
    )
    $SPOFileName = Split-Path $localFilePath -Leaf
    $UriResource = "drives/$($SPODriveId)/items/root:/$($SPOFolder)/$($SPOFileName):/content"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
    $Headers = @{
        Authorization = "Bearer $accessToken"
    }
    $resultUpload = Invoke-RestMethod -Method "PUT" -Uri $Uri -Headers $headers -Credential $ADCredential -InFile $localFilePath -ContentType "application/octet-stream" | Out-Null
    if ($ResetArchiveAttribute) {
        $file = Get-Item -Path $LocalFilePath
        $file.Attributes = $file.Attributes -band (-bnot [System.IO.FileAttributes]::Archive)
        if ($file.Attributes -band [System.IO.FileAttributes]::Archive) {
            Write-Host "Failed to clear the Archive attribute."
        } 
    }
}

########################################################################################
# Sync-FolderToSharePoint
########################################################################################
function Sync-FolderToSharePoint {
    param (
        [string]$AccessToken,
        [string]$ADCredential,
        [string]$currentLocalPath,
        [string]$SPOSiteId,
        [string]$SPODriveId,
        [string]$SPOLibraryName,
        [string]$SPOCurrentPath,
        [switch]$RecursiveCall,
        [switch]$ResetArchiveAttribute,
        [string][ValidateSet("Rename","Fail","Replace")]$conflictBehavior = "Rename"
    )
    $LocalItems = $RemoteItems = $null
    $Headers = @{
        Authorization = "Bearer $accessToken"
    }
    $FolderName = Split-Path $SPOCurrentPath -Leaf
    Try {
        $ParentFolder = Split-Path $SPOCurrentPath -Parent
    }
    Catch {
        $ParentFolder = "root"
    }

    Create-SharePointFolder -accessToken $accessToken -SPODriveId $SPODriveId -parentItemId $ParentFolder -folderName $FolderName -conflictBehavior $conflictBehavior

    $LocalItems = Get-ChildItem -Path $currentLocalPath
    
    $UriResource = "drives/$($SPODriveId)/items/root:/$($SPOCurrentPath):/children"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
    $RemoteItems = Get-GraphOutputREST -Uri $Uri -AccessToken $AccessToken -ContentType $ContentTypeJSON
    if ($RemoteItems.Count -gt 0) {
        foreach ($RemoteItem in $RemoteItems) {
            if ($RemoteItem.folder) {
                $itemType = "folder"
            } else {
                $itemType = "file"
            }
            if (-not ($LocalItems.Name -contains $RemoteItem.Name)) {
                write-host "deleting  $($SPOCurrentPath)/$($RemoteItem.Name)" -ForegroundColor Red
                $UriResource = "sites/$($SPOSiteId)/drive/items/$($RemoteItem.id)"
                $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
                $resultDelete = Invoke-WebRequest -Method "DELETE" -Uri $Uri -Headers $Headers
            }
        }
    }
    if ($LocalItems.Count -gt 0) {
        foreach ($LocalItem in $LocalItems) {
            if ($LocalItem.PSIsContainer) {
                $FunctionParams = @{
                    accessToken = $AccessToken
                    ADCredential = $ADCredential
                    currentLocalPath = $LocalItem.FullName
                    SPOSiteId = $SPOSiteId
                    SPODriveId = $SPODriveId
                    SPOLibraryName = $SPOLibraryName 
                    SPOCurrentPath = "$SPOCurrentPath/$($LocalItem.Name)"
                    conflictBehavior = "Replace"
                    RecursiveCall = $true
                    ResetArchiveAttribute = $ResetArchiveAttribute
                }
                Sync-FolderToSharePoint @FunctionParams
            } 
            else {
                if ((-not $RemoteItems.Name -contains $LocalItem.Name) -or ($LocalItem.Mode.Substring(1,1) -eq "a")) {
                    $FunctionParams = @{
                        accessToken = $AccessToken
                        ADCredential = $ADCredential
                        LocalFilePath = $LocalItem.FullName
                        SPOSiteId = $SPOSiteId
                        SPODriveId = $SPODriveId
                        SPOLibraryName = $SPOLibraryName 
                        SPOFolder = $SPOCurrentPath
                        ConflictBehavior = "Replace"
                        ResetArchiveAttribute = $ResetArchiveAttribute
                    }
                    Write-Host "uploading $($localItem.FullName)"
                    Upload-FileToSharePoint @FunctionParams
                }
                else {
                    write-host "skipping  $($localItem.FullName)" -ForegroundColor DarkGray
                }
            }
        }
    }
}

########################################################################################
# UTILITY FUNCTIONS
########################################################################################

########################################################################################
# Test-FileArchiveAttribute
# Returns $true if the archive attribute is set on the file
########################################################################################
function Test-FileArchiveAttribute {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath
    )
    try {
        $file = Get-Item -Path $FilePath -ErrorAction Stop
        return (($file.Attributes -band [System.IO.FileAttributes]::Archive) -ne 0)
    }
    catch {
        Write-Log "Error checking archive attribute for $($FilePath): $($_.Exception.Message)" -MessageType Error
        return $false
    }
}

########################################################################################
# Clear-FileArchiveAttribute
# Clears the archive attribute on the file
########################################################################################
function Clear-FileArchiveAttribute {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath
    )
    try {
        $file = Get-Item -Path $FilePath -ErrorAction Stop
        $file.Attributes = $file.Attributes -band (-bnot [System.IO.FileAttributes]::Archive)
        return $true
    }
    catch {
        Write-Log "Error clearing archive attribute for $($FilePath): $($_.Exception.Message)" -MessageType Error
        return $false
    }
}

########################################################################################
# Connect-NetworkShareWithCredential
# Maps a network share with the provided credentials, returns the mapped drive letter
########################################################################################
function Connect-NetworkShareWithCredential {
    param(
        [Parameter(Mandatory=$true)][string]$SharePath,
        [Parameter(Mandatory=$true)][PSCredential]$Credential
    )
    try {
        $netUse = net use $SharePath /user:$($Credential.UserName) $($Credential.GetNetworkCredential().Password) 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Log "Connected to network share: $SharePath"
            return $true
        }
        else {
            Write-Log "Failed to connect to network share: $SharePath - $netUse" -MessageType Error
            return $false
        }
    }
    catch {
        Write-Log "Error connecting to network share $($SharePath): $($_.Exception.Message)" -MessageType Error
        return $false
    }
}

########################################################################################
# Disconnect-NetworkShare
# Disconnects a network share
########################################################################################
function Disconnect-NetworkShare {
    param(
        [Parameter(Mandatory=$true)][string]$SharePath
    )
    try {
        $netUse = net use $SharePath /delete /y 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Log "Disconnected from network share: $SharePath"
            return $true
        }
        return $false
    }
    catch {
        Write-Log "Error disconnecting network share $($SharePath): $($_.Exception.Message)" -MessageType Error
        return $false
    }
}

########################################################################################
# SITE/DRIVE DISCOVERY FUNCTIONS
########################################################################################

########################################################################################
# Get-SPOSiteId
# Resolves a SharePoint site URL to its site ID
# SiteUrl format: "tenant.sharepoint.com:/sites/sitename" or "tenant.sharepoint.com,guid,guid"
########################################################################################
function Get-SPOSiteId {
    param(
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][string]$SiteUrl
    )
    $Headers = @{ Authorization = "Bearer $AccessToken" }

    # If already in ID format (contains commas), return as-is
    if ($SiteUrl -match ',') {
        return $SiteUrl
    }

    # Parse the site URL
    $UriResource = "sites/$SiteUrl"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource

    try {
        $Site = Invoke-RestMethod -Method "GET" -Uri $Uri -Headers $Headers -ContentType $ContentTypeJSON
        Write-Log "Resolved site ID: $($Site.id)"
        return $Site.id
    }
    catch {
        Write-Log "Error resolving site URL $($SiteUrl): $($_.Exception.Message)" -MessageType Error
        return $null
    }
}

########################################################################################
# Get-SPODriveId
# Gets the drive ID for a document library by name
########################################################################################
function Get-SPODriveId {
    param(
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][string]$SPOSiteId,
        [Parameter(Mandatory=$true)][string]$LibraryName
    )
    $Headers = @{ Authorization = "Bearer $AccessToken" }
    $UriResource = "sites/$($SPOSiteId)/drives"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource

    try {
        $Drives = Get-GraphOutputREST -Uri $Uri -AccessToken $AccessToken -ContentType $ContentTypeJSON
        foreach ($Drive in $Drives) {
            if (($Drive.driveType -eq "documentLibrary") -and ($Drive.name -eq $LibraryName)) {
                Write-Log "Resolved drive ID for '$LibraryName': $($Drive.id)"
                return $Drive.id
            }
        }
        Write-Log "Document library '$LibraryName' not found in site" -MessageType Error
        return $null
    }
    catch {
        Write-Log "Error getting drives for site $($SPOSiteId): $($_.Exception.Message)" -MessageType Error
        return $null
    }
}

########################################################################################
# Get-SPODriveItemByPath
# Gets a drive item (file or folder) by its path
########################################################################################
function Get-SPODriveItemByPath {
    param(
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][string]$SPODriveId,
        [Parameter(Mandatory=$true)][string]$ItemPath
    )
    $Headers = @{ Authorization = "Bearer $AccessToken" }
    $ItemPath = $ItemPath.Replace("\", "/").TrimStart("/")

    if ([string]::IsNullOrEmpty($ItemPath) -or $ItemPath -eq "root") {
        $UriResource = "drives/$($SPODriveId)/root"
    }
    else {
        $UriResource = "drives/$($SPODriveId)/root:/$($ItemPath)"
    }
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource

    try {
        $Item = Invoke-RestMethod -Method "GET" -Uri $Uri -Headers $Headers -ContentType $ContentTypeJSON
        return $Item
    }
    catch {
        if ($_.Exception.Response.StatusCode -eq 404) {
            return $null
        }
        Write-Log "Error getting item at path '$ItemPath': $($_.Exception.Message)" -MessageType Error
        return $null
    }
}

########################################################################################
# Get-SPOFolderChildren
# Lists all items in a SharePoint folder
########################################################################################
function Get-SPOFolderChildren {
    param(
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][string]$SPODriveId,
        [string]$FolderPath = "",
        [string]$ItemId = ""
    )
    $FolderPath = $FolderPath.Replace("\", "/").TrimStart("/")

    if (-not [string]::IsNullOrEmpty($ItemId)) {
        $UriResource = "drives/$($SPODriveId)/items/$($ItemId)/children"
    }
    elseif ([string]::IsNullOrEmpty($FolderPath) -or $FolderPath -eq "root") {
        $UriResource = "drives/$($SPODriveId)/root/children"
    }
    else {
        $UriResource = "drives/$($SPODriveId)/root:/$($FolderPath):/children"
    }
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource

    try {
        $Children = Get-GraphOutputREST -Uri $Uri -AccessToken $AccessToken -ContentType $ContentTypeJSON
        return $Children
    }
    catch {
        Write-Log "Error listing folder contents at '$FolderPath': $($_.Exception.Message)" -MessageType Error
        return @()
    }
}

########################################################################################
# Test-SPOItemExists
# Checks if an item exists at the specified path
########################################################################################
function Test-SPOItemExists {
    param(
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][string]$SPODriveId,
        [Parameter(Mandatory=$true)][string]$ItemPath
    )
    $Item = Get-SPODriveItemByPath -AccessToken $AccessToken -SPODriveId $SPODriveId -ItemPath $ItemPath
    return ($null -ne $Item)
}

########################################################################################
# FILE OPERATION FUNCTIONS
########################################################################################

########################################################################################
# Remove-SPOItem
# Deletes a file or folder from SharePoint
########################################################################################
function Remove-SPOItem {
    param(
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][string]$SPOSiteId,
        [Parameter(Mandatory=$true)][string]$ItemId
    )
    $Headers = @{ Authorization = "Bearer $AccessToken" }
    $UriResource = "sites/$($SPOSiteId)/drive/items/$($ItemId)"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource

    try {
        Invoke-RestMethod -Method "DELETE" -Uri $Uri -Headers $Headers | Out-Null
        return $true
    }
    catch {
        Write-Log "Error deleting item $($ItemId): $($_.Exception.Message)" -MessageType Error
        return $false
    }
}

########################################################################################
# LARGE FILE UPLOAD FUNCTIONS
########################################################################################

########################################################################################
# New-SPOUploadSession
# Creates a resumable upload session for large files
########################################################################################
function New-SPOUploadSession {
    param(
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][string]$SPODriveId,
        [Parameter(Mandatory=$true)][string]$TargetPath,
        [Parameter(Mandatory=$true)][string]$FileName,
        [ValidateSet("rename","replace","fail")][string]$ConflictBehavior = "replace"
    )
    $Headers = @{
        Authorization = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }
    $TargetPath = $TargetPath.Replace("\", "/").TrimStart("/")

    if ([string]::IsNullOrEmpty($TargetPath)) {
        $UriResource = "drives/$($SPODriveId)/root:/$($FileName):/createUploadSession"
    }
    else {
        $UriResource = "drives/$($SPODriveId)/root:/$($TargetPath)/$($FileName):/createUploadSession"
    }
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource

    $Body = @{
        item = @{
            "@microsoft.graph.conflictBehavior" = $ConflictBehavior
            name = $FileName
        }
    } | ConvertTo-Json

    try {
        $Session = Invoke-RestMethod -Method "POST" -Uri $Uri -Headers $Headers -Body $Body
        Write-Log "Created upload session for $FileName (expires: $($Session.expirationDateTime))"
        return $Session
    }
    catch {
        Write-Log "Error creating upload session for $($FileName): $($_.Exception.Message)" -MessageType Error
        return $null
    }
}

########################################################################################
# Send-SPOFileChunk
# Uploads a single chunk of a large file
########################################################################################
function Send-SPOFileChunk {
    param(
        [Parameter(Mandatory=$true)][string]$UploadUrl,
        [Parameter(Mandatory=$true)][byte[]]$ChunkData,
        [Parameter(Mandatory=$true)][int64]$RangeStart,
        [Parameter(Mandatory=$true)][int64]$RangeEnd,
        [Parameter(Mandatory=$true)][int64]$TotalSize
    )
    $Headers = @{
        "Content-Length" = $ChunkData.Length
        "Content-Range" = "bytes $RangeStart-$RangeEnd/$TotalSize"
    }

    try {
        $Response = Invoke-RestMethod -Method "PUT" -Uri $UploadUrl -Headers $Headers -Body $ChunkData
        return $Response
    }
    catch {
        throw $_
    }
}

########################################################################################
# Upload-LargeFileToSharePoint
# Orchestrates chunked upload for large files (>4MB)
########################################################################################
function Upload-LargeFileToSharePoint {
    param(
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][string]$LocalFilePath,
        [Parameter(Mandatory=$true)][string]$SPODriveId,
        [string]$SPOFolder = "",
        [int]$ChunkSizeMB = 10,
        [ValidateSet("replace","fail")][string]$ConflictBehavior = "replace",
        [switch]$ResetArchiveAttribute,
        [int]$MaxRetries = 3
    )

    $FileName = Split-Path $LocalFilePath -Leaf
    $FileInfo = Get-Item -Path $LocalFilePath
    $FileSize = $FileInfo.Length
    $ChunkSize = $ChunkSizeMB * 1024 * 1024

    # Ensure chunk size is multiple of 320KB (Graph API requirement)
    $ChunkSize = [math]::Floor($ChunkSize / 327680) * 327680
    if ($ChunkSize -lt 327680) { $ChunkSize = 327680 }

    Write-Log "Starting chunked upload for $FileName ($([math]::Round($FileSize/1MB, 2)) MB, chunk size: $([math]::Round($ChunkSize/1MB, 2)) MB)"

    # Create upload session
    $Session = New-SPOUploadSession -AccessToken $AccessToken -SPODriveId $SPODriveId `
        -TargetPath $SPOFolder -FileName $FileName -ConflictBehavior $ConflictBehavior

    if (-not $Session) {
        Write-Log "Failed to create upload session for $FileName" -MessageType Error
        return $null
    }

    $UploadUrl = $Session.uploadUrl
    $FileStream = $null
    $Result = $null

    try {
        $FileStream = [System.IO.File]::OpenRead($LocalFilePath)
        $Position = 0
        $ChunkNumber = 0
        $TotalChunks = [math]::Ceiling($FileSize / $ChunkSize)

        while ($Position -lt $FileSize) {
            $ChunkNumber++
            $BytesToRead = [math]::Min($ChunkSize, $FileSize - $Position)
            $ChunkData = New-Object byte[] $BytesToRead
            $BytesRead = $FileStream.Read($ChunkData, 0, $BytesToRead)

            $RangeStart = $Position
            $RangeEnd = $Position + $BytesRead - 1

            $RetryCount = 0
            $ChunkUploaded = $false

            while (-not $ChunkUploaded -and $RetryCount -lt $MaxRetries) {
                try {
                    Write-Log "Uploading chunk $ChunkNumber/$TotalChunks (bytes $RangeStart-$RangeEnd)" -ForegroundColor DarkGray
                    $Response = Send-SPOFileChunk -UploadUrl $UploadUrl -ChunkData $ChunkData `
                        -RangeStart $RangeStart -RangeEnd $RangeEnd -TotalSize $FileSize
                    $ChunkUploaded = $true

                    # Check if this was the final chunk (response contains file metadata)
                    if ($Response.id) {
                        $Result = $Response
                    }
                }
                catch {
                    $RetryCount++
                    $ErrorMsg = $_.Exception.Message
                    if ($RetryCount -lt $MaxRetries) {
                        $Delay = [math]::Pow(2, $RetryCount) * 5
                        Write-Log "Chunk upload failed, retrying in $Delay seconds... (attempt $RetryCount/$MaxRetries)" -MessageType Warning
                        Start-Sleep -Seconds $Delay
                    }
                    else {
                        Write-Log "Chunk upload failed after $MaxRetries attempts: $ErrorMsg" -MessageType Error
                        throw $_
                    }
                }
            }

            $Position += $BytesRead
        }

        if ($Result) {
            Write-Log "Successfully uploaded $FileName ($([math]::Round($FileSize/1MB, 2)) MB)"

            if ($ResetArchiveAttribute) {
                Clear-FileArchiveAttribute -FilePath $LocalFilePath | Out-Null
            }
        }

        return $Result
    }
    catch {
        Write-Log "Error during chunked upload of $($FileName): $($_.Exception.Message)" -MessageType Error
        return $null
    }
    finally {
        if ($FileStream) {
            $FileStream.Close()
            $FileStream.Dispose()
        }
    }
}

########################################################################################
# SYNC LOGIC FUNCTIONS
########################################################################################

########################################################################################
# Compare-LocalAndRemoteFiles
# Compares local folder contents with SPO folder contents
########################################################################################
function Compare-LocalAndRemoteFiles {
    param(
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][string]$LocalPath,
        [Parameter(Mandatory=$true)][string]$SPODriveId,
        [string]$SPOFolderPath = "",
        [switch]$UseArchiveAttribute
    )

    $Result = [PSCustomObject]@{
        FilesToUpload = @()
        FilesToSkip = @()
        FoldersToCreate = @()
        ItemsToDelete = @()
    }

    # Get local items
    $LocalItems = Get-ChildItem -Path $LocalPath -ErrorAction SilentlyContinue

    # Get remote items
    $RemoteItems = Get-SPOFolderChildren -AccessToken $AccessToken -SPODriveId $SPODriveId -FolderPath $SPOFolderPath
    $RemoteItemNames = @{}
    foreach ($item in $RemoteItems) {
        $RemoteItemNames[$item.name] = $item
    }

    # Check local items
    foreach ($LocalItem in $LocalItems) {
        if ($LocalItem.PSIsContainer) {
            # Folder
            if (-not $RemoteItemNames.ContainsKey($LocalItem.Name)) {
                $Result.FoldersToCreate += $LocalItem
            }
        }
        else {
            # File
            $ShouldUpload = $false

            if (-not $RemoteItemNames.ContainsKey($LocalItem.Name)) {
                # File doesn't exist in SPO
                $ShouldUpload = $true
            }
            elseif ($UseArchiveAttribute) {
                # Check archive attribute
                $ShouldUpload = Test-FileArchiveAttribute -FilePath $LocalItem.FullName
            }
            else {
                # Always upload (for Replace/Version modes)
                $ShouldUpload = $true
            }

            if ($ShouldUpload) {
                $Result.FilesToUpload += $LocalItem
            }
            else {
                $Result.FilesToSkip += $LocalItem
            }
        }
    }

    # Check for items to delete (items in SPO but not locally)
    $LocalItemNames = $LocalItems | ForEach-Object { $_.Name }
    foreach ($RemoteItem in $RemoteItems) {
        if ($LocalItemNames -notcontains $RemoteItem.name) {
            $Result.ItemsToDelete += $RemoteItem
        }
    }

    return $Result
}

########################################################################################
# Sync-FolderToSharePointEx
# Extended sync function with Mirror/Copy modes and large file support
########################################################################################
function Sync-FolderToSharePointEx {
    param(
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][string]$LocalPath,
        [Parameter(Mandatory=$true)][string]$SPOSiteId,
        [Parameter(Mandatory=$true)][string]$SPODriveId,
        [Parameter(Mandatory=$true)][string]$SPOLibraryName,
        [string]$SPOCurrentPath = "",
        [ValidateSet("Mirror","Copy")][string]$SyncMode = "Copy",
        [ValidateSet("Replace","Version")][string]$UploadMode = "Replace",
        [switch]$ResetArchiveAttribute,
        [int64]$LargeFileThreshold = 262144000,
        [int]$ChunkSizeMB = 10,
        [int]$MaxRetries = 3,
        [switch]$WhatIf,
        [switch]$RecursiveCall
    )

    # Initialize statistics (only at top level)
    if (-not $RecursiveCall) {
        $script:SyncStats = [PSCustomObject]@{
            FilesUploaded = 0
            FilesSkipped = 0
            FilesDeleted = 0
            FoldersCreated = 0
            FoldersDeleted = 0
            Errors = 0
            BytesUploaded = 0
        }
    }

    $Headers = @{ Authorization = "Bearer $AccessToken" }
    $SPOCurrentPath = $SPOCurrentPath.Replace("\", "/").TrimStart("/")

    # Ensure target folder exists
    if (-not [string]::IsNullOrEmpty($SPOCurrentPath)) {
        $FolderName = Split-Path $SPOCurrentPath -Leaf
        $ParentFolder = Split-Path $SPOCurrentPath -Parent
        if ([string]::IsNullOrEmpty($ParentFolder)) { $ParentFolder = "" }

        $FolderExists = Test-SPOItemExists -AccessToken $AccessToken -SPODriveId $SPODriveId -ItemPath $SPOCurrentPath
        if (-not $FolderExists) {
            if ($WhatIf) {
                Write-Log "[WhatIf] Would create folder: $SPOCurrentPath"
            }
            else {
                Create-SharePointFolder -accessToken $AccessToken -SPODriveId $SPODriveId `
                    -parentItemId $ParentFolder -folderName $FolderName -conflictBehavior "Fail"
                $script:SyncStats.FoldersCreated++
            }
        }
    }

    # Compare local and remote
    $Comparison = Compare-LocalAndRemoteFiles -AccessToken $AccessToken -LocalPath $LocalPath `
        -SPODriveId $SPODriveId -SPOFolderPath $SPOCurrentPath -UseArchiveAttribute:$ResetArchiveAttribute

    # Process folders to create (and recurse)
    $LocalFolders = Get-ChildItem -Path $LocalPath -Directory -ErrorAction SilentlyContinue
    foreach ($Folder in $LocalFolders) {
        $NewSPOPath = if ([string]::IsNullOrEmpty($SPOCurrentPath)) { $Folder.Name } else { "$SPOCurrentPath/$($Folder.Name)" }

        Sync-FolderToSharePointEx -AccessToken $AccessToken -LocalPath $Folder.FullName `
            -SPOSiteId $SPOSiteId -SPODriveId $SPODriveId -SPOLibraryName $SPOLibraryName `
            -SPOCurrentPath $NewSPOPath -SyncMode $SyncMode -UploadMode $UploadMode `
            -ResetArchiveAttribute:$ResetArchiveAttribute -LargeFileThreshold $LargeFileThreshold `
            -ChunkSizeMB $ChunkSizeMB -MaxRetries $MaxRetries -WhatIf:$WhatIf -RecursiveCall
    }

    # Upload files
    $ConflictBehavior = if ($UploadMode -eq "Replace") { "replace" } else { "replace" }  # Version mode also uses replace, SPO handles versioning

    foreach ($File in $Comparison.FilesToUpload) {
        if ($WhatIf) {
            Write-Log "[WhatIf] Would upload: $($File.FullName) -> $SPOCurrentPath/$($File.Name)"
            continue
        }

        try {
            if ($File.Length -gt $LargeFileThreshold) {
                # Large file - use chunked upload
                Write-Log "Uploading (chunked): $($File.FullName) ($([math]::Round($File.Length/1MB, 2)) MB)"
                $Result = Upload-LargeFileToSharePoint -AccessToken $AccessToken -LocalFilePath $File.FullName `
                    -SPODriveId $SPODriveId -SPOFolder $SPOCurrentPath -ChunkSizeMB $ChunkSizeMB `
                    -ConflictBehavior $ConflictBehavior -ResetArchiveAttribute:$ResetArchiveAttribute -MaxRetries $MaxRetries
            }
            else {
                # Small file - use simple upload
                Write-Log "Uploading: $($File.FullName)"
                Upload-FileToSharePoint -AccessToken $AccessToken -LocalFilePath $File.FullName `
                    -SPOSiteId $SPOSiteId -SPODriveId $SPODriveId -SPOLibraryName $SPOLibraryName `
                    -SPOFolder $SPOCurrentPath -conflictBehavior "Replace" -ResetArchiveAttribute:$ResetArchiveAttribute
                $Result = $true
            }

            if ($Result) {
                $script:SyncStats.FilesUploaded++
                $script:SyncStats.BytesUploaded += $File.Length
            }
            else {
                $script:SyncStats.Errors++
            }
        }
        catch {
            Write-Log "Error uploading $($File.FullName): $($_.Exception.Message)" -MessageType Error
            $script:SyncStats.Errors++
        }
    }

    # Skip files
    foreach ($File in $Comparison.FilesToSkip) {
        Write-Log "Skipping (not modified): $($File.Name)" -ForegroundColor DarkGray
        $script:SyncStats.FilesSkipped++
    }

    # Delete items (Mirror mode only)
    if ($SyncMode -eq "Mirror") {
        foreach ($Item in $Comparison.ItemsToDelete) {
            $ItemType = if ($Item.folder) { "folder" } else { "file" }

            if ($WhatIf) {
                Write-Log "[WhatIf] Would delete $($ItemType): $SPOCurrentPath/$($Item.name)"
                continue
            }

            Write-Log "Deleting $($ItemType): $SPOCurrentPath/$($Item.name)" -ForegroundColor Red
            $Deleted = Remove-SPOItem -AccessToken $AccessToken -SPOSiteId $SPOSiteId -ItemId $Item.id

            if ($Deleted) {
                if ($Item.folder) {
                    $script:SyncStats.FoldersDeleted++
                }
                else {
                    $script:SyncStats.FilesDeleted++
                }
            }
            else {
                $script:SyncStats.Errors++
            }
        }
    }

    # Return statistics at top level
    if (-not $RecursiveCall) {
        return $script:SyncStats
    }
}

