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

