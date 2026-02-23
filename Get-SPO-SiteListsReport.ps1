#######################################################################################################################
# Get-SPO-SiteListsReport
# Reports all lists from a specific SPO site (and optionally subsites), including item count and combined item size.
# Uses purely Graph API calls via LOG_READER app registration - no PS modules required.
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    [Parameter(Mandatory=$true)][string]$SiteUrl,
    [switch]$IncludeSubsites,
    [switch]$IncludeHiddenLists,
    [Parameter(Mandatory=$true)][string]$OutputFile
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1

#######################################################################################################################

$LogFolder          = "exports"
$LogFilePrefix      = "spo-lists-report"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile    = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

[array]$Report = @()
[System.Collections.Generic.Queue[pscustomobject]]$SiteQueue = [System.Collections.Generic.Queue[pscustomobject]]::new()

#######################################################################################################################

. $IncFile_StdLogStartBlock

Write-Log "SiteUrl:            $SiteUrl"
Write-Log "IncludeSubsites:    $IncludeSubsites"
Write-Log "IncludeHiddenLists: $IncludeHiddenLists"
Write-Log "Output file:        $OutputFile"

############################################################
# Authenticate
############################################################
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

############################################################
# Resolve root site URL to site object
############################################################
Write-Log "Resolving site: $SiteUrl"
$ParsedUrl    = [System.Uri]$SiteUrl
$SiteHostPath = "$($ParsedUrl.Host):$($ParsedUrl.AbsolutePath)"
$Uri          = New-GraphUri -Version "v1.0" -Resource "sites/$SiteHostPath" -Select "id,displayName,webUrl"
$Headers      = @{ Authorization = "Bearer $($AuthDB[$AppReg_LOG_READER].AccessToken)" }

Try {
    $RootSite = Invoke-RestMethod -Method GET -Uri $Uri -Headers $Headers -ContentType $ContentTypeJSON
    Write-Log "Resolved: $($RootSite.displayName) | $($RootSite.webUrl) | $($RootSite.id)"
    $SiteQueue.Enqueue([pscustomobject]@{
        id          = $RootSite.id
        displayName = $RootSite.displayName
        webUrl      = $RootSite.webUrl
    })
}
Catch {
    Write-Log "Failed to resolve site '$SiteUrl': $($_.Exception.Message)" -MessageType Error
    . $IncFile_StdLogEndBlock
    Exit 1
}

############################################################
# BFS over site and subsites - get lists for each
############################################################
while ($SiteQueue.Count -gt 0) {
    $CurrentSite = $SiteQueue.Dequeue()
    Write-Log "Processing site: $($CurrentSite.displayName) ($($CurrentSite.webUrl))"

    # Refresh token before each site to handle long-running scripts
    Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
    $Headers = @{ Authorization = "Bearer $($AuthDB[$AppReg_LOG_READER].AccessToken)" }

    ########################################################
    # Discover and enqueue subsites (if requested)
    ########################################################
    if ($IncludeSubsites) {
        $SubsiteUri = New-GraphUri -Version "v1.0" -Resource "sites/$($CurrentSite.id)/sites" -Select "id,displayName,webUrl"
        [array]$Subsites = Get-GraphOutputREST -Uri $SubsiteUri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Silent
        foreach ($Subsite in $Subsites) {
            Write-Log "  Enqueuing subsite: $($Subsite.displayName) ($($Subsite.webUrl))"
            $SiteQueue.Enqueue([pscustomobject]@{
                id          = $Subsite.id
                displayName = $Subsite.displayName
                webUrl      = $Subsite.webUrl
            })
        }
    }

    ########################################################
    # Get all lists for current site
    ########################################################
    $ListsUri = New-GraphUri -Version "v1.0" -Resource "sites/$($CurrentSite.id)/lists" -Select "id,displayName,webUrl,list,createdDateTime,lastModifiedDateTime"
    [array]$Lists = Get-GraphOutputREST -Uri $ListsUri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "Getting lists for '$($CurrentSite.displayName)'" -ProgressDots

    foreach ($List in $Lists) {
        # Skip hidden lists unless explicitly requested
        if ($List.list.hidden -and (-not $IncludeHiddenLists)) {
            continue
        }

        $ItemCount = 0
        $ListSize  = 0
        $IsDocLib  = ($List.list.template -eq "documentLibrary")

        ####################################################
        # Get item count via $count=true + ConsistencyLevel:eventual
        # Uses $top=1 so only 1 item is transferred; total comes from @odata.count
        ####################################################
        Try {
            $UriResource = "sites/$($CurrentSite.id)/lists/$($List.id)/items"
            $UriSelect = "id"
            $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
            # -Count -Top 1
            #write-host $Uri -ForegroundColor Cyan
            $Items = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ConsistencyLevel "eventual"
        }
        Catch {
            Write-Log "  Warning: Could not get item count for '$($List.displayName)': $($_.Exception.Message)" -MessageType Warning
        }

        ####################################################
        # Get combined size from drive root (document libraries only)
        # drive/root.size = recursive sum of all file sizes in bytes
        ####################################################
        if ($IsDocLib) {
            Try {
                $DriveRootUri = New-GraphUri -Version "v1.0" -Resource "sites/$($CurrentSite.id)/lists/$($List.id)/drive/root" -Select "size"
                $DriveRoot    = Invoke-RestMethod -Method GET -Uri $DriveRootUri -Headers $Headers -ContentType $ContentTypeJSON
                $ListSize     = $DriveRoot.size
            }
            Catch {
                Write-Log "  Warning: Could not get size for '$($List.displayName)': $($_.Exception.Message)" -MessageType Warning
            }
        }

        $Report += [pscustomobject]@{
            SiteDisplayName      = $CurrentSite.displayName
            SiteWebUrl           = $CurrentSite.webUrl
            SiteId               = $CurrentSite.id
            ListDisplayName      = $List.displayName
            ListId               = $List.id
            ListWebUrl           = $List.webUrl
            ListTemplate         = $List.list.template
            IsDocumentLibrary    = $IsDocLib
            Hidden               = $List.list.hidden
            ItemCount            = $Items.Count
            SizeBytes            = $ListSize
            SizeMB               = if ($ListSize -gt 0) { [math]::Round($ListSize / 1MB, 2) } else { 0 }
            CreatedDateTime      = $List.createdDateTime
            LastModifiedDateTime = $List.lastModifiedDateTime
        }

        Write-Host ("  {0,-50} | {1,-20} | Items: {2,6} | Size: {3,10} MB{4}" -f `
            $List.displayName, `
            $List.list.template, `
            $Items.Count, `
            ([math]::Round($ListSize / 1MB, 2)), `
            $(if ($List.list.hidden) { " [hidden]" } else { "" }))
    }
}

Write-Log "Total lists in report: $($Report.Count)"

############################################################
# Export report
############################################################
Export-Report -Text "SPO site lists report" -Report $Report -SortProperty "SiteDisplayName" -Path $OutputFile

#######################################################################################################################

. $IncFile_StdLogEndBlock
