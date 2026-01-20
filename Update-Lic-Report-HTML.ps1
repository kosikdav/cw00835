#######################################################################################################################
# Update-Lic-Report-HTML
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    [string]$workloads
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "db"
$LogFilePrefix		= "update-lic-report-html"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

[array]$LicReport = @()
[hashtable]$MSLicData_DB = @{}
[int]$TTL = 30
$outFile = $M365LicFolder + "\default.htm"
$HighlightedSKUs = @(
    "Microsoft 365 E3",
    "Microsoft 365 F3",
    "Microsoft Copilot for Microsoft 365"
)
#######################################################################################################################

. $IncFile_StdLogStartBlock

Invoke-WebRequest $MSLicDataUrl -OutFile $MSLicDataPath
$MSLicDataArray = Import-CSVtoArray -Path $MSLicDataPath
$MSLicDataArray = $MSLicDataArray | Sort-Object -Property "GUID" -Unique
foreach ($sku in $MSLicDataArray) {
    $MSLicData_DB.Add($sku.GUID, $sku.Product_Display_Name)
}

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
$subscriptionIds = $prepaidEnabled = $prepaidlockedOut = $prepaidSuspended = $prepaidWarning = $null
$UriResource  = "subscribedSkus"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
[array]$SKUs = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON

$DateGenerated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
# Convert to HTML table with some styling
$html = @"
<html>
<head>
    <title>CEZ M365 Licensing</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        table { border-collapse: collapse; width: 60%; }
        th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
        t    /* Highlight styles */
        .low { background-color: #ffdddd; }    /* light red */
        .high { background-color: #ddffdd; }   /* light green */
        .notice { background-color: #fff7cc; } /* light yellow */
    </style>
</head>
<body>
    <h2>M365 Licensing Status Report</h2>
    <p>Report updated: $DateGenerated</p>
    <table>
        <tr>
            <th>Display Name</th>
            <th>Total licenses</th>
            <th>Consumed licenses</th>
            <th>Remaining licenses</th>
        </tr>
"@

foreach ($sku in $SKUs) {
    $skuDisplayName = $appliesTo = $status = [string]::Empty
    $prepaidEnabled = $consumedUnits = $remainingUnits = $null
    if ($sku.prepaidUnits -and $sku.capabilityStatus -eq "Enabled" -and $sku.appliesTo -eq "User") {
        if ($MSLicData_DB.ContainsKey($sku.skuId)) {
            $DisplayName = $MSLicData_DB[$sku.skuId]
            $appliesTo = $sku.appliesTo
            $prepaidEnabled = $sku.prepaidUnits.enabled
            $consumedUnits = $sku.consumedUnits
            $remainingUnits = $prepaidEnabled - $sku.consumedUnits
            $SKUObject = [PSCustomObject]@{
                displayName     = $DisplayName
                totalUnits      = $prepaidEnabled
                consumedUnits   = $consumedUnits
                remainingUnits  = $remainingUnits
            }
            $LicReport += $skuObject
        }
    }
}
$LicReport = $LicReport | Sort-Object -Property displayName
foreach ($item in $LicReport) {
    $rowClass = ""
    $fontColor = "black"
    $fontStyle = "normal"
    if ($HighlightedSKUs -contains $item.displayName) {
        $rowClass = "notice"
        $fontStyle = "bold"
        if ($item.remainingUnits -le 10) {
            $fontColor = "red"
        } elseif ($item.remainingUnits -ge 60) {
            $fontColor = "green"
        }
    }
    $html += "        <tr class='$rowClass'><td style='font-weight:$fontStyle;'>$($item.displayName)</td><td>$($item.totalUnits)</td><td>$($item.consumedUnits)</td><td style='color:$fontColor; font-weight:$fontStyle;'>$($item.remainingUnits)</td></tr>`n"
}   

# Close the HTML
$html += @"
    </table>
    <p>This report provides an overview of the licensing status for M365 services.</p>
</body>
</html>
"@

# Output to file

$html | Out-File -FilePath $outFile -Encoding UTF8
    
. $IncFile_StdLogEndBlock
