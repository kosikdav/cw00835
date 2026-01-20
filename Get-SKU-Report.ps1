#######################################################################################################################
# Get-SKU-Repoert
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder 			= "exports"
$LogFilePrefix		= "sku-report"

$OutputFolder 		= "sku\report"
$OutputFilePrefix	= "sku-report"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$OutputFile = New-OutputFile -RootFolder $REF -Folder $OutputFolder -Prefix $OutputFilePrefix -FileDateYesterday -Ext "csv"

[array]$SKUReport = $null
[hashtable]$MSLicData_DB = @{}

##################################################################################################

. $IncFile_StdLogBeginBlock

$MSLicDataArray = Import-CSVtoArray -Path $MSLicDataPath
$MSLicDataArray = $MSLicDataArray | Sort-Object -Property "GUID" -Unique
foreach ($sku in $MSLicDataArray) {
	$MSLicData_DB.Add($sku.GUID, $sku.Product_Display_Name)
}

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

$subscriptionIds = $prepaidEnabled = $prepaidlockedOut = $prepaidSuspended = $prepaidWarning = $null

$UriResource  = "subscribedSkus"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
$SKUs = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
write-host "SKUs: $($SKUs.Count)"
foreach ($sku in $SKUs) {
	$subscriptionIds = $sku.subscriptionIds -join ","
	$skuDisplayName = $sku.skuPartNumber.Replace("_", " ")
	if ($MSLicData_DB.ContainsKey($sku.skuId)) {
		$skuDisplayName = $MSLicData_DB[$sku.skuId]
	}

	if ($sku.prepaidUnits) {
		$prepaidEnabled =   $sku.prepaidUnits.enabled
		$prepaidlockedOut = $sku.prepaidUnits.lockedOut
		$prepaidSuspended = $sku.prepaidUnits.suspended
		$prepaidWarning =   $sku.prepaidUnits.warning
	}
	$SKUObject = [PSCustomObject]@{
		skuId           = $sku.skuId;
		account 	   	= $sku.accountName;
		accountId	  	= $sku.accountId;
		skuPartNumber   = $sku.skuPartNumber;
		skuDisplayName  = $skuDisplayName;
		subscriptionIds = $subscriptionIds;
		appliesTo       = $sku.appliesTo;
		status          = $sku.capabilityStatus;
		consumedUnits   = $sku.consumedUnits;
		availableUnits  = $prepaidEnabled-$sku.consumedUnits;
		prepaidUnits_enabled    = $prepaidEnabled;
		prepaidUnits_lockedOut  = $prepaidlockedOut;
		prepaidUnits_suspended  = $prepaidSuspended;
		prepaidUnits_warning    = $prepaidWarning
	}
	$SKUReport += $SKUObject
}
#######################################################################################################################

Export-Report -Report $SKUReport -Path $OutputFile -SortProperty "skuPartNumber"

. $IncFile_StdLogEndBlock
