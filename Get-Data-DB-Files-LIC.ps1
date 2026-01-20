#######################################################################################################################
# Get-Data-DB-Files-LIC
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
$LogFilePrefix		= "get-data-db-files-lic"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

[array]$DBReportLicensingInfoSKUs   = $null
[array]$DBReportLicensingInfoPlans  = $null
[hashtable]$MSLicData_DB = @{}
[int]$TTL = 30

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
foreach ($sku in $SKUs) {
    $subscriptionIds = $sku.subscriptionIds -join ","
    $skuDisplayName = [string]::Empty
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
        skuPartNumber   = $sku.skuPartNumber;
        skuDisplayName  = $skuDisplayName;
        subscriptionIds = $subscriptionIds;
        appliesTo       = $sku.appliesTo;
        status          = $sku.capabilityStatus;
        consumedUnits   = $sku.consumedUnits;
        prepaidUnits_enabled    = $prepaidEnabled;
        prepaidUnits_lockedOut  = $prepaidlockedOut;
        prepaidUnits_suspended  = $prepaidSuspended;
        prepaidUnits_warning    = $prepaidWarning
    }
    $DBReportLicensingInfoSKUs += $skuObject
    foreach ($plan in $sku.servicePlans) {
        $PlanObject = [PSCustomObject]@{
            PlanId          = $plan.servicePlanId;
            PlanName        = $plan.servicePlanName;
            PlanStatus      = $plan.provisioningStatus;
            PlanAppliesTo   = $plan.appliesTo;
            skuId           = $sku.skuId;
            skuPartNumber   = $sku.skuPartNumber;
        }
        $DBReportLicensingInfoPlans += $PlanObject
    }
}

Export-Report "DBReportLicesningInfoSKUs" -Report $DBReportLicensingInfoSKUs -Path $DBFileLicensingInfoSKUs
Export-Report "DBReportLicesningInfoPlans" -Report $DBReportLicensingInfoPlans -Path $DBFileLicensingInfoPlans
    
. $IncFile_StdLogEndBlock
