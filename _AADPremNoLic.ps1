#######################################################################################################################
# Get-AAD-Guests-Reports
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "exports"
$LogFilePrefix		= "aad-users-report"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

[array]$AADPremNoLic = @()

#######################################################################################################################

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "groups/$($GroupId_AADPremNoLic)/members"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999
#$UsersAADPremNoLic = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "AADPremNoLic members" -ProgressDots

#$UsersAADPremNoLic | ForEach-Object {$AADPremNoLic += $_.id}
#$AADPremNoLic = $UsersAADPremNoLic.id

#write-host $AADPremNoLic.Count

$AADPremLicense = [string]::Empty
$UserId = "0b1c04f9-50c9-4f65-83ff-861a87516d1f"
$UriResource = "users/$($Userid)/licenseDetails"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
try {
    $Query = Invoke-RestMethod -Headers @{ Authorization = "Bearer $accessToken" } -Uri $Uri -Method GET -ContentType "application/json"
    $licensedSKUs = $Query.Value
    Write-Host $licensedSKUs.count -ForegroundColor Yellow
    :sku foreach ($sku in $licensedSKUs) {
        write-host $sku.skuPartNumber -ForegroundColor Yellow
        write-host $sku.servicePlans.count
        :plan foreach ($plan in $sku.servicePlans) {
            write-host $plan.servicePlanName $plan.servicePlanId -ForegroundColor DarkGray
            if (($AADP1LicensePlans.Contains($plan.servicePlanId)) -and ([string]::IsNullOrEmpty($AADPremLicense))) {
                $AADPremLicense = $plan.servicePlanName
                write-host $plan.servicePlanName
            }
            if ($AADP2LicensePlans.Contains($plan.servicePlanId)) {
                $AADPremLicense = $plan.servicePlanName
                write-host $plan.servicePlanName
            }
            if ($EXOLicensePlans.Contains($plan.servicePlanId)) {
                $EXOLicense = "EXO"
                write-host $plan.servicePlanName
            }
            if ($SPOLicensePlans.Contains($plan.servicePlanId)) {
                $SPOLicense = "SPO"
                write-host $plan.servicePlanName
            }
        }
    }
}
catch {
    $licensedSKUs = "n/a"
}
Write-Host $AADPremLicense -ForegroundColor Green
Write-Host $EXO