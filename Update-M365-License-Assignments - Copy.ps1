#######################################################################################################################
# Update-M365-License-Assignments
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "lic-mgmt-M365"
$LogFilePrefix		= "update-license-assignments"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile 	= New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$LogFileMin = New-OutputFile -RootFolder $ROF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$AllowedDirectSKUs = @(
	"f30db892-07e9-47e9-837c-80727f46fd3d",
	"a403ebcc-fae0-4ca2-8c8c-7a907fd6c235",
	"3f9f06f5-3c31-472c-985f-62d9c10ec167",
	"606b54a9-78d8-4298-ad8b-df6ef4481c80"
)
#######################################################################################################################

. $IncFile_StdLogStartBlock

$SKU_DB = Import-CSVToHashDB -Path $DBFileLicensingInfoSKUs -KeyName "skuId"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriSelect = "id,userPrincipalName,displayName,assignedLicenses"
$UriFilter = "assignedLicenses/`$count+ne+0&`$count=true"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Select $UriSelect -Filter $UriFilter
write-host $Uri -ForegroundColor Green
$AADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ConsistencyLevel "eventual" -ProgressDots -Text "licensed AAD users"
write-Log "Licensed AAD users: $($AADUsers.Count)"

foreach ($User in $AADUsers) {
	Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
	$UriResource = "users/$($User.id)/licenseAssignmentStates"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
	try {
		$Query = Invoke-RestMethod -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON
		$licenseAssignmentStates = $Query.Value
	}
	Catch {
		$licenseAssignmentStates = $null
	}
	write-host "$($User.userPrincipalName) $($licenseAssignmentStates.Count)" -ForegroundColor Green
	foreach ($state in $licenseAssignmentStates) {
		if (($null -eq $state.assignedByGroup) -and ( -not ($AllowedDirectSKUs -contains $state.skuid))) {
			Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
			if ($SKU_DB.ContainsKey($state.skuid)) {
				$sku = $SKU_DB[$state.skuid]
			}
			$UriResource = "users/$($User.id)/assignLicense"
			$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
			$GraphBody = [PSCustomObject]@{
				addLicenses = @();
				removeLicenses = @($state.skuid)
			} | ConvertTo-Json
			Try{
				$ResultRemove = Invoke-RestMethod -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Method "POST" -ContentType $ContentTypeJSON -Body $GraphBody
				$LogEntry = "$($User.userPrincipalName): $($state.skuid) $($sku.skupartnumber) $($sku.skuDisplayName) (state:$($state.state)) license removed"
				Write-Log $LogEntry
				Write-Log $LogEntry -AlternateLogfile $LogFileMin
			}
			Catch {
				write-log $_.Exception.Message -MessageType "Error"
			}
		}
	}
}

#######################################################################################################################

. $IncFile_StdLogEndBlock
