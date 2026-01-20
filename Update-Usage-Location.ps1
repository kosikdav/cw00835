#######################################################################################################################
# Update-USage-Location
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "usage-location"
$LogFilePrefix		= "update-usage-location"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$LogFileMin = New-OutputFile -RootFolder $ROF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

#######################################################################################################################

. $IncFile_StdLogStartBlock

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriSelect = "id,userPrincipalName,displayName,usageLocation"
$UriFilter = "usageLocation+eq+null&`$count=true"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect -Filter $UriFilter
$AADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ConsistencyLevel "eventual" -ProgressDots -Text "licensed AAD users"
Write-Log "AAD users without usage location: $(Get-Count($AADUsers))"

if ($AADUsers.Count -gt 0) {
	foreach ($User in $AADUsers) {
		Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
		$UriResource = "users/$($User.id)"
		$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
		$GraphBody = [PSCustomObject]@{
			usageLocation = "CZ"
		} | ConvertTo-Json
		Try{
			$ResultUpdate = Invoke-RestMethod -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Method "PATCH" -ContentType $ContentTypeJSON -Body $GraphBody
			$LogEntry = "$($User.userPrincipalName): usage location updated to CZ"
			Write-Log $LogEntry
			Write-Log $LogEntry -AlternateLogfile $LogFileMin

		}
		Catch {
			write-log $_.Exception.Message -MessageType "Error"
		}
	}
}

#######################################################################################################################

. $IncFile_StdLogEndBlock
