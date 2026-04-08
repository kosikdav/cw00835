#######################################################################################################################
# Update-ADO-Guest-Group-Owners.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Init.ps1

#######################################################################################################################

$LogFolder			= "aad-group-mgmt"
$LogFilePrefix		= "ado-guest-group-owners"

#######################################################################################################################

. $ScriptPath\include-Script-Start-Include.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

#######################################################################################################################

. $IncFile_StdLogStartBlock

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

$UriResource = "groups/$($GroupId_CEZ_OwnerGroup_ADO_Guest_Group_Mgmt)/members"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
[array]$ManagedOwners = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON 
$TOBEOwners = $ManagedOwners.id | Sort-Object -Unique

$UriResource = "groups"
$Uriselect = "id,displayName,onPremisesSyncEnabled,groupTypes,resourceProvisioningOptions,isAssignableToRole"
$UriFilter = "startswith(displayName,'ADO_') and endswith(displayName,'_Guest')"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $Uriselect -Filter $UriFilter
[array]$ADOGuestGroups = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "Getting AAD groups" -ProgressDots -ConsistencyLevel "eventual"

write-host "Processing $($ADOGuestGroups.Count) ADO guest groups"
foreach ($group in $ADOGuestGroups) {
	$UriResource = "groups/$($group.id)/owners"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
	[array]$Owners = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON 
	$ASISOwners = $Owners.id | Sort-Object -Unique

	$Difference = Compare-Object -ReferenceObject $ASISOwners -DifferenceObject $TOBEOwners
	if ($Difference) {
		Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
		write-log "Updating owners for group $($group.displayName)"
		$missingOwners = $TOBEOwners | Where-Object { $ASISOwners -notcontains $_ }
		$extraOwners = $ASISOwners | Where-Object { $TOBEOwners -notcontains $_ }
		if ($missingOwners) {
			foreach ($id in $missingOwners) {
				Write-Log "Adding $($id) to $($group.displayName)"
				Add-GraphGroupOwnerById -GroupId $group.id -userId $id -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -SkipCurrentOwners:$true
			}
		}
		if ($extraOwners) {
			foreach ($id in $extraOwners) {
				Write-Log "Removing $($id) from $($group.displayName)"
				Remove-GraphGroupOwnerById -GroupId $group.id -userId $id -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken
			}
		}
	}
}

#######################################################################################################################

. $IncFile_StdLogEndBlock