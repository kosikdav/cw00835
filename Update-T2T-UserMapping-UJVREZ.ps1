#######################################################################################################################
# Update-T2T-UserMapping-UJVREZ.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1

$LogFolder			= "t2t-ujvrez"
$LogFilePrefix		= "user-mapping"
$OutputFolder		= "t2t-ujvrez"
$OutputFilePrefix	= "user-mapping"

#######################################################################################################################

. $ScriptPath\include-Script-Start-Include.ps1

$LogFile 	= New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$OutputFile	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Ext "csv" -Freq "YMDHMS"

#######################################################################################################################

if ($InteractiveRun) {
	$ADCredentialPath = "c:\cred\qp_aad_grp_mgmt\qp_aad_grp_mgmt_qskosikdav.cred"
}
else {
	$ADCredentialPath = $aad_grp_mgmt_cred
}

$MappingCSVFilePath = "d:\data\t2t-ujvrez\userMapping.csv"
$Dst_AD_OU_list = @(
	"OU=CVREZ,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=EGP,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=EngineeringPraha,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=iCVREZ,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",	
	"OU=NQ-Safe,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=UJVREZ,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=VZUP,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp"
)

$Src_AppReg_LOG_READER 			= $AppReg_UJVREZ_LOG_READER
$Src_AppReg_EXO_MGMT 			= $AppReg_UJVREZ_EXO_MGMT   
$Src_T2T_EXO_MIGRATION_GROUP 	= $UJVREZ_T2T_EXO_MIGRATION_GROUP

$Dst_AppReg_LOG_READER 			= $AppReg_CEZDATA_LOG_READER
$Dst_AppReg_EXO_MGMT 			= $AppReg_CEZDATA_EXO_MGMT   

$Src_map_attr = "UJV_pn"
$Dst_map_attr = "CEZ_pn"
$Src_PN_attr = "extension_93b54ce056df45bd8f5f398753fa17c0_employeeNumber"
$Dst_PN_attr = "extension_008a5d3f841f4052ac1283ff4782c560_employeeNumber"
$Dst_mailAD40 = "extension_008a5d3f841f4052ac1283ff4782c560_msExchExtensionAttribute40"
$CommonEntraAttributes = "id,userPrincipalName,displayName,onpremisesSamAccountName,mail"

[array]$MappingReport = @()
[hashtable]$mapping_DB = @{}
[hashtable]$Dst_UserDB_per_pn = @{}

$ADCredential = Import-Clixml -Path $ADCredentialPath

#######################################################################################################################

. $IncFile_StdLogStartBlock

#######################################################################################################################
# mapping file
#######################################################################################################################
[array]$userMapping = Import-CSVtoArray -Path $MappingCSVFilePath
write-host "User mapping: $($userMapping.count)"
$userMapping = $userMapping | Where-Object { $_.prio -eq 1 }
write-host "User mapping (prio 1): $($userMapping.count)"

$duplicateSrcMappings = $userMapping | Group-Object -Property $Src_map_attr | Where-Object { $_.Count -gt 1 }
foreach ($group in $duplicateSrcMappings) {
	write-host "Duplicate $($Src_map_attr): $($group.Name) - Count: $($group.Count)"
	foreach ($mapping in $group.Group) {
		write-host $mapping
	}
	exit
}
$duplicateDstMappings = $userMapping | Group-Object -Property $Dst_map_attr | Where-Object { $_.Count -gt 1 }
foreach ($group in $duplicateDstMappings) {
	write-host "Duplicate $($Dst_map_attr): $($group.Name) - Count: $($group.Count)"
	foreach ($mapping in $group.Group) {
		write-host $mapping
	}
	exit
}
foreach ($mapping in $userMapping) {
	$mapping_DB.add($mapping.$Src_map_attr, $mapping.$Dst_map_attr)
}
write-host "User mapping DB: $($mapping_DB.count)"

#######################################################################################################################
# get SRC AAD users
#######################################################################################################################
Request-MSALToken -AppRegName $Src_AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriSelect = "$CommonEntraAttributes,$Src_PN_attr"
$UriFilter = "userType eq 'Member' and onpremisesSyncEnabled eq true"
$Uri = New-GraphUri -Resource $UriResource -Version "v1.0" -Select $UriSelect -Filter $UriFilter -Top 999
[array]$SrcAADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$Src_AppReg_LOG_READER].AccessToken -ProgressDots -Text "SRC AAD users"
write-host "UJVREZ AAD users: $($SrcAADUsers.count)"

#######################################################################################################################
# get DST AAD users
#######################################################################################################################
Request-MSALToken -AppRegName $Dst_AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriSelect = "$CommonEntraAttributes,onPremisesDistinguishedName,onpremisesExtensionAttributes,$Dst_PN_attr,$Dst_mailAD40"
$UriFilter = "userType eq 'Member' and onpremisesSyncEnabled eq true"
$Uri = New-GraphUri -Resource $UriResource -Version "v1.0" -Select $UriSelect -Filter $UriFilter -Top 999
[array]$DstAADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$Dst_AppReg_LOG_READER].AccessToken -ProgressDots -Text "DST AAD users"
write-host "CEZDATA AAD users: $($DstAADUsers.count)"
$DstAADUsers = $DstAADUsers | Where-Object { $_.onpremisesSamAccountName -notlike 'Q*' }
write-host "CEZDATA AAD users (Q): $($DstAADUsers.count)"

write-host "CEZDATA AAD users - adding OU property..." -NoNewline
foreach ($user in $DstAADUsers) {    
	$user | Add-Member -NotePropertyName 'OU' -NotePropertyValue (
		$user.onPremisesDistinguishedName -replace '^CN=[^,]+,'
	)
}
write-host "done"
$DstAADUsers = $DstAADUsers | Where-Object { $_.OU -in $Dst_AD_OU_list }
write-host "CEZDATA AAD users (OU): $($DstAADUsers.count)"

$duplicateUsers = $DstAADUsers | Group-Object -Property "extension_008a5d3f841f4052ac1283ff4782c560_employeeNumber" | Where-Object { $_.Count -gt 1 }
foreach ($group in $duplicateUsers) {
	write-host "Duplicate CEZ_pn: $($group.Name) - Count: $($group.Count)"
	foreach ($user in $group.Group) {
		write-host "  User: $($user.DisplayName) - UPN: $($user.UserPrincipalName) - OU: $($user.OU)"
	}
	exit
}

foreach ($user in $DstAADUsers) {
	$userObject = [PSCustomObject]@{
		id = $user.id
		userPrincipalName = $user.userPrincipalName
		displayName = $user.displayName
		onPremisesSamAccountName = $user.onpremisesSamAccountName
		onPremisesSyncEnabled = $user.onpremisesSyncEnabled
		ext10 = $user.onpremisesExtensionAttributes.extensionAttribute10
		employeeNumber = $user.extension_008a5d3f841f4052ac1283ff4782c560_employeeNumber
		mailAD40 = $user.extension_008a5d3f841f4052ac1283ff4782c560_msExchExtensionAttribute40
	}
	$Dst_UserDB_per_pn.add($userObject.employeeNumber, $userObject)
}
write-host "CEZDATA AAD userDB: $($Dst_UserDB_per_pn.count)"

foreach ($user in $SrcAADUsers) {
	$ReportObject = $null
	if ($user.$Src_PN_attr) {
		$ReportObject = [PSCustomObject]@{
			UJV_UPN = $user.userPrincipalName
			UJV_UPNdomain = $user.userPrincipalName.Split("@")[1]
			UJV_DisplayName = $user.displayName
			UJV_mail = $user.mail
			UJV_samAccountName = $user.onpremisesSamAccountName
			UJV_PN = $user.$Src_PN_attr
			Mapped_PN = $null
			CEZ_UPN = $null
			CEZ_DisplayName = $null
			CEZ_mail = $null
			CEZ_samAccountName = $null
			CEZ_PN = $null
			OldExt10 = $null
			NewExt10 = $null
			CEZ_mailAD40 = $null
			mailAD40match = $null
			Result = $null
		}
		if ($mapping_DB.ContainsKey($user.$Src_PN_attr)) {
			$mapped_pn = $mapping_DB[$user.$Src_PN_attr]
			$ReportObject.mapped_pn = $mapped_pn
			if ($Dst_UserDB_per_pn.ContainsKey($mapped_pn)) {
				$DstUser = $Dst_UserDB_per_pn[$mapped_pn]
				$ReportObject.CEZ_UPN = $DstUser.userPrincipalName
				$ReportObject.CEZ_DisplayName = $DstUser.displayName
				$ReportObject.CEZ_mail = $DstUser.mailAD40
				$ReportObject.CEZ_samAccountName = $DstUser.onPremisesSamAccountName
				$ReportObject.CEZ_PN = $DstUser.employeeNumber
				$ReportObject.CEZ_mailAD40 = $DstUser.mailAD40
				if ($DstUser.mailAD40 -and ($DstUser.mailAD40 -eq $user.mail)) {
					$ReportObject.mailAD40match = "YES"
				}
				else {
					$ReportObject.mailAD40match = "NO"
				}
				if ($DstUser.ext10 -ne $user.$Src_PN_attr) { 
					if ($ReportObject.mailAD40match -eq "YES") {
						if ($DstUser.ext10) {
							$currentExt10 = $DstUser.ext10
						}
						else {
							$currentExt10 = "<empty>"
						}
						try {
							#Set-ADUser -Identity $DstUser.onPremisesSamAccountName -Replace @{extensionAttribute10 = $user.$Src_PN_attr} -Credential $ADCredential
							Write-Log "UPDATING:  $($user.userPrincipalName) ($($user.displayName)) PN:$($user.$Src_PN_attr) - $($DstUser.userPrincipalName) $($DstUser.onPremisesSamAccountName) $($currentExt10) -> $($user.$Src_PN_attr)" -ForegroundColor Yellow
							$ReportObject.OldExt10 = $currentExt10
							$ReportObject.NewExt10 = $user.$Src_PN_attr
							$ReportObject.Result = "UPDATED"
						}
						catch {
							Write-Log "Error updating user: $($user.userPrincipalName) ($($user.displayName)) - $($_.Exception.Message)" -ForegroundColor Red -MessageType "ERR"
						}
					}
					else {
						write-Log "MAILERR:   $($user.userPrincipalName) ($($user.displayName)) UJV_PN:$($user.$Src_PN_attr) - $($DstUser.userPrincipalName) $($DstUser.onPremisesSamAccountName) ext10:$($DstUser.ext10)" -ForegroundColor DarkYellow -MessageType "WARN"
						$ReportObject.OldExt10 = $currentExt10
						$ReportObject.NewExt10 = $user.$Src_PN_attr
						$ReportObject.Result = "MAILERR"
					}
				}
				else {
					write-Log "OK:        $($user.userPrincipalName) ($($user.displayName)) UJV_PN:$($user.$Src_PN_attr) - $($DstUser.userPrincipalName) $($DstUser.onPremisesSamAccountName) ext10:$($DstUser.ext10)" -ForegroundColor Green
					$ReportObject.OldExt10 = $currentExt10
					$ReportObject.NewExt10 = $currentExt10
					$ReportObject.Result = "OK"
				}
			}
			else {
				Write-Log "NOTFOUND:  $($user.userPrincipalName) ($($user.displayName)) UJV_PN:$($user.$Src_PN_attr) - expected CEZ_PN: $($mapped_pn)" -ForegroundColor Red -MessageType "ERR"
				$ReportObject.Result = "NOTFOUND"
			}
		}
		else {
			Write-Log "NOMAPPING: $($user.userPrincipalName) ($($user.displayName)) UJV_PN:$($user.$Src_PN_attr)" -ForegroundColor DarkYellow -MessageType "WARN"
			$ReportObject.Result = "NOMAPPING"
		}
	}
	if ($ReportObject) {
		$MappingReport += $ReportObject
	}
}

Export-Report -Text "T2T user mapping report UJVREZ-CEZDATA" -Report $MappingReport -Path $OutputFile -SortProperty "UJV_UPN"

#######################################################################################################################

. $IncFile_StdLogEndBlock