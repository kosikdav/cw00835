#######################################################################################################################
# Get-T2T-UsersWithoutMapping-UJVREZ.ps1
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

$MappingCSV_ALL_FilePath = "d:\data\t2t-ujvrez\userMapping.csv"
$MappingCSV_ENGPRAHA_FilePath = "d:\data\t2t-ujvrez\userMapping-engpraha.csv"

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
$Dst_PN_attr = "employeeNumber"
$Dst_mailAD40 = "msExchExtensionAttribute40"
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
[array]$userMappingALL = Import-CSVtoArray -Path $MappingCSV_ALL_FilePath
[array]$userMappingENGPRAHA = Import-CSVtoArray -Path $MappingCSV_ENGPRAHA_FilePath

$userMapping = $userMappingALL + $userMappingENGPRAHA

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

$SrcAADUsers = $SrcAADUsers | Where-Object { $_.$Src_PN_attr -ne $null -and $_.$Src_PN_attr -ne "" }
write-host "SRC AAD users with $($Src_PN_attr): $($SrcAADUsers.count)"

$SrcAADUsers = $SrcAADUsers | Where-Object { $_.userPrincipalName -notlike "ks.*" }
write-host "SRC AAD users without ks.* UPN: $($SrcAADUsers.count)"

write-host "SRC AAD - checking duplicate $($Src_PN_attr)..." -NoNewline
$duplicateSrcUsers = $SrcAADUsers | Group-Object -Property $Src_PN_attr | Where-Object { $_.Count -gt 1 }
foreach ($group in $duplicateSrcUsers) {
	write-host "Duplicate $($Src_PN_attr): $($group.Name) - Count: $($group.Count)"
	foreach ($mapping in $group.Group) {
		write-host $mapping
	}
	exit
}
write-host "done"

#######################################################################################################################
# get DST AD users
#######################################################################################################################

#check duplicates in ext10 before proceeding
write-host "CEZDATA AD users - checking duplicate ext10..." -NoNewline
$DstExt10Users = Get-ADUser -Filter {extensionAttribute10 -like "*"} -SearchBase "OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp" -Properties extensionAttribute10 -Credential $ADCredential
$duplicateUsers = $DstExt10Users | Group-Object -Property "extensionAttribute10" | Where-Object { $_.Count -gt 1 }
foreach ($group in $duplicateUsers) {
	write-host "Duplicate ext10: $($group.Name) - Count: $($group.Count)"
	foreach ($user in $group.Group) {
		write-host "  User: $($user.DisplayName) - UPN: $($user.UserPrincipalName)"
	}
	exit
}
write-host "done"

#read CEZDATA AD users and filter only those with enabled account
write-host "CEZDATA AD users..." -NoNewline
$DstADUsers = Get-ADUser -Filter {Enabled -eq $true -and ObjectClass -eq "user"} -SearchBase "OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp" -Properties displayName,mail,extensionAttribute10,employeeNumber,msExchExtensionAttribute40 -Credential $ADCredential | Select-Object userPrincipalName,displayName,samAccountName,extensionAttribute10,employeeNumber,msExchExtensionAttribute40,DistinguishedName
write-host "done ($($DstADUsers.count))"

#filter out users with samAccountName starting with Q
$DstADUsers = $DstADUsers | Where-Object { $_.SamAccountName -notlike 'Q*' }
write-host "CEZDATA AD users - (filtered out Q?): $($DstADUsers.count)"

#OU property to each user object by parsing it from DistinguishedName, we will need it for filtering users by OU and for reporting
write-host "CEZDATA AD users - adding OU property..." -NoNewline
foreach ($user in $DstADUsers) {    
	$user | Add-Member -NotePropertyName 'OU' -NotePropertyValue ( $user.DistinguishedName -replace '^CN=[^,]+,' )
}
write-host "done"

#filter out only users from specific OUs
$DstADUsers = $DstADUsers | Where-Object { $_.OU -in $Dst_AD_OU_list }
write-host "CEZDATA AD users (filtered by OU): $($DstADUsers.count)"

#check if we have duplicate employeeNumber in CEZDATA AD users, if yes, we cannot proceed as the mapping is based on employeeNumber and it must be unique
write-host "CEZDATA AD users - checking duplicate employeeNumber..." -NoNewline
$duplicateUsers = $DstADUsers | Group-Object -Property "employeeNumber" | Where-Object { $_.Count -gt 1 }
foreach ($group in $duplicateUsers) {
	write-host "Duplicate CEZ_pn: $($group.Name) - Count: $($group.Count)"
	foreach ($user in $group.Group) {
		write-host "  User: $($user.DisplayName) - UPN: $($user.UserPrincipalName) - OU: $($user.OU)"
	}
	exit
}
write-host "done"

foreach ($user in $DstADUsers) {
	$userObject = [PSCustomObject]@{
		userPrincipalName = $user.userPrincipalName
		displayName = $user.displayName
		samAccountName = $user.samAccountName
		ext10 = $user.extensionAttribute10
		employeeNumber = $user.employeeNumber
		mailAD40 = $user.msExchExtensionAttribute40
	}
	$Dst_UserDB_per_pn.add($userObject.employeeNumber, $userObject)
}
write-host "CEZDATA AAD userDB: $($Dst_UserDB_per_pn.count)"

write-host $string_divider

$countOK = 0
$countUpdated = 0
$countNotFound = 0
$countNoMapping = 0
$countMailErr = 0
$countTotal = 0

foreach ($user in $SrcAADUsers) {
	$ReportObject = $null
	if ($user.$Src_PN_attr) {
		$countTotal++
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
							Set-ADUser -Identity $DstUser.samAccountName -Replace @{extensionAttribute10 = $user.$Src_PN_attr} -Credential $ADCredential
							Write-Log "UPDATING:  $($user.userPrincipalName) ($($user.displayName)) PN:$($user.$Src_PN_attr) - $($DstUser.userPrincipalName) $($DstUser.onPremisesSamAccountName) ext10:$($currentExt10) -> $($user.$Src_PN_attr)" -ForegroundColor Yellow
							$ReportObject.OldExt10 = $currentExt10
							$ReportObject.NewExt10 = $user.$Src_PN_attr
							$ReportObject.Result = "UPDATED"
							$countUpdated++
						}
						catch {
							Write-Log "Error updating user: $($user.userPrincipalName) ($($user.displayName)) - $($_.Exception.Message)" -ForegroundColor Red -MessageType "ERR"
						}
					}
					else {
						write-Log "MAILERR:   $($user.userPrincipalName) ($($user.displayName)) UJV_PN:$($user.$Src_PN_attr) - $($DstUser.onPremisesSamAccountName) $($DstUser.mailAD40) vs $($user.mail)" -ForegroundColor DarkCyan -MessageType "WARN"
						$ReportObject.OldExt10 = $currentExt10
						$ReportObject.NewExt10 = $user.$Src_PN_attr
						$ReportObject.Result = "MAILERR"
						$countMailErr++
					}
				}
				else {
					#write-Log "OK:        $($user.userPrincipalName) ($($user.displayName)) UJV_PN:$($user.$Src_PN_attr) - $($DstUser.userPrincipalName) $($DstUser.onPremisesSamAccountName) ext10:$($DstUser.ext10)" -ForegroundColor Green
					$ReportObject.OldExt10 = $currentExt10
					$ReportObject.NewExt10 = $currentExt10
					$ReportObject.Result = "OK"
					$countOK++
				}
			}
			else {
				Write-Log "NOTFOUND:  $($user.userPrincipalName) ($($user.displayName)) UJV_PN:$($user.$Src_PN_attr) - expected CEZ_PN: $($mapped_pn)" -ForegroundColor Red -MessageType "ERR"
				$ReportObject.Result = "NOTFOUND"
				$countNotFound++
			}
		}
		else {
			Write-Log "NOMAPPING: $($user.userPrincipalName) ($($user.displayName)) UJV_PN:$($user.$Src_PN_attr)" -ForegroundColor DarkYellow -MessageType "WARN"
			$ReportObject.Result = "NOMAPPING"
			$countNoMapping++
		}
	}
	if ($ReportObject) {
		$MappingReport += $ReportObject
	}
}

Export-Report -Text "T2T user mapping report UJVREZ-CEZDATA" -Report $MappingReport -Path $OutputFile -SortProperty "UJV_UPN"

#######################################################################################################################

write-host
write-log "Total:     $countTotal"
write-log "OK:        $("{0:F2}" -f $countOK/$countTotal*100)%" -foregroundcolor green
write-log "Updated:   $("{0:F2}" -f $countUpdated/$countTotal*100)%" -foregroundcolor yellow
write-log "NotFound:  $("{0:F2}" -f $countNotFound/$countTotal*100)%" -foregroundcolor red
write-log "NoMapping: $("{0:F2}" -f $countNoMapping/$countTotal*100)%" -foregroundcolor darkyellow
write-log "MailErr:   $("{0:F2}" -f $countMailErr/$countTotal*100)%" -foregroundcolor darkcyan
write-host

. $IncFile_StdLogEndBlock