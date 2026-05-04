#######################################################################################################################
# Update-T2T-UserMapping-UJVREZ.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1

$LogFolder				= "t2t-ujvrez"
$LogFilePrefix			= "user-mapping"
$OutputFolder			= "t2t-ujvrez"
$OutputFilePrefix		= "user-mapping"
$OutputFileSuffixSRC 	= "src"
$OutputFileSuffixDST 	= "dst-no-mapping"

$SGUMFolder				= $OutputFolder+"\sgum"

#######################################################################################################################

. $ScriptPath\include-Script-Start-Include.ps1

$LogFile 	= New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$OutputFileSRC	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixSRC -Ext "csv" -Freq "YMDHMS"
$OutputFileDST	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixDST -Ext "csv" -Freq "YMDHMS"

$SGUMFile	= New-OutputFile -RootFolder $ROF -Folder $SGUMFolder -Prefix $OutputFilePrefix -Ext "sgum" -Freq "YMDHMS"

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
$CommonEntraAttributes = "id,userPrincipalName,displayName,onPremisesSamAccountName,mail"

[array]$MappingReportSRC = @()
[array]$MappingReportDST = @()
[array]$SGUMReport = @()

[hashtable]$mapping_DB = @{}
[hashtable]$Dst_UserDB_per_pn = @{}

$ADCredential = Import-Clixml -Path $ADCredentialPath

#######################################################################################################################

. $IncFile_StdLogStartBlock

write-log "SRC MAPPING" -ForegroundColor Cyan
Request-MSALToken -AppRegName $Src_AppReg_LOG_READER -TTL 30

#######################################################################################################################
# mapping file
#######################################################################################################################
[array]$userMappingALL = Import-CSVtoArray -Path $MappingCSV_ALL_FilePath
[array]$userMappingENGPRAHA = Import-CSVtoArray -Path $MappingCSV_ENGPRAHA_FilePath

$userMapping = $userMappingALL + $userMappingENGPRAHA

write-host "User mapping: $($userMapping.count)"
$userMapping = $userMapping | Where-Object { $_.prio -eq 1 }
write-host "User mapping (prio 1): $($userMapping.count)"

write-host "Checking duplicate $($Src_map_attr) in mapping file..." -NoNewline
$duplicateSrcMappings = $userMapping | Group-Object -Property $Src_map_attr | Where-Object { $_.Count -gt 1 }
foreach ($group in $duplicateSrcMappings) {
	write-host "Duplicate $($Src_map_attr): $($group.Name) - Count: $($group.Count)"
	foreach ($mapping in $group.Group) {
		write-host $mapping
	}
	exit
}
write-host "done"

write-host "Checking duplicate $($Dst_map_attr) in mapping file..." -NoNewline
$duplicateDstMappings = $userMapping | Group-Object -Property $Dst_map_attr | Where-Object { $_.Count -gt 1 }
foreach ($group in $duplicateDstMappings) {
	write-host "Duplicate $($Dst_map_attr): $($group.Name) - Count: $($group.Count)"
	foreach ($mapping in $group.Group) {
		write-host $mapping
	}
	exit
}
write-host "done"

foreach ($mapping in $userMapping) {
	$mapping_DB.add($mapping.$Src_map_attr, $mapping.$Dst_map_attr)
}
write-host "User mapping DB: $($mapping_DB.count)"

#######################################################################################################################
# get SRC AAD users
#######################################################################################################################
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
write-host "DST AD users - checking duplicate ext10..." -NoNewline
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
write-host "DST AD users..." -NoNewline
$DstADUsers = Get-ADUser -Filter {Enabled -eq $true -and ObjectClass -eq "user"} -SearchBase "OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp" -Properties displayName,mail,extensionAttribute10,employeeNumber,msExchExtensionAttribute40 -Credential $ADCredential | Select-Object userPrincipalName,displayName,samAccountName,extensionAttribute10,employeeNumber,msExchExtensionAttribute40,DistinguishedName
write-host "done ($($DstADUsers.count))"

#filter out users with samAccountName starting with Q
$DstADUsers = $DstADUsers | Where-Object { $_.SamAccountName -notlike 'Q*' }
write-host "DST AD users - (filtered out Q?): $($DstADUsers.count)"

#OU property to each user object by parsing it from DistinguishedName, we will need it for filtering users by OU and for reporting
write-host "DST AD users - adding OU property..." -NoNewline
foreach ($user in $DstADUsers) {    
	$user | Add-Member -NotePropertyName 'OU' -NotePropertyValue ( $user.DistinguishedName -replace '^CN=[^,]+,' )
}
write-host "done"

#filter out only users from specific OUs
$DstADUsers = $DstADUsers | Where-Object { $_.OU -in $Dst_AD_OU_list }
write-host "DST AD users (filtered by OU): $($DstADUsers.count)"

#check if we have duplicate employeeNumber in DST AD users, if yes, we cannot proceed as the mapping is based on employeeNumber and it must be unique
write-host "DST AD users - checking duplicate employeeNumber..." -NoNewline
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
write-host "DST AD userDB: $($Dst_UserDB_per_pn.count)"

write-host $string_divider

$countSRCOK = 0
$countSRCUpdated = 0
$countSRCNotFound = 0
$countSRCNoMapping = 0
$countSRCMailErr = 0
$countSRCTotal = 0

foreach ($user in $SrcAADUsers) {
	$ReportObject = $SGUMObject = $null
	if ($user.$Src_PN_attr) {
		$countSRCTotal++
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
				$ReportObject.CEZ_samAccountName = $DstUser.SamAccountName
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
							Write-Log "UPDATING:  $($user.userPrincipalName) ($($user.displayName)) PN:$($user.$Src_PN_attr) - $($DstUser.userPrincipalName) $($DstUser.SamAccountName) ext10:$($currentExt10) -> $($user.$Src_PN_attr)" -ForegroundColor Yellow
							$ReportObject.OldExt10 = $currentExt10
							$ReportObject.NewExt10 = $user.$Src_PN_attr
							$ReportObject.Result = "UPDATED"
							$SGUMObject = [PSCustomObject]@{
								SourceValue = $user.userPrincipalName
								DestinationValue = $DstUser.userPrincipalName
							}
							$countSRCUpdated++
						}
						catch {
							Write-Log "Error updating user: $($user.userPrincipalName) ($($user.displayName)) - $($_.Exception.Message)" -ForegroundColor Red -MessageType "ERR"
						}
					}
					else {
						write-Log "MAILERR:   $($user.userPrincipalName) ($($user.displayName)) UJV_PN:$($user.$Src_PN_attr) - $($DstUser.SamAccountName) $($DstUser.mailAD40) vs $($user.mail)" -ForegroundColor DarkCyan -MessageType "WARN"
						$ReportObject.OldExt10 = $currentExt10
						$ReportObject.NewExt10 = $user.$Src_PN_attr
						$ReportObject.Result = "MAILERR"
						$countSRCMailErr++
					}
				}
				else {
					#write-Log "OK:        $($user.userPrincipalName) ($($user.displayName)) UJV_PN:$($user.$Src_PN_attr) - $($DstUser.userPrincipalName) $($DstUser.SamAccountName) ext10:$($DstUser.ext10)" -ForegroundColor Green
					$ReportObject.OldExt10 = $currentExt10
					$ReportObject.NewExt10 = $currentExt10
					$ReportObject.Result = "OK"
					$SGUMObject = [PSCustomObject]@{
						SourceValue = $user.userPrincipalName
						DestinationValue = $DstUser.userPrincipalName
					}
					$countSRCOK++
				}
			}
			else {
				Write-Log "NOTFOUND:  $($user.userPrincipalName) ($($user.displayName)) UJV_PN:$($user.$Src_PN_attr) - expected CEZ_PN: $($mapped_pn)" -ForegroundColor Red -MessageType "ERR"
				$ReportObject.Result = "NOTFOUND"
				$countSRCNotFound++
			}
		}
		else {
			Write-Log "NOMAPPING: $($user.userPrincipalName) ($($user.displayName)) UJV_PN:$($user.$Src_PN_attr)" -ForegroundColor DarkYellow -MessageType "WARN"
			$ReportObject.Result = "NOMAPPING"
			$countSRCNoMapping++
		}
	}
	if ($ReportObject) {
		$MappingReportSRC += $ReportObject
	}
	if ($SGUMObject) {
		$SGUMReport += $SGUMObject
	}
}

write-host $string_divider

#######################################################################################################################
# get DST AD users
#######################################################################################################################
write-log "DST MAPPING" -ForegroundColor Cyan
#check duplicates in ext10 before proceeding

if ($countSRCUpdated -gt 0) {
	write-host "DST AD users - checking duplicate ext10..." -NoNewline
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
	#read DST AD users and filter only those with enabled account
	write-host "DST AD users..." -NoNewline
	$DstADUsers = Get-ADUser -Filter {Enabled -eq $true -and ObjectClass -eq "user"} -SearchBase "OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp" -Properties displayName,mail,extensionAttribute10,employeeNumber,employeeId,msExchExtensionAttribute40 -Credential $ADCredential | Select-Object userPrincipalName,displayName,samAccountName,extensionAttribute10,employeeId,employeeNumber,msExchExtensionAttribute40,DistinguishedName
	write-host "done ($($DstADUsers.count))"

	#filter out users with samAccountName starting with Q
	$DstADUsers = $DstADUsers | Where-Object { $_.SamAccountName -notlike 'Q*' }
	write-host "DST AD users - (filtered out Q?): $($DstADUsers.count)"

	#OU property to each user object by parsing it from DistinguishedName, we will need it for filtering users by OU and for reporting
	write-host "DST AD users - adding OU property..." -NoNewline
	foreach ($user in $DstADUsers) {    
		$user | Add-Member -NotePropertyName 'OU' -NotePropertyValue ( $user.DistinguishedName -replace '^CN=[^,]+,' )
	}
	write-host "done"

	#filter out only users from specific OUs
	$DstADUsers = $DstADUsers | Where-Object { $_.OU -in $Dst_AD_OU_list }
	write-host "DST AD users (filtered by OU): $($DstADUsers.count)"

	#check if we have duplicate employeeNumber in DST AD users, if yes, we cannot proceed as the mapping is based on employeeNumber and it must be unique
	write-host "DST AD users - checking duplicate employeeNumber..." -NoNewline
	$duplicateUsers = $DstADUsers | Group-Object -Property "employeeNumber" | Where-Object { $_.Count -gt 1 }
	foreach ($group in $duplicateUsers) {
		write-host "Duplicate CEZ_pn: $($group.Name) - Count: $($group.Count)"
		foreach ($user in $group.Group) {
			write-host "  User: $($user.DisplayName) - UPN: $($user.UserPrincipalName) - OU: $($user.OU)"
		}
		exit
	}
	write-host "done"
}

$CountDSTNoMapping = 0
$UsersByKIP = $DstADUsers | Group-Object -Property "employeeId"
foreach ($group in $UsersByKIP) {
	$mappedUsers = $group.Group | Where-Object { $_.extensionAttribute10 -ne $null }
	if ($mappedUsers.count -eq 0) {
		foreach ($user in $group.Group) {
			write-host "$($user.UserPrincipalName) ($($user.DisplayName)) KIP: $($user.employeeId) OU: $($user.OU)"
			$MappingReportDST += [PSCustomObject]@{
				employeeId = $user.employeeId
				displayName = $user.DisplayName
				userPrincipalName = $user.UserPrincipalName
				OU = $user.OU
			}
			$CountDSTNoMapping++
		}
	}
}

#######################################################################################################################

Write-Host
Write-Log "SRC mapping summary:" -ForegroundColor Cyan
Write-Log "----------------------------" -ForegroundColor Cyan
Write-Log "Total:     $countSRCTotal"
Write-Log "OK:        $($countSRCOK) $(($countSRCOK/$countSRCTotal*100).ToString("##.##"))%" -foregroundcolor green
Write-Log "Updated:   $($countSRCUpdated) $(($countSRCUpdated/$countSRCTotal*100).ToString("##.##"))%" -foregroundcolor yellow
Write-Log "NotFound:  $($countSRCNotFound) $(($countSRCNotFound/$countSRCTotal*100).ToString("##.##"))%" -foregroundcolor red
Write-Log "NoMapping: $($countSRCNoMapping) $(($countSRCNoMapping/$countSRCTotal*100).ToString("##.##"))%" -foregroundcolor darkyellow
Write-Log "MailErr:   $($countSRCMailErr) $(($countSRCMailErr/$countSRCTotal*100).ToString("##.##"))%" -foregroundcolor darkcyan
Write-Host

Write-Log "DST mapping summary:" -ForegroundColor Cyan
Write-Log "----------------------------" -ForegroundColor Cya
Write-Log "Total:     $($DstADUsers.count)"
Write-Log "NoMapping: $($CountDSTNoMapping) $(($CountDSTNoMapping/$DstADUsers.count*100).ToString("##.##"))%" -foregroundcolor darkyellow
Write-Host

Export-Report -Text "SRC - T2T user mapping report UJVREZ-CEZDATA" -Report $MappingReportSRC -Path $OutputFileSRC -SortProperty "UJV_UPN"
Export-Report -Text "SRC - T2T user mapping SGUM file UJVREZ-CEZDATA" -Report $SGUMReport -Path $SGUMFile
Export-Report -Text "DST - users without mapping UJVREZ-CEZDATA by KIP" -Report $MappingReportDST -Path $OutputFileDST -SortProperty "displayName"

#######################################################################################################################

. $IncFile_StdLogEndBlock