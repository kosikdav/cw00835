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
$LogFilePrefix		= "users-nomapping"
$OutputFolder		= "t2t-ujvrez"
$OutputFilePrefix	= "users-nomapping"

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

$Dst_AD_OU_list = @(
	"OU=CVREZ,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=EGP,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=EngineeringPraha,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=iCVREZ,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",	
	"OU=NQ-Safe,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=UJVREZ,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=VZUP,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp"
)

$Dst_AppReg_LOG_READER 			= $AppReg_CEZDATA_LOG_READER
$Dst_AppReg_EXO_MGMT 			= $AppReg_CEZDATA_EXO_MGMT   

$Src_map_attr = "UJV_pn"
$Dst_map_attr = "CEZ_pn"
$Src_PN_attr = "extension_93b54ce056df45bd8f5f398753fa17c0_employeeNumber"
$Dst_PN_attr = "employeeNumber"
$Dst_mailAD40 = "msExchExtensionAttribute40"
$CommonEntraAttributes = "id,userPrincipalName,displayName,onpremisesSamAccountName,mail"

[array]$MappingReport = @()

$ADCredential = Import-Clixml -Path $ADCredentialPath

#######################################################################################################################

. $IncFile_StdLogStartBlock

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
$DstADUsers = Get-ADUser -Filter {Enabled -eq $true -and ObjectClass -eq "user"} -SearchBase "OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp" -Properties displayName,mail,extensionAttribute10,employeeNumber,employeeId,msExchExtensionAttribute40 -Credential $ADCredential | Select-Object userPrincipalName,displayName,samAccountName,extensionAttribute10,employeeId,employeeNumber,msExchExtensionAttribute40,DistinguishedName
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

$UsersByKIP = $DstADUsers | Group-Object -Property "employeeId"
foreach ($group in $UsersByKIP) {
	$mappedUsers = $group.Group | Where-Object { $_.extensionAttribute10 -ne $null }
	if ($mappedUsers.count -eq 0) {
		foreach ($user in $group.Group) {
			write-host "$($user.UserPrincipalName) ($($user.DisplayName)) KIP: $($user.employeeId) OU: $($user.OU)"
			$MappingReport += [PSCustomObject]@{
				employeeId = $user.employeeId
				displayName = $user.DisplayName
				userPrincipalName = $user.UserPrincipalName
				OU = $user.OU
			}
		}
	}
}

#######################################################################################################################

Export-Report -Text "T2T users without mapping UJVREZ-CEZDATA by KIP" -Report $MappingReport -Path $OutputFile -SortProperty "UJV_UPN"

. $IncFile_StdLogEndBlock