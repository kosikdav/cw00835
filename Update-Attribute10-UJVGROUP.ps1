#######################################################################################################################
# Update-AD-Groups.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1
. $ScriptPath\include-Script-Start-Include.ps1

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
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################

$Src_AppReg_LOG_READER 			= $AppReg_UJVREZ_LOG_READER
$Src_AppReg_EXO_MGMT 			= $AppReg_UJVREZ_EXO_MGMT   
$Src_T2T_EXO_MIGRATION_GROUP 	= $UJVREZ_T2T_EXO_MIGRATION_GROUP

$Dst_AppReg_LOG_READER 			= $AppReg_CEZDATA_LOG_READER
$Dst_AppReg_EXO_MGMT 			= $AppReg_CEZDATA_EXO_MGMT   


$ADCredential = Import-Clixml -Path $ADCredentialPath

[array]$userMapping = Import-CSVtoArray -Path $MappingCSVFilePath
write-host "User mapping: $($userMapping.count)"
$userMapping = $userMapping | Where-Object { $_.prio -eq 1 }
write-host "User mapping: $($userMapping.count)"

Request-MSALToken -AppRegName $Src_AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriSelect = "id,userPrincipalName,displayName,onpremisesSamAccountName,mail,onpremisesSyncEnabled,extension_93b54ce056df45bd8f5f398753fa17c0_employeeNumber"
$UriFilter = "userType eq 'Member' and onpremisesSyncEnabled eq true"
$Uri = New-GraphUri -Resource $UriResource -Version "v1.0" -Select $UriSelect -Filter $UriFilter -Top 999
[array]$SrcAADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$Src_AppReg_LOG_READER].AccessToken -ProgressDots -Text "SRC AAD users"
write-host "UJVREZ AAD users: $($SrcAADUsers.count)"

Request-MSALToken -AppRegName $Dst_AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriSelect = "id,userPrincipalName,displayName,onpremisesSamAccountName,onpremisesSyncEnabled,onPremisesDistinguishedName,onpremisesExtensionAttributes,extension_008a5d3f841f4052ac1283ff4782c560_employeeNumber,extension_008a5d3f841f4052ac1283ff4782c560_msExchExtensionAttribute40"
$UriFilter = "userType eq 'Member' and onpremisesSyncEnabled eq true"
$Uri = New-GraphUri -Resource $UriResource -Version "v1.0" -Select $UriSelect -Filter $UriFilter -Top 999
[array]$DstAADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$Dst_AppReg_LOG_READER].AccessToken -ProgressDots -Text "DST AAD users"
write-host "CEZDATA AAD users: $($DstAADUsers.count)"
$DstAADUsers = $DstAADUsers | Where-Object { $_.onpremisesSamAccountName -notlike 'Q*' }
write-host "CEZDATA AAD users (Q): $($DstAADUsers.count)"

<#
foreach ($user in $DstAADUsers) {
	$userObject = [PSCustomObject]@{
		id = $user.id
		UserPrincipalName = $user.userPrincipalName
		DisplayName = $user.displayName
		OnPremisesSamAccountName = $user.onpremisesSamAccountName
		OnPremisesSyncEnabled = $user.onpremisesSyncEnabled
		ext10 = $user.onpremisesExtensionAttributes.extensionAttribute10
		EmployeeNumber = $user.extension_008a5d3f841f4052ac1283ff4782c560_employeeNumber
		MailAD40 = $user.extension_008a5d3f841f4052ac1283ff4782c560_msExchExtensionAttribute40
	}
}
#>

write-host "CEZDATA AAD adding OU property..." -NoNewline
foreach ($user in $DstAADUsers) {    
	$user | Add-Member -NotePropertyName 'OU' -NotePropertyValue (
		$user.onPremisesDistinguishedName -replace '^CN=[^,]+,'
	)
	#write-host $user
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
}
