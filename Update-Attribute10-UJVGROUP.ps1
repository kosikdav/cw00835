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

#######################################################################################################################
#######################################################################################################################
#######################################################################################################################

$Src_AppReg_LOG_READER 			= $AppReg_UJVREZ_LOG_READER
$Src_AppReg_EXO_MGMT 			= $AppReg_UJVREZ_EXO_MGMT   
$Src_T2T_EXO_MIGRATION_GROUP 	= $UJVREZ_T2T_EXO_MIGRATION_GROUP

$Dst_AppReg_LOG_READER 			= $AppReg_CEZDATA_LOG_READER
$Dst_AppReg_EXO_MGMT 			= $AppReg_CEZDATA_EXO_MGMT   

$ADCredential = Import-Clixml -Path $ADCredentialPath

Request-MSALToken -AppRegName $Dst_AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriSelect = "id,userPrincipalName,displayName,onpremisesSamAccountName,mail,onpremisesSyncEnabled,extension_93b54ce056df45bd8f5f398753fa17c0_employeeNumber"
$UriFilter = "userType eq 'Member' and onpremisesSyncEnabled eq true"
$Uri = New-GraphUri -Resource $UriResource -Version "v1.0" -Select $UriSelect -Filter $UriFilter -Top 999
[array]$SrcAADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$Dst_AppReg_LOG_READER].AccessToken -ProgressDots -Text "AAD users"
write-host "UJVREZ AAD users: $($SrcAADUsers.count)"

Request-MSALToken -AppRegName $Dst_AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriSelect = "id,userPrincipalName,displayName,onpremisesSamAccountName,onpremisesSyncEnabled,onpremisesExtensionAttributes,extension_008a5d3f841f4052ac1283ff4782c560_employeeNumber,extension_008a5d3f841f4052ac1283ff4782c560_msExchExtensionAttribute40"
$UriFilter = "userType eq 'Member' and onpremisesSyncEnabled eq true"
$Uri = New-GraphUri -Resource $UriResource -Version "v1.0" -Select $UriSelect -Filter $UriFilter -Top 999
[array]$DstAADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$Dst_AppReg_LOG_READER].AccessToken -ProgressDots -Text "AAD users"
write-host "CEZDATA AAD users: $($DstAADUsers.count)"
$DstAADUsers = $DstAADUsers | Where-Object { $_.onpremisesSamAccountName -notlike 'Q*' }
write-host "CEZDATA AAD users: $($DstAADUsers.count)"