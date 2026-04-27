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

$ADCredential = Import-Clixml -Path $ADCredentialPath

Set-ADUser -Identity "vlcekpet2" -Replace @{extensionAttribute10 = "1025137"} -Credential $ADCredential
Set-ADUser -Identity "hamplmir1" -Replace @{extensionAttribute10 = "1800312"} -Credential $ADCredential
Set-ADUser -Identity "zazvorkapet" -Replace @{extensionAttribute10 = "1090454"} -Credential $ADCredential
