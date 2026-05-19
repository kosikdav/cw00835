
#######################################################################################################################
# Get-MIP-Backup
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Init.ps1

#######################################################################################################################

$LogFolder 					= "mip-backup"
$LogFilePrefix				= "mip-backup"

$OutputFolder 				= "mip-backup"
$OutputFilePrefix			= "mip-backup"

$OutputFileSuffixLabels	= "labels"
$OutputFileSuffixPolicies = "policies"

#######################################################################################################################

. $ScriptPath\include-Script-Start-Include.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputXMLFileLabels = New-OutputFile -RootFolder $OLF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixLabels -Ext "xml"
$OutputXMLFilePolicies = New-OutputFile -RootFolder $OLF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixPolicies -Ext "xml"

#######################################################################################################################

. $IncFile_StdLogStartBlock

. "$ScriptPath\cezdata\include-appreg-CEZDATA_PURVIEW_MGMT.ps1"
Connect-IPPSSession -AppId $ClientId -CertificateThumbprint $Thumbprint -Organization $TenantName

$Labels = Get-Label
write-host "MIP labels found: " $Labels.count -foregroundColor Yellow
foreach($Label in $Labels){
    write-host "Label: " $Label.Name -foregroundColor Cyan
    write-host $Label -ForegroundColor Cyan
}
Export-Clixml -InputObject $Labels -Path $OutputXMLFileLabels

$Polices = Get-LabelPolicy
write-host "MIP policies found: " $Polices.count -foregroundColor Yellow
foreach($Policy in $Polices){
    write-host "Policy: " $Policy.Name -foregroundColor Cyan
    write-host $Policy -ForegroundColor Cyan
}
Export-Clixml -InputObject $Polices -Path $OutputXMLFilePolicies

#######################################################################################################################

. $IncFile_StdLogEndBlock