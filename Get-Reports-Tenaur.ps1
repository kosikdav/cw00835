#######################################################################################################################
# Get-Reports-Tenaur
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			    = "exports"
$LogFilePrefix		    = "reports-tenaur"

$OutputFolderTMS	    = "tenaur\teams"
$OutputFolderGRP	    = "tenaur\groups"
$OutputFolderUSR	    = "tenaur\users"

$OutputFilePrefixTMS	= "teams"
$OutputFilePrefixGRP	= "m365-groups"
$OutputFilePrefixUSR	= "aad-users"

$OutputFileSuffixList	= "list-tnr"
$OutputFileSuffixMem    = "members-tnr"
$OutputFileSuffixUSR    = "tnr"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"


$OutputFileGRPList 	= New-OutputFile -RootFolder $ROF -Folder $OutputFolderGRP -Prefix $OutputFilePrefixGRP -Suffix $OutputFileSuffixList -Ext "csv"
$OutputFileGRPMem 	= New-OutputFile -RootFolder $ROF -Folder $OutputFolderGRP -Prefix $OutputFilePrefixGRP -Suffix $OutputFileSuffixMem -Ext "csv"
$OutputFileTMSList 	= New-OutputFile -RootFolder $ROF -Folder $OutputFolderTMS -Prefix $OutputFilePrefixTMS -Suffix $OutputFileSuffixList -Ext "csv"
$OutputFileTMSMem 	= New-OutputFile -RootFolder $ROF -Folder $OutputFolderTMS -Prefix $OutputFilePrefixTMS -Suffix $OutputFileSuffixMem -Ext "csv"
$OutputFileUSR 		= New-OutputFile -RootFolder $ROF -Folder $OutputFolderUSR -Prefix $OutputFilePrefixUSR -Suffix $OutputFileSuffixUSR -Ext "csv"

[array]$ReportGRPList = @()
[array]$ReportGRPMembers = @()
[array]$ReportTMSList = @()
[array]$ReportTMSMembers = @()
[array]$ReportUSRList = @()

#######################################################################################################################

. $IncFile_StdLogBeginBlock


# M365 Groups #########################################################################################################

Write-Log "Getting Tenaur M365 groups report as of: $(Get-Date -Format "yyyy-MM-dd HH:MM:ss")"

$M365Groups = Import-CSVtoArray -Path $DBFileGroupsM365
ForEach ($Group in $M365Groups) {
	if (($Group.Mail.ToLower()).StartsWith("tnr")) {
		$ReportGRPList += $Group
	}
}
Export-Report "Tenaur M365 groups list report" -Report $ReportGRPList -Path $OutputFileGRPList

#-----------------------------------------------------------------------------------

$M365GroupMembers = Import-CSVtoArray -Path $DBFileGroupsMembers
foreach ($Member in $M365GroupMembers) {
	if (($Member.companyName.ToLower()).StartsWith("tenaur") -and $Member.unified) {
		$ReportGRPMembers += $Member
	}
}
Export-Report "Tenaur M365 groups membership report" -Report $ReportGRPMembers -Path $OutputFileGRPMem


# Teams ###############################################################################################################

$Teams = Import-CSVtoArray -Path $DBFileTeams
foreach ($Team in $Teams) {
	if (($team.Mail.ToLower()).StartsWith("tms_tnr")) {
		$ReportTMSList += $Team
	}
}
Export-Report "Tenaur Teams list report" -Report $ReportTMSList -Path $OutputFileTMSList

#-----------------------------------------------------------------------------------

$TeamMembers = Import-CSVtoArray -Path $DBFileTeamsMembers
foreach ($Member in $TeamMembers) {
	if (($Member.companyName.ToLower()).StartsWith("tenaur")) {
		$ReportTMSMembers += $Member
	}
}
Export-Report "Tenaur Teams membership report" -Report $ReportTMSMembers -Path $OutputFileTMSMem


# Users ###############################################################################################################

$Users = Import-CSVtoArray -Path $DBFileUsers
foreach ($User in $Users) {
	if (($user.companyName.ToLower()).StartsWith("tenaur")) {
		$ReportUSRList += $User
	}
}
Export-Report "Tenaur users report" -Report $ReportUSRList -Path $OutputFileUSR

. $IncFile_StdLogEndBlock