#######################################################################################################################
# Get-AAD-Users-Reports.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder					= "exports"
$LogFilePrefix				= "teams-reports-chat-domains"

$OutputFolder				= "teams\reports"
$OutputFilePrefix			= "teams"
$OutputFileSuffix			= "chat-domains"


#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputFile 	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffix -Ext "csv"
$ChatDomainsReport = [System.Collections.Generic.List[Object]]::new()
$DateLimit = (Get-Date).AddDays(-365)

#######################################################################################################################

. $IncFile_StdLogBeginBlock

$AllAADUsers = Import-CSVtoArray -Path $DBFileUsersMemLic 
$TeamsLicensedUsers = $AllAADUsers | Where-Object {($_.TMSLicense -eq "True")}
write-host "Total Teams licensed users: $($TeamsLicensedUsers.Count)"
$Counter = 0
foreach ($User in $TeamsLicensedUsers) {
	Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
	write-host "Processing user: $($User.DisplayName) " -ForegroundColor Green -NoNewline
	$UriResource = "users/$($User.Id)/chats"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
	[array]$Chats = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON 
	if ($Chats.Count -gt 0) {
		Write-Host "($($Chats.Count))" -ForegroundColor Green
		foreach ($Chat in $Chats) {
			if ($Chat.lastUpdatedDateTime -lt $DateLimit) { 
				#write-host "Chat is older than 1 year, skipping" -ForegroundColor DarkGreen
				continue 
			}
			if ($Chat.id) {
				$UriResource = "chats/$($Chat.id)/members"
				$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
				[array]$ChatMembers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
				if ($ChatMembers.Count -gt 0) {
					foreach ($ChatMember in $ChatMembers) {
						if ($ChatMember.email) {
							$MailDomain = $ChatMember.email.Split("@")[1].ToLower()
							if (-not($ChatDomainsReport.Contains($MailDomain))) {
								$Counter++
								write-host "$($MailDomain) ($($Counter))" -ForegroundColor DarkGray
								Add-Content -Path $OutputFile -Value $MailDomain
								$ChatDomainsReport.Add($MailDomain)
							}
						}
					}
				}
			}
		}
	}
	else {
		write-host "(0)" -ForegroundColor DarkGreen
	}
}

Export-Report -Text "AAD users report" -Report $ChatDomainsReport -Path $OutputFile

#######################################################################################################################

. $IncFile_StdLogStartBlock
