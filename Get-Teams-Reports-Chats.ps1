#######################################################################################################################
# Get-Teams-Reports-Chats
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "exports"
$LogFilePrefix		= "teams-reports"
$LogFileSuffix		= "chats"

$OutputFolder		= "teams\reports"
$OutputFilePrefix	= "teams"
$OutputFileSuffix   = "members-mimoradky"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Suffix $LogFileSuffix -Ext "log"

$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffix -Ext "csv"

$chatId="19:e292ebccb63c45b2b44405b326ad93e2@thread.v2"
[array]$TeamsChatMembersReport = $null

##################################################################################################

. $IncFile_StdLogStartBlock

Write-Log "Teams chat members report:  $($OutputFile)"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "chats/$($chatId)/members"
$Uri = New-GraphUri -Version "beta" -Resource $UriResource
$ChatMembers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "chat members" -ProgressDots

Initialize-ProgressBarMain -Activity "Building chat members report" -Total $ChatMembers.Count
foreach ($ChatMember in $ChatMembers) {
    Update-ProgressBarMain
    Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
	$UserId = $ChatMember.userId
	$Role = "member"
	if ($ChatMember.Roles.Contains("owner")) {
		$Role = "owner"
	}
	$UserProperties = "id,AccountEnabled,Mail,CreatedDateTime,UserPrincipalName,DisplayName,UserType,companyName,department,onPremisesSamAccountName,signInActivity"
	$User = Get-UserFromGraphById -id $UserId -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -Properties $UserProperties -Version "v1.0"
    if ($User) {
		$SIA = $User.signInActivity
		$Mail = "n/a"
		if ($User.Mail) {
			$Mail = $User.Mail
		}
		if (($null -eq $SIA.LastSignInDateTime) -or ($SIA.LastSignInDateTime -eq '1/1/0001 1:00:00 AM')){
			$LastSignInDateTime = "never"
			$DaysSinceLastSignIn = "n/a"
		} 
		Else {
			$LastSignInDateTime	= [DateTime]$SIA.LastSignInDateTime
			$DaysSinceLastSignIn = (New-TimeSpan -Start $SIA.LastSignInDateTime -End $Today).Days
		}
		$TeamsChatMembersReport += [pscustomobject]@{
			Enabled					= $User.AccountEnabled;
			Role					= $Role
			UserType 				= $User.UserType; 
			UserPrincipalName 		= $User.UserPrincipalName; 
			UPNDomain				= $User.UserPrincipalName.Split("@")[1]; 
			KPJM                    = $User.onPremisesSamAccountName;
			DisplayName 			= $User.DisplayName; 
			Mail 					= $Mail; 
			MailDomain	 			= $Mail.Split("@")[1];
			CompanyName				= $User.companyName;
			Department				= $User.department;
			CreatedDateTime 		= $User.CreatedDateTime;
			LastSignIn				= $LastSignInDateTime;
			DaysSinceLastSignIn		= $DaysSinceLastSignIn;
			visibleHistoryStart     = $ChatMember.visibleHistoryStartDateTime
		}
	}
}
Complete-ProgressBarMain

#######################################################################################################################

Export-Report -Text "chat members report" -Report $TeamsChatMembersReport -Path $OutputFile

. $IncFile_StdLogEndBlock