#######################################################################################################################
# Get-Teams-Reports
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

$OutputFolder			= "teams\reports"
$OutputFolderCEZTpl		= "teams\reports\TAS"
$OutputFolderMTR		= "teams\reports\MTR"
$OutputFolderPSTN		= "teams\reports\PSTN"
$OutputFilePrefix		= "teams"

$OutputFileSuffixMembers		= "members"
$OutputFileSuffixTmsChnl		= "teams-channels-owners"
$OutputFileSuffixTmsCEZTpl		= "teams-TAS"
$OutputFileSuffixTmsCEZTplMem	= "teams-TAS-members"
$OutputFileSuffixTmsList		= "list"
$OutputFileSuffixCSUsers		= "cs-users"
$OutputFileSuffixPSTN			= "pstn-users"
$OutputFileSuffixStatGrp		= "stats-per-group"
$OutputFileSuffixStatTms		= "stats-per-team"
$OutputFileSuffixStatUsr		= "stats-per-user"
$OutputFileSuffixStatUsrMTR		= "stats-per-user-MTRPro"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputFileTmsChnl 		= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixTmsChnl -Ext "csv"
$OutputFileTmsCEZTpl	= New-OutputFile -RootFolder $ROF -Folder $OutputFolderCEZTpl -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixTmsCEZTpl -Ext "csv"
$OutputFileTmsCEZTplMem	= New-OutputFile -RootFolder $ROF -Folder $OutputFolderCEZTpl -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixTmsCEZTplMem -Ext "csv"
$OutputFileMembers 		= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixMembers -Ext "csv"
$OutputFileTmsList 		= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixTmsList -Ext "csv"
$OutputFileCSUsers		= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixCSUsers -Ext "csv"
$OutputFilePSTNUsers	= New-OutputFile -RootFolder $ROF -Folder $OutputFolderPSTN -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixPSTN -Ext "csv"
$OutputFileStatGrp		= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixStatGrp -Ext "csv"
$OutputFileStatTmsD7	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $($OutputFileSuffixStatTms+"-D7") -Ext "csv"
$OutputFileStatTmsD180	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $($OutputFileSuffixStatTms+"-D180") -Ext "csv"
$OutputFileStatUsr 		= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $($OutputFileSuffixStatUsr+"-D180") -Ext "csv"
$OutputFileStatUsrMTR	= New-OutputFile -RootFolder $ROF -Folder $OutputFolderMTR -Prefix $OutputFilePrefix -Suffix $($OutputFileSuffixStatUsrMTR+"-D180") -Ext "csv"

[array]$ReportTeamsMembers = @()
[array]$ReportTeamsChannelsOwners = @()
[array]$ReportTeamsCEZTpl = @()
[array]$ReportTeamsCEZTplMem = @()
[array]$ReportTeamsList = @()
[System.Collections.ArrayList]$ReportCSUsers = @()
[System.Collections.ArrayList]$ReportPSTNUsers = @()
[array]$ReportStatsUsers = @()
[array]$ReportStatsUsersMTRPro = @()

[hashtable]$M365Group_DB = @{}
[hashtable]$M365GroupStat_DB = @{}

function Initialize-TempVars {
	$script:UPN = [string]::Empty
	$script:MemberUPNDomain = [string]::Empty
	$script:Department = [string]::Empty
	$script:CompanyName = [string]::Empty
	$script:JobTitle = [string]::Empty
	$script:Mail = [string]::Empty
	$script:MailDomain = [string]::Empty
	$script:ChannelMemberRole = "Channel member"
	$script:UserType = [string]::Empty
}

function Set-TempVars {
	param (
		[parameter(Mandatory = $true)]$UserId
	)
	# main function body ##################################
	$script:CurrentUser = $script:AADUsers_DB.Item($UserId)
	If ($script:CurrentUser.mail) {
		$script:Mail = $script:CurrentUser.mail
		$script:MailDomain = $script:Mail.Split("@")[1]
	} 
	If ($script:CurrentUser.userPrincipalName) {
		$script:UPN = $script:CurrentUser.userPrincipalName
		$script:Department = $script:CurrentUser.department
		$script:MemberUPNDomain = $script:UPN.Split("@")[1]
		$script:CompanyName = $script:CurrentUser.companyName
		$script:JobTitle = $script:CurrentUser.jobTitle
	}
}

#######################################################################################################################

. $ScriptPath\include-Script-StartLog-Generic.ps1
write-Log "OutputFileMembers: $($OutputFileMembers)"
write-Log "OutputFileTmsChnl: $($OutputFileTmsChnl)"
write-Log "OutputFileTmsCEZTpl: $($OutputFileTmsCEZTpl)"
write-Log "OutputFileCSUsers: $($OutputFileCSUsers)"
write-Log "OutputFilePSTNUsers: $($OutputFilePSTNUsers)"
write-Log "OutputFileTmsList: $($OutputFileTmsList)"
write-Log "OutputFileStatGrp: $($OutputFileStatGrp)"
write-Log "OutputFileStatTmsD7: $($OutputFileStatTmsD7)"
write-Log "OutputFileStatTmsD180: $($OutputFileStatTmsD180)"
write-Log "OutputFileStatUsr: $($OutputFileStatUsr)"

#######################################################################################################################
# CSOnlineUsers report 
#######################################################################################################################

Connect-Teams -AppRegName $AppReg_TMS_MGMT -TTL 60
$CsOnlineUsers = Get-CsOnlineUser
write-log "CsOnlineUsers: $($CsOnlineUsers.Count)" -ForegroundColor Yellow
foreach ($CsOnlineUser in $CsOnlineUsers) {
	$PhoneNumberAssignment = $null
	if ($CsOnlineUser.LineURI) {
		$result = Get-CsPhoneNumberAssignment -AssignedPstnTargetId $CsOnlineUser.Identity
		$PhoneNumberAssignment = [PSCustomObject]@{
			TelephoneNumber 	= $result.TelephoneNumber;
			OperatorId 			= $result.OperatorId;
			NumberType 			= $result.NumberType;
			ActivationState 	= $result.ActivationState;
			AssignmentCategory 	= $result.AssignmentCategory;
			Capability 			= $result.Capability;
			LocationUpdateSupported = $result.LocationUpdateSupported;
			PstnAssignmentStatus = $result.PstnAssignmentStatus;
			NumberSource 		= $result.NumberSource
		}
		$userObjectMin = [pscustomobject]@{
			UserId 					= $CsOnlineUser.Identity;
			UserPrincipalName 		= $CsOnlineUser.UserPrincipalName;
			Alias 					= $CsOnlineUser.Alias;
			DisplayName 			= $CsOnlineUser.DisplayName;
			Department 				= $CsOnlineUser.Department;
			CompanyName 			= $CsOnlineUser.Company;
			TenantId 				= $CsOnlineUser.TenantId;
			LineUri 				= $CsOnlineUser.LineUri;
			SipAddress 				= $CsOnlineUser.SipAddress;
			FeatureTypes 			= $CsOnlineUser.FeatureTypes -join ";";
			ProvisionedPlan 		= $CsOnlineUser.ProvisionedPlan -join ";";
			EnterpriseVoiceEnabled 	= $CsOnlineUser.EnterpriseVoiceEnabled;
			ExternalAccessPolicy 	= $CsOnlineUser.ExternalAccessPolicy;
	
			TelephoneNumber 		= $PhoneNumberAssignment.TelephoneNumber;
			NumberType 				= $PhoneNumberAssignment.NumberType;
			ActivationState 		= $PhoneNumberAssignment.ActivationState;
			AssignmentCategory 		= $PhoneNumberAssignment.AssignmentCategory;
			Capability 				= $PhoneNumberAssignment.Capability -join ";";
			LocationUpdateSupported = $PhoneNumberAssignment.LocationUpdateSupported;
			PstnAssignmentStatus 	= $PhoneNumberAssignment.PstnAssignmentStatus;
			NumberSource 			= $PhoneNumberAssignment.NumberSource;
		}
		$ReportPSTNUsers += $userObjectMin
	}
	$userObject	= [pscustomobject]@{
		AccountEnabled 			= $CsOnlineUser.AccountEnabled;
		AccountType 			= $CsOnlineUser.AccountType;
		UserId 					= $CsOnlineUser.Identity;
		UserPrincipalName 		= $CsOnlineUser.UserPrincipalName;
		Alias 					= $CsOnlineUser.Alias;
		FirstName 				= $CsOnlineUser.GivenName;
		LastName 				= $CsOnlineUser.LastName;
		DisplayName 			= $CsOnlineUser.DisplayName;
		UserDirSyncEnabled 		= $CsOnlineUser.UserDirSyncEnabled;
		Title 					= $CsOnlineUser.Title;
		CompanyName				= $CsOnlineUser.Company;
		TenantId 				= $CsOnlineUser.TenantId;
		Department 				= $CsOnlineUser.Department;
		HideFromAddressLists 	= $CsOnlineUser.HideFromAddressLists;
		AdmUnitReference 		= $CsOnlineUser.AdministrativeUnitReference -join ";";
		ApplicationAccessPolicy = $CsOnlineUser.ApplicationAccessPolicy;
		CallingLineIdentity 	= $CsOnlineUser.CallingLineIdentity;
		City 					= $CsOnlineUser.City;
		Street 					= $CsOnlineUser.Street;
		StateOrProvince 		= $CsOnlineUser.StateOrProvince;
		PostalCode				= $CsOnlineUser.PostalCode;
		Country 				= $CsOnlineUser.Country;
		CountryAbbreviation 	= $CsOnlineUser.CountryAbbreviation;
		DialPlan 				= $CsOnlineUser.DialPlan;
		EnterpriseVoiceEnabled 	= $CsOnlineUser.EnterpriseVoiceEnabled;
		ExternalAccessPolicy 	= $CsOnlineUser.ExternalAccessPolicy;
		FeatureTypes 			= $CsOnlineUser.FeatureTypes -join ";";
		ProvisionedPlan 		= $CsOnlineUser.ProvisionedPlan -join ";";
		HostingProvider 		= $CsOnlineUser.HostingProvider;
		InterpretedUserType 	= $CsOnlineUser.InterpretedUserType;
		IsSipEnabled 			= $CsOnlineUser.IsSipEnabled;
		LineUri 				= $CsOnlineUser.LineUri;
		TelephoneNumber 		= $PhoneNumberAssignment.TelephoneNumber;
		OperatorId 				= $PhoneNumberAssignment.OperatorId;
		NumberType 				= $PhoneNumberAssignment.NumberType;
		ActivationState 		= $PhoneNumberAssignment.ActivationState;
		AssignmentCategory 		= $PhoneNumberAssignment.AssignmentCategory;
		Capability 				= $PhoneNumberAssignment.Capability -join ";";
		LocationUpdateSupported = $PhoneNumberAssignment.LocationUpdateSupported;
		PstnAssignmentStatus 	= $PhoneNumberAssignment.PstnAssignmentStatus;
		NumberSource 			= $PhoneNumberAssignment.NumberSource;

		OnPremEnterpriseVoiceEnabled = $CsOnlineUser.OnPremEnterpriseVoiceEnabled;
		OnPremHostingProvider 	= $CsOnlineUser.OnPremHostingProvider;
		OnPremLineUri 			= $CsOnlineUser.OnPremLineUri;
		OnPremOptionFlags 		= $CsOnlineUser.OnPremOptionFlags;
		OnPremSIPEnabled 		= $CsOnlineUser.OnPremSIPEnabled;
		OnPremSipAddress 		= $CsOnlineUser.OnPremSipAddress;
		OnlineAudioConferencingRoutingPolicy = $CsOnlineUser.OnlineAudioConferencingRoutingPolicy;
		OnlineDialOutPolicy 	= $CsOnlineUser.OnlineDialOutPolicy;
		OnlineVoiceRoutingPolicy = $CsOnlineUser.OnlineVoiceRoutingPolicy;
		OnlineVoicemailPolicy 	= $CsOnlineUser.OnlineVoicemailPolicy;
		OwnerUrn 				= $CsOnlineUser.OwnerUrn;
		PreferredDataLocation 	= $CsOnlineUser.PreferredDataLocation;
		PreferredLanguage 		= $CsOnlineUser.PreferredLanguage;
		UsageLocation 			= $CsOnlineUser.UsageLocation; 
		ProxyAddresses 			= $CsOnlineUser.ProxyAddresses -join ";";
		ShadowProxyAddresses 	= $CsOnlineUser.ShadowProxyAddresses -join ";";
		SipAddress 				= $CsOnlineUser.SipAddress;
		SipProxyAddress 		= $CsOnlineUser.SipProxyAddress; 
		SoftDeletionTimestamp 	= $CsOnlineUser.SoftDeletionTimestamp; 
		
		TeamsAppPermissionPolicy = $CsOnlineUser.TeamsAppPermissionPolicy;
		TeamsAppSetupPolicy 	= $CsOnlineUser.TeamsAppSetupPolicy;
		TeamsAudioConferencingPolicy = $CsOnlineUser.TeamsAudioConferencingPolicy;
		TeamsCallHoldPolicy 	= $CsOnlineUser.TeamsCallHoldPolicy;
		TeamsCallParkPolicy 	= $CsOnlineUser.TeamsCallParkPolicy;
		TeamsCallingPolicy 		= $CsOnlineUser.TeamsCallingPolicy;
		TeamsCarrierEmergencyCallRoutingPolicy = $CsOnlineUser.TeamsCarrierEmergencyCallRoutingPolicy;
		TeamsChannelsPolicy 	= $CsOnlineUser.TeamsChannelsPolicy;
		TeamsComplianceRecordingPolicy = $CsOnlineUser.TeamsComplianceRecordingPolicy;
		TeamsCortanaPolicy 		= $CsOnlineUser.TeamsCortanaPolicy;
		TeamsEducationAssignmentsAppPolicy = $CsOnlineUser.TeamsEducationAssignmentsAppPolicy;
		TeamsEmergencyCallRoutingPolicy = $CsOnlineUser.TeamsEmergencyCallRoutingPolicy;
		TeamsEmergencyCallingPolicy = $CsOnlineUser.TeamsEmergencyCallingPolicy;
		TeamsEnhancedEncryptionPolicy = $CsOnlineUser.TeamsEnhancedEncryptionPolicy;
		TeamsEventsPolicy 		= $CsOnlineUser.TeamsEventsPolicy;
		TeamsFeedbackPolicy 	= $CsOnlineUser.TeamsFeedbackPolicy;
		TeamsFilesPolicy 		= $CsOnlineUser.TeamsFilesPolicy;
		TeamsIPPhonePolicy 		= $CsOnlineUser.TeamsIPPhonePolicy;
		TeamsMediaLoggingPolicy = $CsOnlineUser.TeamsMediaLoggingPolicy;
		TeamsMeetingBrandingPolicy = $CsOnlineUser.TeamsMeetingBrandingPolicy;
		TeamsMeetingBroadcastPolicy = $CsOnlineUser.TeamsMeetingBroadcastPolicy;
		TeamsMeetingPolicy 		= $CsOnlineUser.TeamsMeetingPolicy;
		TeamsMessagingPolicy 	= $CsOnlineUser.TeamsMessagingPolicy;
		TeamsMobilityPolicy 	= $CsOnlineUser.TeamsMobilityPolicy;
		TeamsNetworkRoamingPolicy = $CsOnlineUser.TeamsNetworkRoamingPolicy;
		TeamsNotificationAndFeedsPolicy = $CsOnlineUser.TeamsNotificationAndFeedsPolicy;
		TeamsOwnersPolicy 		= $CsOnlineUser.TeamsOwnersPolicy;
		TeamsRoomVideoTeleConferencingPolicy = $CsOnlineUser.TeamsRoomVideoTeleConferencingPolicy;
		TeamsSharedCallingRoutingPolicy	= $CsOnlineUser.TeamsSharedCallingRoutingPolicy;
		TeamsShiftsAppPolicy	= $CsOnlineUser.TeamsShiftsAppPolicy;
		TeamsShiftsPolicy 		= $CsOnlineUser.TeamsShiftsPolicy;
		TeamsSurvivableBranchAppliancePolicy = $CsOnlineUser.TeamsSurvivableBranchAppliancePolicy;
		TeamsSyntheticAutomatedCallPolicy = $CsOnlineUser.TeamsSyntheticAutomatedCallPolicy;
		TeamsTargetingPolicy 	= $CsOnlineUser.TeamsTargetingPolicy;
		TeamsTasksPolicy 		= $CsOnlineUser.TeamsTasksPolicy;
		TeamsTemplatePermissionPolicy = $CsOnlineUser.TeamsTemplatePermissionPolicy;
		TeamsUpdateManagementPolicy = $CsOnlineUser.TeamsUpdateManagementPolicy;
		TeamsUpgradeEffectiveMode = $CsOnlineUser.TeamsUpgradeEffectiveMode;
		TeamsUpgradeNotificationsEnabled = $CsOnlineUser.TeamsUpgradeNotificationsEnabled;
		TeamsUpgradeOverridePolicy = $CsOnlineUser.TeamsUpgradeOverridePolicy;
		TeamsUpgradePolicy 	= $CsOnlineUser.TeamsUpgradePolicy;
		TeamsUpgradePolicyIsReadOnly = $CsOnlineUser.TeamsUpgradePolicyIsReadOnly;
		TeamsVdiPolicy 		= $CsOnlineUser.TeamsVdiPolicy;
		TeamsVerticalPackagePolicy = $CsOnlineUser.TeamsVerticalPackagePolicy;
		TeamsVideoInteropServicePolicy = $CsOnlineUser.TeamsVideoInteropServicePolicy;
		TeamsVoiceApplicationsPolicy  = $CsOnlineUser.TeamsVoiceApplicationsPolicy;
		TenantDialPlan 		= $CsOnlineUser.TenantDialPlan;
	}
	$ReportCSUsers += $userObject
}
Export-Report "CSOnlineUsers" -Report $ReportCSUsers -Path $OutputFileCSUsers
Export-Report "PSTNUsers" -Report $ReportPSTNUsers -Path $OutputFilePSTNUsers
Remove-Variable ReportCSUsers
Remove-Variable ReportPSTNUsers

#######################################################################################################################
# TEAMS MEMBERSHIP REPORT + O365Group_DB 
#######################################################################################################################

$AADUsers_DB = Import-CSVtoHashDB -Path $DBFileUsersAllStd -KeyName "id"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
[array]$ReportTeamsMembers = @()
[array]$ReportTeamsList = @()
$UriResource = "groups"
$UriFilter = "groupTypes/any(c:c+eq+'Unified')"
$Uri = New-GraphUri -Version "beta" -Resource $UriResource -Top 999 -Filter $UriFilter
[array]$M365Groups = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "unified groups" -ProgressDots

foreach ($M365Group in $M365Groups) {
	#write-host "Group: $($M365Group.displayName)" -ForegroundColor DarkYellow
	Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
	# TEAM specific info (channels)
	if ($M365Group.resourceProvisioningOptions.Contains("Team")) {
		Write-Host "Team: $($M365Group.displayName)" -ForegroundColor Yellow
		$UriResource = "teams/$($M365Group.id)"
		$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource		
		$TeamDetails = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON

		$UriResource = "teams/$($M365Group.id)/allChannels"
		$UriFilter = "membershipType+ne+'standard'"
		$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Filter $UriFilter 
		[array]$NonStdChannels = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -includeUnknownEnumMembers
		
		$UriResource = "teams/$($M365Group.id)/incomingChannels"
		$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
		[array]$IncomingChannels = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -includeUnknownEnumMembers
		if ($IncomingChannels) {
			$IncomingChannelIds = $IncomingChannels.id
		}
		# Team channels
		if ($NonStdChannels.Count -gt 0) {
			foreach ($Channel in $NonStdChannels) {
				$ChannelType = ($Channel.membershipType).ToLower()
				$TypeColor = "Cyan"
				If ($Channel.id -in $IncomingChannelIds) {
					$ChannelType = "shared - incoming"
				}
				if ($ChannelType.StartsWith("shared")) {
					$TypeColor = "Magenta"
				}
				$UriResource = "teams/$($M365Group.id)/channels/$($Channel.Id)/members"
				$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
				[array]$ChannelMembers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
				# Team channel members/owners

				Write-Host "  Channel: " -ForegroundColor DarkGray -NoNewline
				write-host "$($Channel.displayName) " -ForegroundColor White -NoNewline
				Write-Host "(" -ForegroundColor DarkGray -NoNewline
				Write-Host $ChannelType -ForegroundColor $TypeColor -NoNewline
				write-host ") Members: " -ForegroundColor DarkGray -NoNewline
				write-host "$($ChannelMembers.Count)" -ForegroundColor White
				
				if ($ChannelMembers.Count -gt 0) {
					foreach ($ChannelMember in $ChannelMembers) {
						Initialize-TempVars
						If ($ChannelMember.email) {
							$Mail = $ChannelMember.email
							$MailDomain = $Mail.Split("@")[1]
						}
						Set-TempVars -UserId $ChannelMember.UserId
						If ($ChannelMember.Roles -Contains "Owner") {
							$ChannelMemberRole = "Channel owner"
						}
						$TeamsMembersRecord = [pscustomobject]@{
							TeamId				= $M365Group.id
							TeamDisplayName		= $M365Group.displayName
							TeamMailAddr		= $M365Group.mail
							UserId				= $ChannelMember.UserId
							MemberDisplayname	= $ChannelMember.DisplayName
							MemberUPN 			= $UPN
							MemberUPNDomain		= $MemberUPNDomain
							MemberEmail			= $Mail
							MemberEmailDomain	= $MailDomain
							CompanyName			= $CompanyName
							Department			= $Department
							JobTitle			= $JobTitle
							Channel				= $Channel.DisplayName
							ChannelType			= $ChannelType
							Role				= $ChannelMemberRole
						}
						$ReportTeamsMembers += $TeamsMembersRecord
					}
				}
			}
		}
	} #if ($M365Group.ProvisionedAs -eq "Team")
	
	$UriResource = "groups/$($M365Group.id)/members"
	$Uri = New-GraphUri -Version "beta" -Resource $UriResource
	[array]$AllMemberObjects = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
	
	$Uri = "https://graph.microsoft.com/beta/groups/$($M365Group.id)/owners"
	[array]$AllOwnerObjects = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
	
	$AllOwnersUPN = $AllOwnerObjects.userPrincipalName
	$OwnerCount 	= 0
	$MemberCount 	= 0
	$GuestCount 	= 0

	# Team owners
	if ($AllOwnerObjects.Count -gt 0) {
		ForEach ($Owner in $AllOwnerObjects) {
			Initialize-TempVars
			Set-TempVars -UserId $Owner.Id
				$TeamsMembersRecord = [pscustomobject]@{
					TeamId				= $M365Group.id
					TeamDisplayName		= $M365Group.displayName
					TeamMailAddr		= $M365Group.mail
					UserId				= $Owner.Id
					MemberDisplayname	= $Owner.DisplayName
					MemberUPN 			= $UPN
					MemberUPNDomain		= $MemberUPNDomain
					MemberEmail			= $Mail
					MemberEmailDomain	= $MailDomain
					CompanyName			= $CompanyName
					Department			= $Department
					JobTitle			= $JobTitle
					Channel				= 'n/a'
					ChannelType			= 'n/a'
					Role				= 'Team owner'
				}
				$ReportTeamsMembers += $TeamsMembersRecord
				$OwnerCount++
		} #ForEach ($Owner in $AllOwnerObjects)
	} #if ($AllOwnerObjects.Count -gt 0)
	
	# Team members
	if ($AllMemberObjects.Count -gt 0) {
		ForEach ($Member in $AllMemberObjects) {
			if ($Member.UserPrincipalName -notin $AllOwnersUPN) {
				Initialize-TempVars
				Set-TempVars -UserId $Member.Id
				If ($Member.UserType) {
					$UserType = $Member.UserType.ToLower()
				}
					$TeamsMembersRecord = [pscustomobject]@{
						TeamId 				= $M365Group.Id
						TeamDisplayName		= $M365Group.DisplayName
						TeamMailAddr		= $M365Group.Mail
						UserId				= $Member.Id
						MemberDisplayname	= $Member.DisplayName
						MemberUPN			= $UPN
						MemberUPNDomain		= $MemberUPNDomain
						MemberEmail			= $Mail
						MemberEmailDomain	= $MailDomain
						CompanyName			= $CompanyName
						Department			= $Department
						JobTitle			= $JobTitle
						Channel				= 'n/a'
						ChannelType			= 'n/a'
						Role				= 'Team ' + $UserType
					}
					$ReportTeamsMembers += $TeamsMembersRecord
					if ($Member.UserType -eq "Member") {
						$MemberCount++
					}
					if ($Member.UserType -eq "Guest") {
						$GuestCount++
					}
			} #if ($Member.UserPrincipalName -notin $AllOwnersUPN)
		} #ForEach ($Member in $AllMemberObjects
	} #if ($AllMemberObjects.Count -gt 0)
	
	$M365GroupData = [pscustomobject]@{
		isArchived							= $TeamDetails.isArchived;
		InternalId							= $TeamDetails.internalId;
		Id									= $M365Group.id;
		DisplayName							= $M365Group.displayName;
		description							= $M365Group.description;	
		createdDateTime						= $M365Group.createdDateTime;
		createdByAppId						= $M365Group.createdbyAppId;
		Mail								= $M365Group.Mail;
		MailNickname						= $M365Group.MailNickname;
		Owners								= $OwnerCount;
		Members								= $MemberCount;
		Guests								= $GuestCount;
		#ResourceBehaviorOptions				= $M365Group.resourceBehaviorOptions;
		#ResourceProvisioningOptions			= $M365Group.resourceProvisioningOptions;
		
		AllowOnlyMembersToPost 				= $M365Group.AllowOnlyMembersToPost;
		HideGroupInOutlook 					= $M365Group.HideGroupInOutlook;
		SubscribeNewGroupMembers 			= $M365Group.SubscribeNewGroupMembers;
		WelcomeEmailDisabled 				= $M365Group.WelcomeEmailDisabled;
		Visibility							= $M365Group.Visibility;
		#GroupTypes							= $M365Group.GroupTypes;
		Classification						= $TeamDetails.classification;
		IsMembershipLimitedToOwners			= $TeamDetails.isMembershipLimitedToOwners;
		ShowInTeamsSearchAndSuggestions		= $TeamDetails.DiscoverySettingsShowInTeamsSearchAndSuggestions;
		GuestsAllowCreateUpdateChannels 	= $TeamDetails.GuestSettings.allowCreateUpdateChannels;
		GuestsAllowDeleteChannels 			= $TeamDetails.GuestSettings.allowDeleteChannels;
		
		FunAllowGiphy						= $TeamDetails.FunSettings.allowGiphy;
		FunGiphyContentRating				= $TeamDetails.FunSettings.FunGiphyContentRating;
		FunAllowStickersAndMemes			= $TeamDetails.FunSettings.allowStickersAndMemes;
		FunAllowCustomMemes					= $TeamDetails.FunSettings.allowCustomMemes;

		AllowUserEditMessages				= $TeamDetails.MessagingSettings.allowUserEditMessages;
		AllowUserDeleteMessages				= $TeamDetails.MessagingSettings.allowUserDeleteMessages;
		AllowOwnerDeleteMessages			= $TeamDetails.MessagingSettings.allowOwnerDeleteMessages;
		AllowTeamMentions					= $TeamDetails.MessagingSettings.allowTeamMentions;

		AllowCreateUpdateChannels			= $TeamDetails.MemberSettings.allowCreateUpdateChannels; 
		AllowCreatePrivateChannels			= $TeamDetails.MemberSettings.allowCreatePrivateChannels; 
		AllowDeleteChannels					= $TeamDetails.MemberSettings.allowDeleteChannels; 
		AllowAddRemoveApps					= $TeamDetails.MemberSettings.allowAddRemoveApps; 
		AllowCreateUpdateRemoveTabs			= $TeamDetails.MemberSettings.allowCreateUpdateRemoveTabs;
		AllowCreateUpdateRemoveConnectors	= $TeamDetails.MemberSettings.allowCreateUpdateRemoveConnectors;
		
		WebURL								= $TeamDetails.webUrl
	}
	$M365Group_DB.Add($M365Group.id,$M365GroupData)
	$ReportTeamsList += $M365GroupData
}

Export-Report "Teams user membership info" -Report $ReportTeamsMembers -Path $OutputFileMembers
Export-Report "Teams user membership info (DB folder)" -Report $ReportTeamsMembers -Path $DBFileTeamsMembers
Export-Report "Teams list" -Report $ReportTeamsList -Path $OutputFileTmsList

Remove-Variable ReportTeamsList

##################################################################################################
# PER-TEAM TEAMS STATISTICS ######################################################################
##################################################################################################

############################################################################
#getOffice365GroupsActivityDetail
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
[array]$ReportStatsGrp = @()
$UriResource = "reports/getOffice365GroupsActivityDetail"
$UriReportPeriod = "D180"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -ReportPeriod $UriReportPeriod
$ActivityReport = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeCSV

foreach ($GroupStat in $ActivityReport) {
	$CurrentGroup = $null
	if ($M365Group_DB.ContainsKey($GroupStat."Group Id")) {
		$CurrentGroup =  $M365Group_DB.Item($GroupStat."Group Id")
	}
	$StatObject = [pscustomobject]@{
        TeamId                          	= $GroupStat."Group Id";
        IsArchived							= $CurrentGroup.isArchived;
        GroupName                 			= $CurrentGroup.displayName;
		CreatedDateTime						= $CurrentGroup.createdDateTime;
		LastActivityDate                	= $GroupStat."Last Activity Date";
		OwnerCount							= $CurrentGroup.owners;
		MemberCount							= $CurrentGroup.members;
		GuestCount							= $CurrentGroup.guests;
		GroupType                       	= $GroupStat."Group Type";
		ExchangeReceivedEmailCount      	= $GroupStat."Exchange Received Email Count";
        SharePointActiveFileCount       	= $GroupStat."SharePoint Active File Count";
        ExchangeMailboxTotalItemCount   	= $GroupStat."Exchange Mailbox Total Item Count";
        ExchangeMailboxStorageUsedBytes 	= $GroupStat."Exchange Mailbox Storage Used (Byte)";
        SharePointTotalFileCount        	= $GroupStat."SharePoint Total File Count";
        SharePointSiteStorageUsedBytes  	= $GroupStat."SharePoint Site Storage Used (Byte)";
        ReportPeriod                    	= $GroupStat."Report Period";
		WebURL								= $CurrentGroup.webUrl
    }
	$ReportStatsGrp += $StatObject
	$M365GroupStat_DB.Add($GroupStat."Group Id",$StatObject)
}
Export-Report "per-team Teams statistics (getOffice365GroupsActivityDetail)" -Report $ReportStatsGrp -Path $OutputFileStatGrp
Remove-Variable ReportStatsGrp
Remove-Variable ActivityReport

############################################################################
#getTeamsTeamActivityDetail-D7
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
[array]$ReportStatsTms = @()
$UriResource = "reports/getTeamsTeamActivityDetail"
$UriReportPeriod = "D7"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -ReportPeriod $UriReportPeriod
$ActivityReport = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeCSV

foreach ($GroupStat in $ActivityReport) {	
	$CurrentGroup = $null
	if ($M365Group_DB.ContainsKey($GroupStat."Team Id")) {
		$CurrentGroup =  $M365Group_DB.Item($GroupStat."Team Id")
	}
	$ReportStatsTms += [pscustomobject]@{
        TeamId                  = $GroupStat."Team Id";
        M365GroupIsArchived		= $CurrentGroup.isArchived;
		TeamIsDeleted			= $GroupStat."Is Deleted";
        InternalId				= $CurrentGroup.internalId;
		TeamName 				= $GroupStat."Team Name";
		CreatedDateTime			= $CurrentGroup.createdDateTime;
		LastActivityDate        = $GroupStat."Last Activity Date";
		OwnerCount				= $CurrentGroup.owners;
		MemberCount				= $CurrentGroup.members;
		GuestCount				= $CurrentGroup.guests;
		GroupType               = $GroupStat."Group Type";
		TeamType 				= $GroupStat."Team Type";
		ActiveUsers				= $GroupStat."Active Users";
		ActiveChannels 			= $GroupStat."Active Channels";
		ChannelMessages 		= $GroupStat."Channel Messages";
		Reactions 				= $GroupStat."Reactions";
		MeetingsOrganized 		= $GroupStat."Meetings Organized";
		PostMessages 			= $GroupStat."Post Messages";
		ReplyMessages 			= $GroupStat."Reply Messages";
		UrgentMessages 			= $GroupStat."Urgent Messages";
		Mentions 				= $GroupStat."Mentions";
		Guests 					= $GroupStat."Guests";
		ActiveSharedChannels 	= $GroupStat."Active Shared Channels";
		ActiveExternalUsers 	= $GroupStat."Active External Users";
		WebURL					= $CurrentGroup.webUrl
    }
}
Export-Report "per-team Teams statistics (getTeamsTeamActivityDetail) - D7" -Report $ReportStatsTms -Path $OutputFileStatTmsD7
Remove-Variable ReportStatsTms
Remove-Variable ActivityReport

############################################################################
#getTeamsTeamActivityDetail-D180
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
[array]$ReportStatsTms = @()
$UriResource = "reports/getTeamsTeamActivityDetail"
$UriReportPeriod = "D180"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -ReportPeriod $UriReportPeriod
$ActivityReport = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeCSV

foreach ($GroupStat in $ActivityReport) {
	Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
	$CurrentGroup = $null
	if ($M365Group_DB.ContainsKey($GroupStat."Team Id")) {
		$CurrentGroup =  $M365Group_DB.Item($GroupStat."Team Id")
	}
	$ReportStatsTms += [pscustomobject]@{
        TeamId                  = $GroupStat."Team Id";
        M365GroupIsArchived		= $CurrentGroup.isArchived;
		TeamIsDeleted			= $GroupStat."Is Deleted";
        InternalId				= $CurrentGroup.internalId;
		DisplayName             = $CurrentGroup.displayName;
		TeamName 				= $GroupStat."Team Name";
		CreatedDateTime			= $CurrentGroup.createdDateTime;
		LastActivityDate        = $GroupStat."Last Activity Date";
		Owners					= $CurrentGroup.owners;
		Members					= $CurrentGroup.members;
		GroupType               = $GroupStat."Group Type";
		TeamType 				= $GroupStat."Team Type";
		ActiveUsers				= $GroupStat."Active Users";
		ActiveChannels 			= $GroupStat."Active Channels";
		ChannelMessages 		= $GroupStat."Channel Messages";
		Reactions 				= $GroupStat."Reactions";
		MeetingsOrganized 		= $GroupStat."Meetings Organized";
		PostMessages 			= $GroupStat."Post Messages";
		ReplyMessages 			= $GroupStat."Reply Messages";
		UrgentMessages 			= $GroupStat."Urgent Messages";
		Mentions 				= $GroupStat."Mentions";
		Guests 					= $GroupStat."Guests";
		ActiveSharedChannels 	= $GroupStat."Active Shared Channels";
		ActiveExternalUsers 	= $GroupStat."Active External Users";
		WebURL					= $CurrentGroup.webUrl
    }
}
Export-Report "per-team Teams statistics (getTeamsTeamActivityDetail) - D180" -Report $ReportStatsTms -Path $OutputFileStatTmsD180
Remove-Variable ReportStatsTms
Remove-Variable ActivityReport

##################################################################################################
# PER-USER TEAMS STATISTICS ######################################################################
##################################################################################################
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
[array]$ReportStatsUsers = @()
$UriResource = "reports/getTeamsUserActivityUserDetail"
$UriReportPeriod = "D180"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -ReportPeriod $UriReportPeriod
$ActivityReport = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeCSV
foreach ($UserStat in $ActivityReport) {
	if ($AADUsers_DB.ContainsKey($UserStat."User Id")) {
		$CurrentUser =  $AADUsers_DB.Item($UserStat."User Id")
		$upn = $UserStat."User Principal Name"
		$maildomain = $null
		if ($CurrentUser.Mail) {
			$maildomain = $mail.Split("@")[1]
		}
		$MTRProLicensed = Test-ContainsMTRProLicense -LicenseString $UserStat."Assigned Products"

		$UserObject = [pscustomobject]@{
			UserId 					= $UserStat."User Id";
			UPN						= $upn;
			UPNDomain 				= $upn.Split("@")[1];
			DisplayName 			= $CurrentUser.DisplayName;
			UserType 				= $CurrentUser.UserType;
			ServiceAccount 			= Test-IsServiceAccount -Upn $upn;
			AccountEnabled 			= $CurrentUser.accountEnabled;
			Mail 					= $CurrentUser.Mail;
			Maildomain 				= $maildomain;
			onPremisesSyncEnabled 	= $CurrentUser.onPremisesSyncEnabled;
			CompanyName 			= $CurrentUser.CompanyName;
			LastActivityDate 		= $UserStat."Last Activity Date";
			IsLicensed 				= $UserStat."Is Licensed";
			OfficeLicensed 			= Test-ContainsOfficeLicense -LicenseString $UserStat."Assigned Products";
			MTRProLicensed 			= $MTRProLicensed
						
			TeamChatMessageCount 	= $UserStat."Team Chat Message Count";
			PrivateChatMessageCount = $UserStat."Private Chat Message Count";
			CallCount 				= $UserStat."Call Count";
			MeetingCount 			= $UserStat."Meeting Count";
			
			PostMessages 			= $UserStat."Post Messages";
			ReplyMessages 			= $UserStat."Reply Messages";
			UrgentMessages 			= $UserStat."Urgent Messages";
			
			MtgsOrganizedCount 					= $UserStat."Meetings Organized Count";
			AdHocMtgsOrganizedCount 			= $UserStat."Ad Hoc Meetings Organized Count";
			Scheduled1timeMtgsOrganizedCount 	= $UserStat."Scheduled One-time Meetings Organized Count";
			ScheduledRecurrMtgsOrganizedCount 	= $UserStat."Scheduled Recurring Meetings Organized Count";
			
			MtgsAttendedCount 					= $UserStat."Meetings Attended Count";
			AdHocMtgsAttendedCount 				= $UserStat."Ad Hoc Meetings Attended Count";
			Scheduled1timeMtgsAttendedCount 	= $UserStat."Scheduled One-time Meetings Attended Count";
			ScheduledRecurrMtgsAttendedCount 	= $UserStat."Scheduled Recurring Meetings Attended Count";
			
			AudioDuration 						= Convert-SecToHMS -Seconds $UserStat."Audio Duration In Seconds";
			AudioDurationSec 					= $UserStat."Audio Duration In Seconds";
			VideoDuration 						= Convert-SecToHMS -Seconds $UserStat."Video Duration In Seconds";
			VideoDurationSec 					= $UserStat."Video Duration In Seconds";
			ScreenShareDuration 				= Convert-SecToHMS -Seconds $UserStat."Screen Share Duration In Seconds";
			ScreenShareDurationSec 				= $UserStat."Screen Share Duration In Seconds";
			
			HasOtherAction 						= $UserStat."Has Other Action";
			AssignedProducts 					= $UserStat."Assigned Products"
		}
		$ReportStatsUsers += $UserObject
		if ($MTRProLicensed) {
			$ReportStatsUsersMTRPro += $UserObject
		}
	}
}

Export-Report "per-user Teams statistics" -Report $ReportStatsUsers -Path $OutputFileStatUsr
Export-Report "per-user Teams statistics (MTRPro)" -Report $ReportStatsUsersMTRPro -Path $OutputFileStatUsrMTR

Remove-Variable ActivityReport
Remove-Variable ReportStatsUsers
Remove-Variable ReportStatsUsersMTRPro

#######################################################################################################################
# TEAMS-CHANNELS-OWNERS 
#######################################################################################################################

$AADUsers_DB = Import-CSVtoHashDB -Path $DBFileUsersMemMin -KeyName "userPrincipalName"
$ReportTeamsChannelsOwners = Import-CSVtoArray -Path $DBFileTeamsChannelsOwners
Export-Report "Teams-Channels-Owners" -Report $ReportTeamsChannelsOwners -Path $OutputFileTmsChnl

#CEZ Teplarenska - report teams owned by CEZ Teplarenska users
$Teams = $ReportTeamsChannelsOwners | Where-Object { $_.Level -eq "Team"}
foreach ($Team in $Teams) {
	if ($Team.Owners) {
		$GroupStat = $Group = $null
		if ($M365GroupStat_DB.ContainsKey($Team.TeamId)) {
			$GroupStat = $M365GroupStat_DB.Item($Team.TeamId)
		}
		if ($M365Group_DB.ContainsKey($Team.TeamId)) {
			$Group = $M365Group_DB.Item($Team.TeamId)
		}
		$Owners = $Team.Owners.Split(";")
		foreach ($Owner in $Owners) {
			$OwnerRecord = $null
			$Owner = $Owner.Trim()
			if ($AADUsers_DB.ContainsKey($Owner)) {
				$OwnerRecord = $AADUsers_DB.Item($Owner)
				if ($OwnerRecord.Department -like "14_*") {
					$TeamObject = [pscustomobject]@{
						TeamId 			= $Team.TeamId;
						TeamName 		= $Team.TeamName;

						Description 	= $Group.Description;
						Mail 			= $Team.Mail;

						OwnerUPN			= $OwnerRecord.UserPrincipalName;
						OwnerName			= $OwnerRecord.DisplayName;
						OwnerEmail			= $OwnerRecord.Mail;
						CompanyName			= $OwnerRecord.CompanyName;
						Department			= $OwnerRecord.Department;

						FilesFolderUrl 	= $Team.FilesFolderUrl;
						CreatedDateTime = $Team.CreatedDateTime;
						Sensitivity 	= $Team.Sensitivity;
						Dynamic			= $Team.DynamicMembership;
						Visibility 		= $Team.TeamVisibility;
						isArchived 		= $Team.isArchived;
						AllowExtSenders = $Team.AllowExtSenders;
						AutoSubscribe 	= $Team.AutoSubscribe;
						#Owners 			= $Team.Owners;
						#TeamOwnerCount 	= $Team.TeamOwnerCount;
						TeamOwners 		= $Team.TeamOwners;
						TeamMembers 	= $Team.TeamMembers;
						TeamGuests 		= $Team.TeamGuests;
						LastActivityDate                	= $GroupStat.LastActivityDate;
						ExchangeReceivedEmailCount      	= $GroupStat.ExchangeReceivedEmailCount;
						SharePointActiveFileCount       	= $GroupStat.SharePointActiveFileCount;
						ExchangeMailboxTotalItemCount   	= $GroupStat.ExchangeMailboxTotalItemCount;
						ExchangeMailboxStorageUsedBytes 	= $GroupStat.ExchangeMailboxStorageUsedBytes;
						SharePointTotalFileCount        	= $GroupStat.SharePointTotalFileCount;
						SharePointSiteStorageUsedBytes  	= $GroupStat.SharePointSiteStorageUsedBytes;
						StatDataReportPeriod               	= $GroupStat.ReportPeriod
					}
					$ReportTeamsCEZTpl += $TeamObject
				}
			}
		}
	}
}

$CEZTplTeams = $ReportTeamsCEZTpl.TeamId | Sort-Object -Unique
$ReportTeamsCEZTplMem = $ReportTeamsMembers | Where-Object {$_.TeamId -in $CEZTplTeams}

Export-Report "Teams-CEZ-Teplarenska" -Report $ReportTeamsCEZTpl -Path $OutputFileTmsCEZTpl
Export-Report "Teams-CEZ-Teplarenska-members" -Report $ReportTeamsCEZTplMem -Path $OutputFileTmsCEZTplMem

Remove-Variable ReportTeamsChannelsOwners
Remove-Variable ReportTeamsCEZTpl
Remove-Variable ReportTeamsCEZTplMem

##################################################################################################

. $ScriptPath\include-Script-EndLog-Generic.ps1
