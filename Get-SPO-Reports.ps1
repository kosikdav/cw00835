#######################################################################################################################
# Get-SPO-Reports
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1

#######################################################################################################################

$LogFolder			= "exports"
$LogFilePrefix		= "spo-reports"

$OutputFolder		= "spo\reports"
$OutputFilePrefix	= "sharepoint"

$OutputFileSuffixSPOSites		= "spo-sites"
$OutputFileSuffixSPOSitesOwners	= "spo-sites-owners"
$OutputFileSuffixODfBSites		= "odfb-sites"
$OutputFileSuffixStatSPOSte		= "spo-stats-per-site"
$OutputFileSuffixStatSPOUsr		= "spo-stats-per-user"
$OutputFileSuffixStatODfBSte 	= "odfb-stats-per-site"
$OutputFileSuffixStatODfBUsr 	= "odfb-stats-per-user"

[array]$SiteOwners_IgnoredSites = @(
	"https://cezdata.sharepoint.com/search",
	"https://cezdata.sharepoint.com",
	"https://cezdata.sharepoint.com/sites/app_catalog"
)

[array]$SiteOwners_IgnoredUrlPatterns = @(
	"*/portals/*"
)

[array]$SiteOwners_IgnoredTemplatePatterns = @(
	"TEAMCHANNEL*",
	#"GROUP#0",
	"RedirectSite*"
)

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputFileSPOSites 	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixSPOSites -Ext "csv"
$OutputFileSPOSiteOwners 	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixSPOSitesOwners -Ext "csv"
$OutputFileODfBSites 	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixODfBSites -Ext "csv"
$OutputFileStatSPOSte 	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixStatSPOSte -Ext "csv"
$OutputFileStatSPOUsr 	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixStatSPOUsr -Ext "csv"
$OutputFileStatODfBSte 	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixStatODfBSte -Ext "csv"
$OutputFileStatODfBUsr 	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixStatODfBUsr -Ext "csv"

[array]$ReportSPOSites	= @()
[array]$ReportSPOSiteOwners = @()
[array]$ReportODfBSites	= @()
[array]$ReportStatsSPOSites = @()
[array]$ReportStatsSPOUsers = @()
[array]$ReportStatsODfBSites = @()
[array]$ReportStatsODfBUsers = @()
[hashtable]$GraphSites_DB = @{}

#######################################################################################################################

. $IncFile_StdLogStartBlock

write-log "$($OutputFileSPOSites)"
write-log "$($OutputFileSPOSiteOwners)"
write-log "$($OutputFileODfBSites)"
write-log "$($OutputFileStatSPOSte)"
write-log "$($OutputFileStatSPOUsr)"
write-log "$($OutputFileStatODfBSte)"
write-log "$($OutputFileStatODfBUsr)"

############################################################
# Get SPO site owners report
############################################################
Connect-SPOServicePnP -AppRegName $AppReg_SPO_REPORT_PnP 
$SPOSiteCollections = Get-PnPTenantSite

:main foreach ($SiteCollection in $SPOSiteCollections) {
	if ($SiteOwners_IgnoredSites -Contains $SiteCollection.Url.TrimEnd("/")) {
		continue
	}
	foreach ($Pattern in $SiteOwners_IgnoredUrlPatterns) {
		if ($SiteCollection.Url -like $Pattern) {
			continue main
		}
	}
	foreach ($Pattern in $SiteOwners_IgnoredTemplatePatterns) {
		if ($SiteCollection.Template -like $Pattern) {
			continue main
		}
	}

	Connect-SPOServicePnP -AppRegName $AppReg_SPO_REPORT_PnP -Url $SiteCollection.Url -Silent:$true
	
	$Admins = $Owners = [string]::Empty
	$SiteAdmins = Get-PnPSiteCollectionAdmin
	Try {
		$SiteOwners  = Get-PnPGroupMember -Group (Get-PnPGroup -AssociatedOwnerGroup)
	}
	Catch {
		$SiteOwners = [array]::Empty
	}
	foreach ($Admin in $SiteAdmins) {
		if ($Admin.LoginName -match "[^|]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$") {
			$Admins = $Admins + $Matches[0] + ","
		}
	}
	$Admins = $Admins.TrimEnd(",")

	foreach ($Owner in $SiteOwners) {
		if ($Owner.LoginName -match "[^|]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$") {
    		$Owners = $Owners + $Matches[0] + ","
		}
	}
	$Owners = $Owners.TrimEnd(",")

	$ReportSPOSiteOwners += [pscustomobject]@{
		#id = $SiteCollection.Id
		SiteTitle = $SiteCollection.Title
		URL = $SiteCollection.Url
		Owners = $Owners
		Admins = $Admins
		Template = $SiteCollection.Template
	} 
}

Export-Report -Text "SPO sites owners report" -Report $ReportSPOSiteOwners -SortProperty "UserPrincipalName" -Path $OutputFileSPOSiteOwners -Delimiter ";"

############################################################
# SPO site report
############################################################

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "sites"
$UriSelect = "id,createdDateTime,webUrl"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect -Top 999
$GraphSites = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "Getting SPO sites" -ProgressDots
foreach ($Site in $GraphSites) {
	$SiteObject = [pscustomobject]@{
		id = $Site.id;
		createdDateTime = $Site.createdDateTime
	}
	$GraphSites_DB.Add($Site.webUrl, $SiteObject)
}

$AADUsers_DB = Import-CSVtoHashDB -Path $DBFileUsersAllMin -KeyName "userPrincipalName"

Connect-SPOServicePnP -AppRegName $AppReg_SPO_REPORT_PnP -TTL 120
Write-Host "Getting all tenant SPO sites via PnP..." -NoNewline
$SPOSites = Get-PnPTenantSite -Detailed -IncludeOneDriveSites
Write-Host "done ($($SPOSites.Count))"
foreach ($Site in $SPOSites) {
	$PersonalSite = $false
	$SiteId = $createdDateTime = $ownerId = $ownerMail = $ownerName = $ownerCompany = $ownerDepartment = [string]::Empty
	if ($Site.Url.StartsWith($RootODfBURL)) {
		$PersonalSite = $true 
		if ($AADUsers_DB.ContainsKey($Site.Owner)) {
			$CurrentOwner = $AADUsers_DB.Item($Site.Owner)
			if ($CurrentOwner.Mail) {
				$ownerId = $CurrentOwner.id
				$ownerMail = $CurrentOwner.mail
				$ownerName = $CurrentOwner.displayName
				$ownerCompany = $CurrentOwner.companyName
				$ownerDepartment = $CurrentOwner.department
			}
		}
	}
	if ($GraphSites_DB.ContainsKey($Site.Url)) {
		$SiteId = $GraphSites_DB.Item($Site.Url).id
		$createdDateTime = $GraphSites_DB.Item($Site.Url).createdDateTime
	}
	$SiteObject = [pscustomobject]@{
		Title = $Site.Title;
		Url = $Site.Url;
		SiteId = $SiteId;
		Template = $Site.Template;
		PersonalSite = $PersonalSite;
		GroupId = $Site.GroupId;
		HubSiteId = $Site.HubSiteId;
		Status = $Site.Status;
		CreatedDateTime = $createdDateTime;
		LastModifiedDateTime = $Site.LastContentModifiedDate;		
		AllowDownloadingNonWebViewableFiles = $Site.AllowDownloadingNonWebViewableFiles;
		AllowEditing = $Site.AllowEditing;
		AllowSelfServiceUpgrade = $Site.AllowSelfServiceUpgrade;
		AnonymousLinkExpirationInDays = $Site.AnonymousLinkExpirationInDays;
		BlockDownloadLinksFileType = $Site.BlockDownloadLinksFileType;
		CommentsOnSitePagesDisabled = $Site.CommentsOnSitePagesDisabled;
		CompatibilityLevel = $Site.CompatibilityLevel;
		ConditionalAccessPolicy = $Site.ConditionalAccessPolicy;
		DefaultLinkPermission = $Site.DefaultLinkPermission;
		DefaultLinkToExistingAccess = $Site.DefaultLinkToExistingAccess;
		DefaultSharingLinkType = $Site.DefaultSharingLinkType;
		DenyAddAndCustomizePages = $Site.DenyAddAndCustomizePages;
		Description = $Site.Description;
		DisableAppViews = $Site.DisableAppViews;
		DisableCompanyWideSharingLinks = $Site.DisableCompanyWideSharingLinks;
		DisableFlows = $Site.DisableFlows;
		DisableSharingForNonOwnersStatus = $Site.DisableSharingForNonOwnersStatus;
		ExternalUserExpirationInDays = $Site.ExternalUserExpirationInDays;
		InformationSegment = $Site.InformationSegment;
		IsHubSite = $Site.IsHubSite;
		LimitedAccessFileType = $Site.LimitedAccessFileType;
		LocaleId = $Site.LocaleId;
		LockIssue = $Site.LockIssue;
		LockState = $Site.LockState;
		OverrideTenantAnonymousLinkExpirationPolicy = $Site.OverrideTenantAnonymousLinkExpirationPolicy;
		OverrideTenantExternalUserExpirationPolicy = $Site.OverrideTenantExternalUserExpirationPolicy;
		Owner = $Site.Owner;
		OwnerId = $ownerId;
		OwnerEmail = $ownerMail;
		OwnerName = $ownerName;
		OwnerCompany = $ownerCompany;
		OwnerDepartment = $ownerDepartment;
		ProtectionLevelName = $Site.ProtectionLevelName;
		PWAEnabled = $Site.PWAEnabled;
		RelatedGroupId = $Site.RelatedGroupId;
		ResourceQuota = $Site.ResourceQuota;
		ResourceQuotaWarningLevel = $Site.ResourceQuotaWarningLevel;
		ResourceUsageAverage = $Site.ResourceUsageAverage;
		ResourceUsageCurrent = $Site.ResourceUsageCurrent;
		RestrictedToGeo = $Site.RestrictedToGeo;
		SandboxedCodeActivationCapability = $Site.SandboxedCodeActivationCapability;
		SensitivityLabel = $Site.SensitivityLabel;
		SharingAllowedDomainList = $Site.SharingAllowedDomainList;
		SharingBlockedDomainList = $Site.SharingBlockedDomainList;
		SharingCapability = $Site.SharingCapability;
		SharingDomainRestrictionMode = $Site.SharingDomainRestrictionMode;
		ShowPeoplePickerSuggestionsForGuestUsers = $Site.ShowPeoplePickerSuggestionsForGuestUsers;
		SiteDefinedSharingCapability = $Site.SiteDefinedSharingCapability;
		SocialBarOnSitePagesDisabled = $Site.SocialBarOnSitePagesDisabled;
		StorageQuota = $Site.StorageQuota;
		StorageQuotaType = $Site.StorageQuotaType;
		StorageQuotaWarningLevel = $Site.StorageQuotaWarningLevel;
		StorageUsageCurrent = $Site.StorageUsageCurrent;
		WebsCount = $Site.WebsCount;
	}
	if ($PersonalSite) {
		$ReportODfBSites += $SiteObject
	}
	else {
		$ReportSPOSites += $SiteObject
	}
}
Export-Report -Text "SPO sites report" -Report $ReportSPOSites -SortProperty "UserPrincipalName" -Path $OutputFileSPOSites
Export-Report -Text "SPO sites (DB folder)" -Report $ReportSPOSites -SortProperty "UserPrincipalName" -Path $DBFileSPOSites
Export-Report -Text "ODfB sites report" -Report $ReportODfBSites -SortProperty "UserPrincipalName" -Path $OutputFileODfBSites
Export-Report -Text "ODfB sites (DB folder)" -Report $ReportODfBSites -SortProperty "UserPrincipalName" -Path $DBFileODfBSites

############################################################
# Get per-site SPO statistics
############################################################
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "reports/getSharePointSiteUsageDetail"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -ReportPeriod "D180"
Write-Host "Getting per-site SPO statistics..." -NoNewline
$SharePointSiteUsageDetailReport = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeCSV -ProgressDots

foreach ($SiteStat in $SharePointSiteUsageDetailReport) {
	$ReportStatsSPOSites += [pscustomobject]@{
        SiteId 				= $SiteStat."Site Id";
		SiteURL 			= $SiteStat."Site URL";
		OwnerPrincipalName 	= $SiteStat."Owner Principal Name";
		OwnerDisplayName 	= $SiteStat."Owner Display Name";
		IsDeleted 			= $SiteStat."Is Deleted";
		LastActivityDate 	= $SiteStat."Last Activity Date";
		FileCount 			= $SiteStat."File Count";
		ActiveFileCount 	= $SiteStat."Active File Count";
		PageViewCount 		= $SiteStat."Page View Count";
		VisitedPageCount 	= $SiteStat."Visited Page Count";
		StorageUsed 		= $SiteStat."Storage Used (Byte)";
		StorageAllocated 	= $SiteStat."Storage Allocated (Byte)";
		RootWebTemplate 	= $SiteStat."Root Web Template"
    }
}
Export-Report -Text "per-site SPO statistic" -Report $ReportStatsSPOSites -SortProperty "UserPrincipalName" -Path $OutputFileStatSPOSte

############################################################
# Get per-user SPO statistics
############################################################
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "reports/getSharePointActivityUserDetail"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -ReportPeriod "D180"
Write-Host "Getting per-user SPO statistics..." -NoNewline
$SharePointActivityUserDetailReport = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeCSV -ProgressDots

foreach ($UserStat in $SharePointActivityUserDetailReport) {
	$UPN = $UserStat."User Principal Name"
	if ($AADUsers_DB.ContainsKey($UPN)) {
		$CurrentUser = $AADUsers_DB.Item($upn)
		$mail = $CurrentUser.Mail
		if ($null -ne $mail) {
			$maildomain = $mail.Split("@")[1] 
		} 
		else { 
			$maildomain = $null 
		}
		$ReportStatsSPOUsers += [pscustomobject]@{
			UPN							= $upn;
			UPNDomain 					= $upn.Split("@")[1];
			DisplayName 				= $CurrentUser.DisplayName;
			UserType 					= $CurrentUser.UserType;
			ServiceAccount 				= Test-IsServiceAccount -Upn $upn;
			AccountEnabled 				= $CurrentUser.accountEnabled;
			Mail 						= $mail;
			Maildomain 					= $maildomain;
			onPremisesSyncEnabled 		= $CurrentUser.onPremisesSyncEnabled;
			CompanyName 				= $CurrentUser.CompanyName;
			LastActivityDate 			= $UserStat."Last Activity Date";
			IsOfficeLicensed 			= Test-ContainsOfficeLicense -LicenseString $UserStat."Assigned Products";
			ViewedOrEditedFileCount 	= $UserStat."Viewed Or Edited File Count";
			SyncedFileCount 			= $UserStat."Synced File Count";
			SharedInternallyFileCount	= $UserStat."Shared Internally File Count";
			SharedExternallyFileCount	= $UserStat."Shared Externally File Count";
			VisitedPageCount			= $UserStat."Visited Page Count";
			AssignedProducts 			= $UserStat."Assigned Products"
		}
	}
}
Export-Report -Text "per-user SPO statistic" -Report $ReportStatsSPOUsers -SortProperty "UserPrincipalName" -Path $OutputFileStatSPOUsr

############################################################
# Get per-site ODfB statistics
############################################################
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "reports/getOneDriveUsageAccountDetail"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -ReportPeriod "D180"
Write-Host "Getting per-site ODfB statistics..." -NoNewline
$OneDriveUsageAccountDetailReport = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeCSV -ProgressDots
foreach ($SiteStat in $OneDriveUsageAccountDetailReport) {
	$ReportStatsODfBSites += [pscustomobject]@{
		SiteURL 			= $SiteStat."Site URL";
		OwnerPrincipalName 	= $SiteStat."Owner Principal Name";
		OwnerDisplayName 	= $SiteStat."Owner Display Name";
		IsDeleted 			= $SiteStat."Is Deleted";
		LastActivityDate 	= $SiteStat."Last Activity Date";
		FileCount 			= $SiteStat."File Count";
		ActiveFileCount 	= $SiteStat."Active File Count";
		StorageUsed 		= $SiteStat."Storage Used (Byte)";
		StorageAllocated 	= $SiteStat."Storage Allocated (Byte)";
    }
}
Complete-ProgressBarMain
Export-Report -Text "per-site ODfB statistic" -Report $ReportStatsODfBSites -SortProperty "UserPrincipalName" -Path $OutputFileStatODfBSte

############################################################
# Get per-user ODfB statistics
############################################################
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "reports/getOneDriveActivityUserDetail"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -ReportPeriod "D180"
Write-Host "Getting per-user ODfB statistics..." -NoNewline
$OneDriveActivityUserDetailReport = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER]AccessToken -ContentType $ContentTypeCSV -ProgressDots
Initialize-ProgressBarMain -Activity "Building per-site ODfB statistics" -Total $OneDriveActivityUserDetailReport.Count
foreach ($UserStat in $OneDriveActivityUserDetailReport) {
	Update-ProgressBarMain
	$UPN = $UserStat."User Principal Name"
	if ($AADUsers_DB.ContainsKey($UPN)) {
		$CurrentUser = $AADUsers_DB.Item($UPN)
		$mail = $CurrentUser.Mail
		if ($null -ne $mail) { $maildomain = $mail.Split("@")[1] } else { $maildomain = $null }
		$ReportStatsODfBUsers += [pscustomobject]@{
			UPN							= $upn;
			UPNDomain 					= $upn.Split("@")[1];
			DisplayName 				= $CurrentUser.DisplayName;
			UserType 					= $CurrentUser.UserType;
			ServiceAccount 				= Test-IsServiceAccount -Upn $upn;
			AccountEnabled 				= $CurrentUser.accountEnabled;
			Mail 						= $mail;
			Maildomain 					= $maildomain;
			onPremisesSyncEnabled 		= $CurrentUser.onPremisesSyncEnabled;
			CompanyName 				= $CurrentUser.CompanyName;
			LastActivityDate 			= $UserStat."Last Activity Date";
			IsOfficeLicensed 			= Test-ContainsOfficeLicense -LicenseString $UserStat."Assigned Products";
			ViewedOrEditedFileCount 	= $UserStat."Viewed Or Edited File Count";
			SyncedFileCount 			= $UserStat."Synced File Count";
			SharedInternallyFileCount	= $UserStat."Shared Internally File Count";
			SharedExternallyFileCount	= $UserStat."Shared Externally File Count";
			AssignedProducts 			= $UserStat."Assigned Products"
		}
	}
}
Export-Report -Text "per-user ODfB statistic" -Report $ReportStatsODfBUsers -SortProperty "UserPrincipalName" -Path $OutputFileStatODfBUsr

#######################################################################################################################

. $IncFile_StdLogEndBlock
