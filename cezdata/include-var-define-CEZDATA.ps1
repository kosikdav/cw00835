###############################
# include-var-define-CEZDATA
###############################
$TenantId           = "b233f9e1-5599-4693-9cef-38858fe25406"
$TenantName         = "cezdata.onmicrosoft.com"
$TenantShortName    = "CEZDATA"
$RootSPOURL         = "https://cezdata.sharepoint.com/sites"
$RootODfBURL        = "https://cezdata-my.sharepoint.com/personal"
$PnPURL             = "https://cezdata.sharepoint.com"
$PnPTenant          = "cezdata.onmicrosoft.com"
$TenantAdminURL     = "https://cezdata-admin.sharepoint.com"
$GraphV1            = "https://graph.microsoft.com/v1.0"
$GraphBeta          = "https://graph.microsoft.com/beta"
$GuestUPNSuffix     = "#ext#@cezdata.onmicrosoft.com"
$DefaultUPNSuffix     = "@cezdata.onmicrosoft.com"
$DelGuestUPNSuffix  = "#ext#_cezdata_onmicrosoft_com"
$DelUPNPrefixLength = 32
$SearchUALIndexErrorMaxTotal = 100
$SearchUALIndexErrorMaxCycle = 20
$LitigationHoldDuration = 1825
$LHDuration = $LitigationHoldDuration
$MaxReceiveSize = "150MB"
$MaxSendSize 	= "150MB"
$MSGraphResourceId  = "5796a0fc-bded-497c-ad31-4b35e292bc88"
$InactivityLimitGuests = 190
$InactivityLimitPendingInvites = 30

$EXOMbxReportPermissions = $false
$EXOMbxReportTNR = $true

$AADUserReportTNR = $true
$AADUserReportTNR_attr_label = "DepartmentTNR"
$AADUserReportTNR_ext_name = "ext_msDScloudExtensionAttribute1"
$AADUserReportAuthMobile_attr_label = "ext_cEZIntuneMFAAuthMobile"
$AADUserReportAuthMobile_ext_name = "ext_cEZIntuneMFAAuthMobile"


$AADUserReportGroupMemberCount = $true

$string_divider = "-------------------------------------------------------------------------------"

$ps5exe = "c:\windows\system32\windowspowershell\v1.0\powershell.exe"
$ps7exe = "c:\windows\system32\windowspowershell\v1.0\powershell.exe"
$psexe  = $ps5exe

$root_log_folder    = "d:\logs\cezdata"
$root_output_folder = "d:\exports\cezdata"
$scriptsFolder      = "d:\scripts"
$incFolder          = "d:\scripts\cezdata"
$DBFolderName       = "db"
$AADCredsFolder     = "c:\inetpub\sites\aadcreds"
$M365LicFolder      = "c:\inetpub\sites\m365lic"

$aad_grp_mgmt_cred      = "c:\cred\qp_aad_grp_mgmt\qp_aad_grp_mgmt.cred"
$aadauthmobmgmt_cred    = "c:\cred\qp_aadauthmobmgmt\qp_aadauthmobmgmt.cred"

$IncFile_Var_Init           = [System.IO.Path]::Combine($scriptsFolder,"include-var-init.ps1")
$IncFile_Functions_Common   = [System.IO.Path]::Combine($scriptsFolder,"include-functions-common.ps1")
$IncFile_Functions_Audit    = [System.IO.Path]::Combine($scriptsFolder,"include-functions-audit.ps1")

$IncFile_StdLogBeginBlock   = [System.IO.Path]::Combine($scriptsFolder,"include-Script-StdLogStartBlock.ps1")
$IncFile_StdLogStartBlock   = $IncFile_StdLogBeginBlock
$IncFile_StdLogEndBlock     = [System.IO.Path]::Combine($scriptsFolder,"include-Script-StdLogEndBlock.ps1")

$AppReg_LOG_READER          = "CEZDATA_LOG_READER"
$AppReg_LOG_READER_MIN      = "CEZDATA_LOG_READER_MIN"
$AppReg_TMS_MGMT            = "CEZDATA_TEAMS_MGMT"
#$AppReg_MFA_MGMT           = "CEZDATA_AAD_MFA_config"
$AppReg_USR_MGMT            = "CEZDATA_AAD_USR_MGMT"
$AppReg_APP_MGMT            = "CEZDATA_AAD_APP_MGMT"
$AppReg_EXO_MGMT            = "CEZDATA_EXO_MGMT"
$AppReg_SPO_REPORT_PnP      = "CEZDATA_SPO_REPORTS_PnP"
$AppReg_SPO_MGMT            = "CEZDATA_SPO_MGMT"
$AppReg_ROLE_MGMT           = "CEZDATA_ROLE_MGMT"
$AppReg_ICTS_Fabric_Mgmt    = "CEZDATA_ICTS_Fabric_Mgmt"
$AppReg_PowerBI_REST_API_Reader = "CEZDATA_PowerBI_REST_API_Reader"

$IncFile_AppReg_LOG_READER          = [System.IO.Path]::Combine($incFolder,"include-appreg-" + $AppReg_LOG_READER + ".ps1")
$IncFile_AppReg_TMS_MGMT            = [System.IO.Path]::Combine($incFolder,"include-appreg-" + $AppReg_TMS_MGMT + ".ps1")
$IncFile_AppReg_MFA_MGMT            = [System.IO.Path]::Combine($incFolder,"include-appreg-" + $AppReg_MFA_MGMT + ".ps1")
$IncFile_AppReg_USR_MGMT            = [System.IO.Path]::Combine($incFolder,"include-appreg-" + $AppReg_USR_MGMT + ".ps1")
$IncFile_AppReg_EXO_MGMT            = [System.IO.Path]::Combine($incFolder,"include-appreg-" + $AppReg_EXO_MGMT + ".ps1")
$IncFile_AppReg_SPO_REPORT_PnP      = [System.IO.Path]::Combine($incFolder,"include-appreg-" + $AppReg_SPO_REPORT_PnP + ".ps1")
$IncFile_AppReg_ICTS_Fabric_Mgmt    = [System.IO.Path]::Combine($incFolder,"include-appreg-" + $AppReg_ICTS_Fabric_Mgmt + ".ps1")
$IncFile_AIP_labels                 = [System.IO.Path]::Combine($incFolder,"include-var-define-CEZDATA-AIP-labels.ps1")

$ConfigFile_AADGroupMirror             = [System.IO.Path]::Combine($incFolder,"aad-group-mirror-config.csv")

$RLF = $root_log_folder
$ROF = $root_output_folder

$DBFileAADAppList           = [System.IO.Path]::Combine($ROF,$DBFolderName,"aad-app.csv")
$DBFileAADSPList            = [System.IO.Path]::Combine($ROF,$DBFolderName,"aad-sp.csv")
$DBFileLicensingInfoSKUs    = [System.IO.Path]::Combine($ROF,$DBFolderName,"m365-licensing-SKUs.csv")
$DBFileLicensingInfoPlans   = [System.IO.Path]::Combine($ROF,$DBFolderName,"m365-licensing-plans.csv")
$DBFileLicensingAADP1       = [System.IO.Path]::Combine($ROF,$DBFolderName,"m365-licensing-aadp1.csv")
$DBFileLicensingAADP2       = [System.IO.Path]::Combine($ROF,$DBFolderName,"m365-licensing-aadp2.csv")
$DBFileLicensingEXO         = [System.IO.Path]::Combine($ROF,$DBFolderName,"m365-licensing-exo.csv")
$DBFileLicensingSPO         = [System.IO.Path]::Combine($ROF,$DBFolderName,"m365-licensing-spo.csv")

$DBFileAADResourceActions   = [System.IO.Path]::Combine($ROF,$DBFolderName,"aad-resource-actions.csv")
$DBFileAADOAuthScopes       = [System.IO.Path]::Combine($ROF,$DBFolderName,"aad-oauth-scopes.csv")
$DBFileAADAppRoles          = [System.IO.Path]::Combine($ROF,$DBFolderName,"aad-app-roles.csv")
$DBFileAADResourcePerms     = [System.IO.Path]::Combine($ROF,$DBFolderName,"aad-resource-perms.csv")
$DBFileAADAdmRoles          = [System.IO.Path]::Combine($ROF,$DBFolderName,"aad-admin-roles.csv")
$DBFileAADPartnerTenants    = [System.IO.Path]::Combine($ROF,$DBFolderName,"aad-partner-tenants.csv")
$DBFileExtAADTenants        = [System.IO.Path]::Combine($ROF,$DBFolderName,"aad-ext-tenants.csv")
$DBFileTenantDomains        = [System.IO.Path]::Combine($ROF,$DBFolderName,"m365-tenant-domains.csv")
$DBFileGuestsStd            = [System.IO.Path]::Combine($ROF,$DBFolderName,"guests-std.csv")
$DBFileGuests               = $DBFileGuestsStd
$DBFileGuestsName           = [System.IO.Path]::Combine($ROF,$DBFolderName,"guests-name.csv")
$DBFileGuestsSIA            = [System.IO.Path]::Combine($ROF,$DBFolderName,"guests-SIA.csv")
$DBFileUsersAllName         = [System.IO.Path]::Combine($ROF,$DBFolderName,"users-all-name.csv")

$DBFileUsers                = [System.IO.Path]::Combine($ROF,$DBFolderName,"users-aad.csv")
$DBFileUsersAllMin          = [System.IO.Path]::Combine($ROF,$DBFolderName,"users-all-min.csv")
$DBFileUsersAllStd          = [System.IO.Path]::Combine($ROF,$DBFolderName,"users-all-std.csv")
$DBFileUsersAllSIA          = [System.IO.Path]::Combine($ROF,$DBFolderName,"users-all-SIA.csv")
$DBFileUsersMemName         = [System.IO.Path]::Combine($ROF,$DBFolderName,"users-mem-name.csv")
$DBFileUsersMemMin          = [System.IO.Path]::Combine($ROF,$DBFolderName,"users-mem-min.csv")
$DBFileUsersMemStd          = [System.IO.Path]::Combine($ROF,$DBFolderName,"users-mem-std.csv")
$DBFileUsersMemSIA          = [System.IO.Path]::Combine($ROF,$DBFolderName,"users-mem-SIA.csv")
$DBFileUsersMemLic          = [System.IO.Path]::Combine($ROF,$DBFolderName,"users-mem-lic.csv")
$DBFileGroups               = [System.IO.Path]::Combine($ROF,$DBFolderName,"groups.csv")
$DBFileGroupsMembers        = [System.IO.Path]::Combine($ROF,$DBFolderName,"groups-members.csv")
$DBFileGroupsAllMin         = [System.IO.Path]::Combine($ROF,$DBFolderName,"groups-all-min.csv")
$DBFileGroupsM365           = [System.IO.Path]::Combine($ROF,$DBFolderName,"groups-m365.csv")
$DBFileGroupsO365           = $DBFileGroupsM365
$DBFileO365Groups           = $DBFileGroupsM365
$DBFileTeams                = [System.IO.Path]::Combine($ROF,$DBFolderName,"teams.csv")
$DBFileTeamsMembers         = [System.IO.Path]::Combine($ROF,$DBFolderName,"teams-members.csv")
$DBFileTeamsChannelsOwners  = [System.IO.Path]::Combine($ROF,$DBFolderName,"teams-channels-owners.csv")
$DBFileSPOSites             = [System.IO.Path]::Combine($ROF,$DBFolderName,"spo-sites.csv")
$DBFileSPOSitesGraph        = [System.IO.Path]::Combine($ROF,$DBFolderName,"spo-sites-graph.csv")
$DBFileODfBSites            = [System.IO.Path]::Combine($ROF,$DBFolderName,"odfb-sites.csv")
$DBFilePwrEnvironments      = [System.IO.Path]::Combine($ROF,$DBFolderName,"pwrplat-environments.csv")

$DBFileLicensingPlans       = "d:\data\m365-licensing\plans.csv"
$DBFileLicensingSKUs        = "d:\data\m365-licensing\skus.csv"
$DBFileAADCreds             = "d:\data\entra-app-creds\app-creds.xml"
$DBFileAADPermCustom        = "d:\data\entra-app-permissions\app-permissions-custom.csv"
$DBFileAADPermissions       = "d:\data\entra-app-permissions\app-permissions.csv"
$DBFileTMS_CloudPoradna     = "d:\data\tms_icts_cloud_poradna\poradna-members-db.xml"
$DBFileMGMTAPI_Audit_SPO    = "d:\data\o365-mgmt-api\audit-spo-blobs-db.xml"
$DBFileEXOMboxMgmt          = "d:\data\exo-mailbox-mgmt\mailbox-mgmt-db.xml"
$DBFileEXOMobileDeviceMgmt  = "d:\data\exo-mobile-device-mgmt\mobile-device-mgmt-db.xml"
$DBFileMFAMgmt              = "d:\data\entra-mfa-mgmt\mfa-mgmt-db.xml"

$outputFileReportSuffix     = "report"
$outputFileAuditSuffix 	    = "audit"

$ContentTypeJSON            = "application/json"
$ContentTypeCSV             = "text/csv"
$ConsistencyLevelEventual   = "eventual"

$CCIgnoreCase = "CurrentCultureIgnoreCase"

$TypeGroup = "#microsoft.graph.group"
$TypeUser = "#microsoft.graph.user"

$UriAnd = "+and+"

$AuditAdminActionsAll = @("ApplyRecord","Copy","Create","FolderBind","HardDelete","MailItemsAccessed","Move","MoveToDeletedItems","RecordDelete","Send","SendAs","SendOnBehalf","SoftDelete","Update","UpdateCalendarDelegation","UpdateFolderPermissions","UpdateInboxRules")
$AuditDelegateActionsAll = @("ApplyRecord","Create","FolderBind","HardDelete","MailItemsAccessed","Move","MoveToDeletedItems","RecordDelete","SendAs","SendOnBehalf","SoftDelete","Update","UpdateFolderPermissions","UpdateInboxRules")
$AuditOwnerActionsAll = @("ApplyRecord","Create","HardDelete","MailboxLogin","MailItemsAccessed","Move","MoveToDeletedItems","RecordDelete","SearchQueryInitiated","Send","SoftDelete","Update","UpdateCalendarDelegation","UpdateFolderPermissions","UpdateInboxRules")

$T2TTenant_DB = @{
    "87bd4c77-46eb-4cdf-9c33-a5597730faa6" = "CAPEXUS"
    "65afc824-f110-42ab-8a83-247c89d0eed8" = "ENESA"
    "1dabd27c-3764-4c3e-9072-2370ef0ba2cc" = "CEZ ENERGO"
    "4fdc467c-11a1-4d87-b1dd-25dc97b3ab6e" = "HORMEN"
    "c9d29adc-dcc2-4c1d-adfe-9cc2a31ab123" = "CEZ ESL"
    "071feddb-1946-4c52-a3c8-aa8ed2c44d14" = "AZ KLIMA"
    "bde9495d-be92-4384-868b-07ae5b5bd834" = "EP ROZNOV"
    "1f11ae71-9614-4218-9c09-4ceb975c8bec" = "KART"
    "3687fd79-edff-4560-9dda-317079330262" = "AIRPLUS"
    "d3a5825a-e344-414e-9422-6b5883d4cafc" = "DOMAT"
    "a6150323-e050-490e-b983-07cb945d9add" = "CEZNET"
    "56b31968-ca9e-4cc3-9257-477c3699b885" = "UJV"
    "c6807474-993d-4891-994b-79c685818b9c" = "ELENG"
    "12437861-f55d-4e74-8b78-47996c60686a" = "SDAS"
}

$OU_AAD_Root = "OU=AAD,OU=Cloud,OU=skupiny,DC=cezdata,DC=corp"
$OU_ADO     = "OU=DevOps,OU=Azure" + "," + $OU_AAD_Root
$OU_AzSub   = "OU=Subskripce,OU=Azure" + "," + $OU_AAD_Root
$OU_SPO     = "OU=SPO,OU=M365" + "," + $OU_AAD_Root
$OU_EXO     = "OU=EXO,OU=M365" + "," + $OU_AAD_Root
$OU_MIP     = "OU=MIP,OU=M365" + "," + $OU_AAD_Root
$OU_Pwr     = "OU=Pwr,OU=M365" + "," + $OU_AAD_Root
$OU_Lic     = "OU=License,OU=M365" + "," + $OU_AAD_Root
$OU_Intune  = "OU=Intune,OU=M365" + "," + $OU_AAD_Root
$OU_OCP     = "OU=OCP,OU=skupiny,DC=cezdata,DC=corp"


$GuestAzureApps =  @(
    #Azure DevOps	
    "499b84ac-1321-427f-aa17-267ca6975798",
    #Azure Portal
	"c44b4083-3bb0-49c1-b47d-974e53cbdf3c"
)

$NoEnumerationGroups = @(
)

$TMS_CloudPoradna_TagName_Azure = "Azure"

$knownServiceAccountPrefixes    = @("Q1","QB","QE","QL","QN","QP")
$NoExtSharingAccountPrefixes    = @("Q1","QA","QB","QD","QE","QF","QG","QL","QN","QP","QQ","QR","QS","QU","QW","QY","QZ")
$NoMFAPhoneMgmtAccountPrefixes  = @("Q1","QA","QB","QD","QE","QF","QG","QL","QN","QP","QQ","QR","QS","QU","QW","QY","QZ")
$NoPIMAccountPrefixes           = @("QP","QA","QB","QD","QE","QF","QG","QL","QN","QP","QQ","QW","QY","QZ")

$NoPIMAdminRoles = (
    #Directory Readers    
    "88d8e3e3-8f55-4a1e-953a-9b9898b8876b",
    #Message Center Privacy Reader
    "ac16e43d-7b2d-40e0-ac05-243ff356ab5b",
    #Message Center Reader
    "790c1fb9-7f7d-4f88-86a1-ef1f95c05c1b",
    #Reports Reader
    "4a5d8f65-41da-4de4-8968-e035b65339cf",
	##AAD Read Only Admin
	"731eb480-5bf2-48cc-b40f-a9296d762c10",
    "cd06886f-7ecd-424d-be67-253d025ee536",
    #Entra Id Admin Portal Access (no permissions)
    "983d2e33-2e49-4add-8269-80d7ae8fc52a",
    "a0959257-c3dc-4c84-8a16-13073a33f95e"
)

$VPN_ADM_DeviceGroup = "CiscoAnyConnect_ADM_dev"



$NoPIMUsers = (
    "Sync_WINPADSYNC_05a1af03192b@cezdata.onmicrosoft.com",
    "Sync_WINPADSYNC_3dc2419ef613@cezdata.onmicrosoft.com",
    "Sync_WINPADSYNC_f44e076c4fc3@cezdata.onmicrosoft.com",
    "Sync_WINPADSYNC2_1a95d747b373@cezdata.onmicrosoft.com",
    "Sync_WINPADSYNC2_89e795671e1c@cezdata.onmicrosoft.com",
    "admin_cezO365@cezdata.onmicrosoft.com"
)

$ExtSharingWhitelist = @(
    #qshrecinmar@cez.cz    
    "6f98eecf-3922-4627-b0b5-8edbddd1faee",
    #qssejbavit@cez.cz
    "2906dc54-f5e3-403a-8114-758ed3c3bf03",
    #QPBIDIVIZEGR@cez.cz
    "10a2b84d-cf06-4aee-a67b-09b2b326d667"
)
#CEZ_QS_69_803000_Licence_Kontrakty
$CEZ_QS_LicMgmt_Team = "328850b0-4eb3-4799-b279-a8dcf68800f2"

#CEZ_Azure_Subscription_Owners
$GroupId_AzureSubOwnersGroup = "b792181e-0095-49db-a3de-e524df001d9a"
#CEZ_Azure_Subscription_Owners_Std
$GroupId_AzureSubOwnersStd = "562285f7-9a0a-43c6-acc2-f84df6e01db5"
$GroupId_AzureSubSA = $GroupId_AzureSubOwnersStd
$GroupName_AzureSubOwnersStd = "CEZ_Azure_Subscription_Owners_Std"
$GroupName_AzureSubSA = $GroupName_AzureSubOwnersStd
#CEZ_Azure_Subscription_Owners_Std
$GroupId_AzureSubOwnersQS = "c9ecf2ee-92ad-448c-a8ac-49f9274fb88b"
$GroupName_AzureSubOwnersQS = "CEZ_Azure_Subscription_Owners_QS"
#CEZ_AZURE_BSI_ADMINS
$GroupId_AzureBSIAdmins = "e2f7c68b-ff54-4147-92ce-f3310082eecb"

#TMS_ICTS_Cloud_Poradna
$GroupId_TMS_CloudPoradna = "1959c89f-faa4-4ffb-8b53-af7b6e8561df"
#$GroupId_TMS_CloudPoradna = "dc78fa91-ec04-4b63-aef9-b33cfa4e5dbc"
#CEZ_EXO_Allow_SMTP_Client_Auth
$GroupId_SMTPAuthEnabled = "7d0dd440-f4e6-40a4-ab14-cecc74040a42"
#CEZ_EXO_Allow_IMAP_Client_access
$GroupId_IMAPEnabled = "751798f5-5e7e-47b0-836c-69b0c5bfbdd3"
#CEZ_EXO_Allow_POP3_Client_Access
$GroupId_POPEnabled = "9c7c879d-308b-44e5-8224-5dabe58656ba"
#CEZ_EXO_LitigationHold_Disable
$GroupId_LHDisabled = "5ba9efa7-eb74-46a2-876f-1b77208e5b1b"
#Poradna ČEZ ICT Services/MS_ICTS_support_cez@cez.cloud
$GroupId_O365Poradna = "d9e1b8c5-3dab-416e-9e3d-c332ce177c7b"
#CEZ_Lic_Unlicensed_Users_AAD_Prem
$GroupId_AADPremNoLic = "c4aafd32-8def-4dba-ba84-241d5f5bfc57"
#CEZ_Lic_Licensed_Users_SPO
$GroupId_SPOLicensedUsers = "4277d8d7-c0e1-426b-952d-7bc9ddd2cc99"
#CEZ_AAD_Usr_No_External_Sharing
$GroupId_NoExtSharing = "e4dba462-e76e-428b-97d3-b5522150a93a"
#CEZ_AAD_Usr_External_Sharing_Allowed
$GroupId_ExtSharingAllowed = "9c078240-7e30-40c6-9ede-4ca0a88618c3"
#CEZ_AADSync_Ext_noMail
$GroupId_ExtNoMail = "59f4326a-3147-489a-9f56-45e0d670987a"
#CEZ_AADSync_All_Users
$GroupId_AADSync_All_Users = "081a88b2-3ee5-4934-b92e-645cb2e3622d"
#CEZ_AADSync_OU_Resources
$GroupId_OUResources = "d40325bd-feac-4e19-98ad-ece37a32aa00"
#CEZ_AADSync_OU_SyncToAzure
$GroupId_OUSyncToAzure = "f13b7ff9-17cb-4fb8-8bb8-bed1a4186687"
#TMS_ICTS_DEVOPS
$GroupId_OCP_DevOpsTeam = "67a021d9-6fd9-4c0b-91f6-56d81ba6a40f"


$AADP2LicenseGroup = "CEZ_Lic_AAD_Prem_P2"
$E5SecLicenseGroup = "CEZ_Lic_MCAS_BasUsr"
$M365F3SharedLicenseGroup = "CEZ_Lic_M365_BasSrd_F3"

$OUResourcesADGroup = "CEZ_AADSync_OU_Resources"
$OUResourcesDN = "OU=Resources,OU=Groupware,OU=aplikacni,OU=uzivatele,DC=cezdata,DC=corp"
$OUSyncToAzureADGroup = "CEZ_AADSync_OU_SyncToAzure"
$OUSyncToAzureDN = "OU=SYNC_to_AZURE,OU=aplikacni,OU=uzivatele,DC=cezdata,DC=corp"

$phoneNumberRemoveChars = (" ","-","(",")","#","+")

$prefixLength1 = (1,7)
$prefixLength2 = (20,27,30,31,32,33,34,36,39,40,41,43,44,45,46,47,48,49,51,52,53,54,55,56,56,57,58,60,61,62,63,64,64,65,66,81,82,84,86,90,91,92,93,94,95,98)
$prefixLength3 = (211,212,213,216,218,220,
                221,222,223,224,225,226,227,228,229,
                230,231,232,233,234,235,236,237,238,239,
                240,241,242,243,244,245,246,247,248,249,
                250,251,252,253,254,255,256,257,258,
                260,261,262,263,264,265,266,267,268,269,290,
                291,297,298,299,
                350,351,352,353,354,355,356,357,358,359,370,371,372,373,374,375,376,377,378,380,381,382,383,385,386,387,389,
                420,421,423,
                500,501,502,503,504,505,506,507,508,509,590,591,592,593,594,595,596,597,598,599,
                670,673,674,675,676,677,678,679,680,681,682,683,685,686,687,688,689,690,691,692,
                800,808,850,852,853,855,856,870,880,881,882,883,886,
                960,961,962,963,964,965,966,967,968,970,971,972,973,974,975,976,977,992,993,994,995,996,997,998
)

$MSLicDataUrl = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
$MSLicDataPath = "d:\data\m365-licensing\ms-licensing-product-names-service-plans.csv"

################################

$M365E3SKUIds = @("05e9a617-0261-4cee-bb44-138d3ef5d965")
$M365E3UATSKUIds = @("c2ac2ee4-9bb1-47e4-8541-d689c7e83371")
$M365E5SKUIds = @("06ebc4ee-1bb5-47dd-8120-11324bc54e06")
$M365F3SKUIds = @("66b55226-6b4f-492c-910c-a3b7a3c9d993")
$M365CopilotSKUIds = @("639dec6b-bb19-468b-871c-c5c441c4b0cb")
$M365E5SecSKUIds = @("26124093-3d78-432b-b5dc-48bf992543d5")
$TeamsPremSKUIds = @("36a0f3b3-adb5-49ea-bf66-762134cf063a")
$PwrAutPremSKUIds = @("eda1941c-3c4f-4995-b5eb-e85a42175ab9")
$PwrAppPremSKUIds = @("b30411f5-fea1-4a59-9ad9-3db7c7ead579")

################################

$AADP1LicensePlans = @("41781fb2-bc02-4b7c-bd55-b576c07bb09d")

$AADP2LicensePlans = @("eec0eb4f-6444-4f95-aba0-50c24d67f998")

$EXOLicensePlans = @("efb87545-963c-4e0d-99df-69c6916d9eb0",
                "4a82b400-a79f-41a4-b4e2-e94f5787b113",
                "9aaf7827-d63c-4b61-89c3-182f06f82e5c"
)

$SPOLicensePlans = @("e95bec33-7c88-4a70-8e19-b10bd9d0c014",
                "5dbe027f-2339-4123-9542-606e4d348a72",
                "fe71d6c3-a2ea-4499-9778-da042bf08063",
                "902b47e5-dcb2-4fdc-858b-c63a90a2bdb9",
                "c7699d2e-19aa-44de-8edf-1736da088ca1"
)

$TMSLicensePlans = @("57ff2da0-773e-42df-b2af-ffb7a2317929",
                "9104f592-f2a7-4f77-904c-ca5a5715883f",
                "cc8c0802-a325-43df-8cba-995d0c6cb373",
                "f8b44f54-18bb-46a3-9658-44ab58712968",
                "78b58230-ec7e-4309-913c-93a45cc4735b",
                "0504111f-feb8-4a3c-992a-70280f9a2869",
                "92c6b761-01de-457a-9dd9-793a975238f7",
                "0374d34c-6be4-4dbb-b3f0-26105db0b28a",
                "ec17f317-f4bc-451e-b2da-0167e5c260f9",
                "8081ca9c-188c-4b49-a8e5-c23b5e9463a8"
)
$M365CopilotLicensePlans = @("82d30987-df9b-4486-b146-198b21d164c7",
                "931e4a88-a67f-48b5-814f-16a5f1e6028d",
                "b95945de-b3bd-46db-8437-f2beb6ea2347",
                "a62f8878-de10-42f3-b68f-6149a25ceb97",
                "3f30311c-6b1e-48a4-ab79-725b469da960",
                "89f1c4c8-0878-40f7-804d-869c9128ab5d"
)

$PwrAutLicensePlans = @("c5002c70-f725-4367-b409-f0eff4fee6c0",
                "7e6d7d78-73de-46ba-83b1-6d25117334ba",
                "76846ad7-7776-4c40-a281-a386362dd1b9",
                "50e68c76-46c6-4674-81f9-75456511b170",
                "5d798708-6473-48ad-9776-3acc301c40af",
                "d20bfa21-e9ae-43fc-93c2-20783f0840c3",
                "07699545-9485-468e-95b6-2fca3738be01",
                "dc789ed8-0170-4b65-a415-eb77d5bb350a",
                "fa200448-008c-4acb-abd4-ea106ed2199d",
                "1ec58c70-f69c-486a-8109-4b87ce86e449",
                "bd91b1a4-9f94-4ecf-b45b-3a65e5c8128a",
                "c7ce3f26-564d-4d3a-878d-d8ab868c85fe",
                "0f9b09cb-62d1-4ff4-9129-43f4996f83f4",
                "375cd0ad-c407-49fd-866a-0bff4f8a9a4d"
)

$PwrAppLicensePlans = @("e61a2945-1d4e-4523-b6e7-30ba39d20f32",
                "874fc546-6efe-4d22-90b8-5c4e7aa59f4b",
                "c68f8d98-5534-41c8-bf36-22fa496fa792",
                "d5368ca3-357e-4acb-9c21-8495fb025d1f",
                "9c0dab89-a30c-4117-86e7-97bda240acd2",
                "ea2cf03b-ac60-46ae-9c1d-eeaeb63cec86",
                "52e619e2-2730-439a-b0d3-d09ab7e8b705",
                "e0287f9f-e222-4f98-9a83-f379e249159a",
                "a2729df7-25f8-4e63-984b-8a8484121554",
                "92f7a6f3-b89b-4bbd-8c30-809e6da5ad1c"
)


[array]$IntuneLicensePlans = @("c1ec4a95-1f05-45b3-a911-aa3fa01094f5",
                "3e170737-c728-4eae-bbb9-3f3360f7184c"
)

[hashtable]$mdmApp_DB = @{
    "0000000a-0000-0000-c000-000000000000" = "Microsoft Intune";
    "54b943f8-d761-4f8d-951e-9cea1846db5a" = "System Center Configuration Manager";
    "7add3ecd-5b01-452e-b4bf-cdaf9df1d097" = "Office 365 Mobile Device Management"
}

$alternateMobilePhoneMethodId  = "b6332ec1-7057-4abe-9331-3d72feddfe41"
$officePhoneMethodId           = "e37fc753-ff3b-4958-9484-eaa9425c82bc"
$mobilePhoneMethodId           = "3179e48a-750b-4051-897c-87b9720928f7"

[hashtable]$UsrToSysMethodConv_DB = @{
    "push" = "phoneappnotification";
    "voiceMobile" = "voice";
    "voiceOffice" = "voice";
    "voiceAlternateMobile" = "voice";
    "sms" = "sms";
    "oath" = "phoneappotp"
}

[hashtable]$SysToUsrMethodConv_DB = @{
    "phoneappnotification" = "push";
    "voice" = "voiceMobile"
}

$AADGuestDomainsShowInGAL = @("elevion.de",
                "koflerenergies.com",
                "inewa.it"
)

[array]$PSColorList  = @(
    "Black",
    "DarkBlue",
    "DarkGreen",
    "DarkCyan",
    "DarkRed",
    "DarkMagenta",
    "DarkYellow",
    "Gray",
    "DarkGray",
    "Blue",
    "Green",
    "Cyan",
    "Red",
    "Magenta",
    "Yellow",
    "White"
)

$GraphAPIRetryableErrors = @(
    # Bad Gateway
    "server returned an error: (502)",
    # Server Unavailable
    "server returned an error: (503)",
    #Gateway Timeout
    "server returned an error: (504)"
)

$GraphAPIThrottlingErrors = @(
    #Too many requests
    "server returned an error: (429)"
)

$GraphAPIAuthErrors = @(
    #Unathorized
    "server returned an error: (401)"
)