#######################################################################################################################
#######################################################################################################################
# INCLUDE-FUNCTIONS-AUDIT-RECORDS
#######################################################################################################################
#######################################################################################################################
#
#
#
########################################################################################
# Get-AuditRecordTypeFromCode
########################################################################################
function Get-AuditRecordTypeFromCode {
	[alias("GetAuditRecordTypeFromCode")]
	param (
		[Parameter(Mandatory)]
		[int]$RecordTypeCode
	)
	# main function body ##################################
	if (($RecordTypeCode -ge 1) -and ($RecordTypeCode -le 230)) {
		switch ($RecordTypeCode) {
			1 { $recordType = "ExchangeAdmin" }
			2 { $recordType = "ExchangeItem" }
			3 { $recordType = "ExchangeItemGroup" }
			4 { $recordType = "SharePoint" }
			6 { $recordType = "SharePointFileOperation" }
			7 { $recordType = "OneDrive" }
			8 { $recordType = "AzureActiveDirectory" }
			9 { $recordType = "AzureActiveDirectoryAccountLogon" }
			10 { $recordType = "DataCenterSecurityCmdlet" }
			11 { $recordType = "ComplianceDLPSharePoint" }
			13 { $recordType = "ComplianceDLPExchange" }
			14 { $recordType = "SharePointSharingOperation" }
			15 { $recordType = "AzureActiveDirectoryStsLogon" }
			16 { $recordType = "SkypeForBusinessPSTNUsage" }
			17 { $recordType = "SkypeForBusinessUsersBlocked" }
			18 { $recordType = "SecurityComplianceCenterEOPCmdlet" }
			19 { $recordType = "ExchangeAggregatedOperation" }
			20 { $recordType = "PowerBIAudit" }
			21 { $recordType = "CRM" }
			22 { $recordType = "Yammer" }
			23 { $recordType = "SkypeForBusinessCmdlets" }
			24 { $recordType = "Discovery" }
			25 { $recordType = "MicrosoftTeams" }
			28 { $recordType = "ThreatIntelligence" }
			29 { $recordType = "MailSubmission" }
			30 { $recordType = "MicrosoftFlow" }
			31 { $recordType = "AeD" }
			32 { $recordType = "MicrosoftStream" }
			33 { $recordType = "ComplianceDLPSharePointClassification" }
			34 { $recordType = "ThreatFinder" }
			35 { $recordType = "Project" }
			36 { $recordType = "SharePointListOperation" }
			37 { $recordType = "SharePointCommentOperation" }
			38 { $recordType = "DataGovernance" }
			39 { $recordType = "Kaizala" }
			40 { $recordType = "SecurityComplianceAlerts" }
			41 { $recordType = "ThreatIntelligenceUrl" }
			42 { $recordType = "SecurityComplianceInsights" }
			43 { $recordType = "MIPLabel" }
			44 { $recordType = "WorkplaceAnalytics" }
			45 { $recordType = "PowerAppsApp" }
			46 { $recordType = "PowerAppsPlan" }
			47 { $recordType = "ThreatIntelligenceAtpContent" }
			48 { $recordType = "LabelContentExplorer" }
			49 { $recordType = "TeamsHealthcare" }
			50 { $recordType = "ExchangeItemAggregated" }
			51 { $recordType = "HygieneEvent" }
			52 { $recordType = "DataInsightsRestApiAudit" }
			53 { $recordType = "InformationBarrierPolicyApplication" }
			54 { $recordType = "SharePointListItemOperation" }
			55 { $recordType = "SharePointContentTypeOperation" }
			56 { $recordType = "SharePointFieldOperation" }
			57 { $recordType = "MicrosoftTeamsAdmin" }
			58 { $recordType = "HRSignal" }
			59 { $recordType = "MicrosoftTeamsDevice" }
			60 { $recordType = "MicrosoftTeamsAnalytics" }
			61 { $recordType = "InformationWorkerProtection" }
			62 { $recordType = "Campaign" }
			63 { $recordType = "DLPEndpoint" }
			64 { $recordType = "AirInvestigation" }
			65 { $recordType = "Quarantine" }
			66 { $recordType = "MicrosoftForms" }
			67 { $recordType = "ApplicationAudit" }
			68 { $recordType = "ComplianceSupervisionExchange" }
			69 { $recordType = "CustomerKeyServiceEncryption" }
			70 { $recordType = "OfficeNative" }
			71 { $recordType = "MipAutoLabelSharePointItem" }
			72 { $recordType = "MipAutoLabelSharePointPolicyLocation" }
			73 { $recordType = "MicrosoftTeamsShifts" }
			75 { $recordType = "MipAutoLabelExchangeItem" }
			76 { $recordType = "CortanaBriefing" }
			78 { $recordType = "WDATPAlerts" }
			82 { $recordType = "SensitivityLabelPolicyMatch" }
			83 { $recordType = "SensitivityLabelAction" }
			84 { $recordType = "SensitivityLabeledFileAction" }
			85 { $recordType = "AttackSim" }
			86 { $recordType = "AirManualInvestigation" }
			87 { $recordType = "SecurityComplianceRBAC" }
			88 { $recordType = "UserTraining" }
			89 { $recordType = "AirAdminActionInvestigation" }
			90 { $recordType = "MSTIC" }
			91 { $recordType = "PhysicalBadgingSignal" }
			93 { $recordType = "AipDiscover" }
			94 { $recordType = "AipSensitivityLabelAction" }
			95 { $recordType = "AipProtectionAction" }
			96 { $recordType = "AipFileDeleted" }
			97 { $recordType = "AipHeartBeat" }
			98 { $recordType = "MCASAlerts" }
			99 { $recordType = "OnPremisesFileShareScannerDlp" }
			100 { $recordType = "OnPremisesSharePointScannerDlp" }
			101 { $recordType = "ExchangeSearch" }
			102 { $recordType = "SharePointSearch" }
			103 { $recordType = "PrivacyInsights" }
			105 { $recordType = "MyAnalyticsSettings" }
			106 { $recordType = "SecurityComplianceUserChange" }
			107 { $recordType = "ComplianceDLPExchangeClassification" }
			109 { $recordType = "MipExactDataMatch" }
			113 { $recordType = "MS365DCustomDetection" }
			147 { $recordType = "CoreReportingSettings" }
			148 { $recordType = "ComplianceConnector" }
			154 { $recordType = "OMEPortal" }
			174 { $recordType = "DataShareOperation" }
			181 { $recordType = "EduDataLakeDownloadOperation" }
			183 { $recordType = "MicrosoftGraphDataConnectOperation" }
			186 { $recordType = "PowerPagesSite" }
			216 { $recordType = "Viva Goals" }
			217 { $recordType = "MicrosoftGraphDataConnectConsent" }
			230 { $recordType = "TeamsUpdates" }
			Default { $recordType = "unknown" }
		}
		return $recordType
	}
	else {
		return $null
	}
}

########################################################################################
# Get-AuditUserTypeFromCode
########################################################################################
function Get-AuditUserTypeFromCode {
	[alias("GetAuditUserTypeFromCode")]
	param (
		[Parameter(Mandatory)]
		[int]$UserTypeCode
	)
	if (($UserTypeCode -ge 0) -and ($UserTypeCode -le 8)) {
		switch ($UserTypeCode) {
			0 { $userType = "regular_user" }
			1 { $userType = "reserved" }
			2 { $userType = "admin" }
			3 { $userType = "dc_admin" }
			4 { $userType = "system" }
			5 { $userType = "application" }
			6 { $userType = "service_principal" }
			7 { $userType = "custom_policy" }
			8 { $userType = "system_policy" }	
		}
		return $userType
	}
	else {
		return $null
	}
}
