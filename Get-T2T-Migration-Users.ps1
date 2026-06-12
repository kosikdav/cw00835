#######################################################################################################################
# Get-T2T-Migration-Users.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1
. $ScriptPath\include\include-functions-T2T.ps1

#######################################################################################################################

$LogFolder			= "t2t"
$LogFilePrefix		= "t2t-migration-users"
$LogFileFreq		= "Y"

$OutputFolder = "t2t-ujvrez\migration-users"
$OutputFilePrefix = "t2t-migration-users"

#######################################################################################################################

. $ScriptPath\include-Script-Start-Include.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Freq $LogFileFreq -Ext "log"
$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Freq "YMDHM" -Ext "csv"

[array]$Report = @()

$BatchPrefix = "UJV"
$MigrationEndpoint = "UJVREZ_T2T_EXO_MIGRATION_ENDPOINT"
$DstMailRoutingDomain = "cezdata.mail.onmicrosoft.com"

$Dst_AD_OU_list = @(
	"OU=CVREZ,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=RadioMedic,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=EGP,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=EngineeringPraha,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=iCVREZ,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",	
	"OU=NQ-Safe,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=UJVREZ,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=VZUP,OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp",
	"OU=UJV,OU=aplikacni,OU=uzivatele,DC=cezdata,DC=corp"
)

$DstADSearchBaseAll  = "OU=uzivatele,DC=cezdata,DC=corp"
$DstADSearchBaseSKC  = "OU=skupinaCEZ,OU=uzivatele,DC=cezdata,DC=corp"

$Dst_AD_OU_list_App = @(
	"OU=UJV,OU=aplikacni,OU=uzivatele,DC=cezdata,DC=corp"
)

$DstADUserProperties = @(
	'displayName',
	'mail',
	'distinguishedName',
	'samAccountName',
	'userPrincipalName',
	'extensionAttribute10'
	'targetAddress'
)

$MigrationUsers = @()
$Report = @()
$Dst_UserDB = @{}

#######################################################################################################################

. $IncFile_StdLogStartBlock

if ($InteractiveRun) {
	$ADCredentialPath = "c:\cred\qp_aad_grp_mgmt\qp_aad_grp_mgmt_qskosikdav.cred"
}
else {
	$ADCredentialPath = $aad_grp_mgmt_cred
}

$ADCredential = Import-Clixml -Path $ADCredentialPath

$DstGetADUserParams = @{
	Filter = "ObjectClass -eq 'user'"
	SearchBase = $DstADSearchBaseAll
	Properties = $DstADUserProperties
	Credential = $ADCredential
}

$Src_AppReg_LOG_READER 			= $AppReg_UJVREZ_LOG_READER
$Src_AppReg_EXO_MGMT 			= $AppReg_UJVREZ_EXO_MGMT
$Dst_AppReg_EXO_MGMT 			= $AppReg_EXO_MGMT
$Dst_AppReg_LOG_READER 			= $AppReg_LOG_READER

$Src_T2T_EXO_MIGRATION_GROUP 	= $UJVREZ_T2T_EXO_MIGRATION_GROUP

$DstMailRoutingDomain = "cezdata.mail.onmicrosoft.com"

$Src_AppReg_LOG_READER 			= $AppReg_UJVREZ_LOG_READER
$Src_AppReg_EXO_MGMT 			= $AppReg_UJVREZ_EXO_MGMT   
$Src_T2T_EXO_MIGRATION_GROUP 	= $UJVREZ_T2T_EXO_MIGRATION_GROUP

#Request-MSALToken -AppRegName $Dst_AppReg_LOG_READER -TTL 30
Connect-EXOService -AppRegName $Dst_AppReg_EXO_MGMT  -TTL 120

$MigrationBatches = Get-MigrationBatch
$MigrationBatches = $MigrationBatches | Where-Object { $_.Identity -like "$BatchPrefix*" }
Write-host "Found $($MigrationBatches.Count) migration batches with prefix '$BatchPrefix'"

foreach ($Batch in $MigrationBatches) {
	write-host ($Batch.Identity.ToString()).PadRight(25) -ForegroundColor Green -NoNewline
	$Result = Get-MigrationUser -BatchId $Batch.Identity
	write-host ": $($Result.Count)"
	$MigrationUsers += $Result
}

#read CEZDATA AD users and filter only those with enabled account
write-host "DST AD users - reading from AD (long running operation)..." -NoNewline
$DstADUsers = Get-ADUser @DstGetADUserParams | Select-Object $DstADUserProperties
write-host "done ($($DstADUsers.count))"

#filter out users with samAccountName starting with QT
$DstADUsers = $DstADUsers | Where-Object { $_.samAccountName -notlike 'QT*' }
write-host "DST AD users - (filtered out QT): $($DstADUsers.count)"

#filter out users with samAccountName starting with QK
$DstADUsers = $DstADUsers | Where-Object { $_.samAccountName -notlike 'QK*' }
write-host "DST AD users - (filtered out QK): $($DstADUsers.count)"

#OU property to each user object by parsing it from DistinguishedName, we will need it for filtering users by OU and for reporting
write-host "DST AD users - adding OU property..." -NoNewline
foreach ($user in $DstADUsers) {    
	$user | Add-Member -NotePropertyName 'OU' -NotePropertyValue ( $user.DistinguishedName -replace '^CN=[^,]+,' )
}
write-host "done"

#filter out only users from specific OUs
$DstADUsers = $DstADUsers | Where-Object { $_.OU -in $Dst_AD_OU_list }
write-host "DST AD users - filtered by OU: $($DstADUsers.count)"

foreach ($user in $DstADUsers) {
	if ($user.userPrincipalName) {
		$userObject = [PSCustomObject]@{
			displayName = $user.displayName
			samAccountName = $user.samAccountName
			OU = $user.OU
			ext10 = $user.extensionAttribute10
			targetAddress = $user.targetAddress
		}
		$Dst_UserDB.add($user.userPrincipalName, $userObject)
	}
}
write-host "DST AD userDB: $($Dst_UserDB.count)"

foreach ($User in $MigrationUsers) {
	$ADUser = $targetAddress = $null
	if ($Dst_UserDB.ContainsKey($User.Identity)) {
		$ADUser = $Dst_UserDB[$User.Identity]
	}
	if ($ADUser.targetAddress -like "SMTP:*") {
		$targetAddress = $ADUser.targetAddress.Substring(5)
	}
	$Report += [PSCustomObject]@{
		Identity = $User.Identity
		displayName = $ADUser.displayName
		targetAddress = $targetAddress
		ext10 = $ADUser.ext10
		samAccountName = $ADUser.samAccountName
		OU = $ADUser.OU
		Guid = $User.Guid
		BatchId = $User.BatchId
		MailboxIdentifier = $User.MailboxIdentifier   
		MailboxEmailAddress = $User.MailboxEmailAddress
		Status = $User.Status
		StatusSummary = $User.StatusSummary
		MigrationType = $User.MigrationType
		State = $User.State
		Flags = $User.Flags
		WorkflowStep = $User.WorkflowStep
		WorkflowStage = $User.WorkflowStage
		EstimatedTotalCount = $User.EstimatedTotalCount
		EstimatedTotalSizeEstimatedArchiveCount = $User.EstimatedTotalSizeEstimatedArchiveCount
		EstimatedArchiveSize = $User.EstimatedArchiveSize
		RemoteIdentifier = $User.RemoteIdentifier
		RecipientType = $User.RecipientType
		SkippedItemCount = $User.SkippedItemCount
		SyncedItemCount = $User.SyncedItemCount
		TransferredItemCount = $User.TransferredItemCount
		SyncedFolderCount = $User.SyncedFolderCount
		MailboxGuid = $User.MailboxGuid
		RequestGuid = $User.RequestGuid
		TriggeredAction = $User.TriggeredAction
		DataConsistencyScore = $User.DataConsistencyScore
		HasUnapprovedSkippedItems = $User.HasUnapprovedSkippedItems
		ErrorSummary = $User.ErrorSummary
		LastSuccessfulSyncTime = $User.LastSuccessfulSyncTime
		LastSubscriptionCheckTime = $User.LastSubscriptionCheckTime
		Diagnostics = $User.Diagnostics
		DiagnosticInfo = $User.DiagnosticInfo
	}
}

Export-Report -Report $Report -Path $OutputFile -SortProperty "Identity"

#######################################################################################################################

. $IncFile_StdLogEndBlock
