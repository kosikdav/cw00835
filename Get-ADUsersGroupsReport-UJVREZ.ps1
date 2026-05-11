#######################################################################################################################
# Get-ADUsersGroupsReport-UJVREZ.ps1
#######################################################################################################################

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. "$ScriptPath\include-function-Request-MSALtoken.ps1"

<#
$appName = "CEZ_UJV_T2T_MIGRATION_STORAGE"
$certYears = 10
$certPassword = "xxxxxxds3$$%x"
Create-SelfSignedCertificate.ps1 -CommonName $appName -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date).AddYears($certYears) -Password (ConvertTo-SecureString $certPassword -AsPlainText -Force) -Force
#>

$Today = (Get-Date).ToString("yyyy-MM-dd")
$Now   = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
$interactiveRun = [Environment]::UserInteractive

$AppName = "CEZ_UJV_T2T_MIGRATION_STORAGE"
$ClientId   = 'c0b8e48f-aacc-4db1-bcae-8a7341ce436d'
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"
$Certificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }

#$ADUserSearchBase  = "OU=uzivatele,DC=cezdata,DC=corp"
#$ADGroupSearchBase  = "OU=uzivatele,DC=cezdata,DC=corp"
$ADUserSearchBase  = "OU=SYNC_to_AZURE,OU=aplikacni,OU=uzivatele,DC=cezdata,DC=corp"
$ADGroupSearchBase = "OU=EXO,OU=M365,OU=AAD,OU=Cloud,OU=skupiny,DC=cezdata,DC=corp"

[array]$ReportUsr = @()
[array]$ReportGrpLst = @()
[array]$ReportGrpMem = @()

#######################################################################################################################

if ($InteractiveRun) {
	$ADCredentialPath = "c:\cred\qp_aad_grp_mgmt\qp_aad_grp_mgmt_qskosikdav.cred"
}
else {
	$ADCredentialPath = $aad_grp_mgmt_cred
}

$ADCredential = Import-Clixml -Path $ADCredentialPath

$ADUserProperties = @(
	'displayName',
	'mail',
	'distinguishedName',
	'samAccountName',
	'userPrincipalName',
	'objectGUID',
	'employeeId'
	'employeeNumber'
)

$ADGroupProperties = @(
	"DistinguishedName",
	"GroupCategory",
	"GroupScope",
	"Name",
	"mail",
	"ObjectGUID",
	"SamAccountName",
	"SID"
)

$GetADUserParams = @{
	Filter = '*'
	SearchBase = $ADUserSearchBase
	Properties = $ADUserProperties
	Credential = $ADCredential
}

$GetADGroupParams = @{
	Filter = '*'
	SearchBase = $ADGroupSearchBase
	Properties = $ADGroupProperties
	Credential = $ADCredential
}


#######################################################################################################################

#read UJVREZ AD users
write-host "AD users - reading from AD..." -NoNewline
$ADUsers = Get-ADUser @GetADUserParams | Select-Object $ADUserProperties
write-host "done ($($ADUsers.count))"

#read UJVREZ AD groups
write-host "AD groups - reading from AD..." -NoNewline
$ADGroups = Get-ADGroup @GetADGroupParams | Select-Object $ADGroupProperties
write-host "done ($($ADGroups.count))"


write-host "AD users - processing user report..." -NoNewline
foreach ($user in $ADUsers) {
	$userObject = [PSCustomObject]@{
		userPrincipalName = $user.userPrincipalName
		mail = $user.mail
		displayName = $user.displayName
		samAccountName = $user.samAccountName
		distinguishedName = $user.distinguishedName
		objectGUID = $user.objectGUID
		employeeNumber = $user.employeeNumber
		employeeId = $user.employeeId
	}
	$ReportUsr += $userObject
}
write-host "done ($($ReportUsr.count))"

write-host "AD groups - processing groups and group membership..."
foreach ($group in $ADGroups) {
	write-host $group.SamAccountName
	$groupObject = [PSCustomObject]@{
		Name = $group.Name
		SamAccountName = $group.SamAccountName
		DistinguishedName = $group.DistinguishedName
		GroupCategory = $group.GroupCategory
		GroupScope = $group.GroupScope
		ObjectGUID = $group.ObjectGUID
		SID = $group.SID
	}
	$ReportGrpLst += $groupObject

	$GroupMembers = Get-ADGroupMember -Identity $group.SamAccountName -Credential $ADCredential -ErrorAction SilentlyContinue
	
	foreach ($member in $GroupMembers) {
		$memberObject = [PSCustomObject]@{
			GroupSamAccountName = $group.SamAccountName
			GroupObjectGUID = $group.ObjectGUID
			MemberSamAccountName = $member.SamAccountName
			MemberObjectGUID = $member.objectGUID
		}
		$ReportGrpMem += $memberObject
	}
}
write-host "done ($($ReportGrpMem.count))"

$ReportUsr | Export-Csv -Path $OutputFileUsr -NoTypeInformation -Encoding UTF8 -Delimiter ','
$ReportGrpLst | Export-Csv -Path $OutputFileGrpLst -NoTypeInformation -Encoding UTF8 -Delimiter ','
$ReportGrpMem | Export-Csv -Path $OutputFileGrpMem -NoTypeInformation -Encoding UTF8 -Delimiter ','

$csvStringUsr = $ReportUsr | ConvertTo-Csv -NoTypeInformation -Encoding UTF8 -Delimiter ',' | Out-String
$csvStringGrpLst = $ReportGrpLst | ConvertTo-Csv -NoTypeInformation -Encoding UTF8 -Delimiter ',' | Out-String
$csvStringGrpMem = $ReportGrpMem | ConvertTo-Csv -NoTypeInformation -Encoding UTF8 -Delimiter ',' | Out-String

$fileBytes  = [Text.Encoding]::UTF8.GetBytes($csvString)
# --- Upload blob (replace $fileBytes in the existing script) ---
$blobName = "export.csv"
$uri      = "https://$storageAccount.blob.core.windows.net/$container/$blobName"
$headers  = @{
	Authorization = "Bearer $accessToken"
	"x-ms-version" = "2020-04-08"
	"x-ms-blob-type" = "BlockBlob"
	"Content-Type" = "text/csv"
}
