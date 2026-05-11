#######################################################################################################################
# Get-ADUsersGroupsReport-UJVREZ.ps1
#######################################################################################################################

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

. "$ScriptPath\include-function-Write-Log.ps1"
. "$ScriptPath\include-function-Request-MSALtoken.ps1"

<#
$appName = "CEZ_UJV_T2T_MIGRATION_STORAGE"
$certYears = 10
$certPassword = "123456xxxxxx"
$StartDate = (Get-Date).AddDays(-1)
$EndDate = (Get-Date).AddYears($certYears)
$Password = ConvertTo-SecureString $certPassword -AsPlainText -Force
Create-SelfSignedCertificate.ps1 -CommonName $AppName -StartDate $StartDate -EndDate $EndDate -Password $Password -Force
#>

#######################################################################################################################
# parameter definitions
#######################################################################################################################

$LogFile = "d:\logs\$($ScriptName).log"

$AppName = "CEZ_UJV_T2T_MIGRATION_STORAGE"
$StorageAccount = "cezt2tstore"
$Container = "ujvrez"

$ClientId   = 'c0b8e48f-aacc-4db1-bcae-8a7341ce436d'
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"

#$ADUserSearchBase  = "OU=uzivatele,DC=cezdata,DC=corp"
#$ADGroupSearchBase  = "OU=uzivatele,DC=cezdata,DC=corp"
$ADUserSearchBase  = "OU=SYNC_to_AZURE,OU=aplikacni,OU=uzivatele,DC=cezdata,DC=corp"
$ADGroupSearchBase = "OU=EXO,OU=M365,OU=AAD,OU=Cloud,OU=skupiny,DC=cezdata,DC=corp"

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

#######################################################################################################################
# variable initialization
#######################################################################################################################

$Today = (Get-Date).ToString("yyyy-MM-dd")
$Now   = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
$interactiveRun = [Environment]::UserInteractive

[array]$ReportUsr = @()
[array]$ReportGrpLst = @()
[array]$ReportGrpMem = @()

$UploadParams = @{
	StorageAccount = $StorageAccount
	Container = $container
}

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
# function definitions
#######################################################################################################################
function Invoke-AzureBlobUpload {
	[CmdletBinding()]
	param (
		[string]$StorageAccount,
		[string]$Container,
		[string]$BlobName,
		[string]$Content,
		[string]$AccessToken
	)

	$uri = "https://$StorageAccount.blob.core.windows.net/$Container/$BlobName"
	$headers  = @{
		Authorization = "Bearer $AccessToken"
		"x-ms-version" = "2020-04-08"
		"x-ms-blob-type" = "BlockBlob"
		"Content-Type" = "text/csv"
	}

	try {
		Invoke-RestMethod -Uri $uri -Method Put -Headers $headers -Body $Content
		Write-Log "Upload successful: $BlobName"
	}
	catch {
		Write-Log "Failed to upload blob: $_" -MessageType "ERR"
	}
}

#######################################################################################################################
# main script logic
#######################################################################################################################
if ($InteractiveRun) {
	$ADCredentialPath = "c:\cred\qp_aad_grp_mgmt\qp_aad_grp_mgmt_qskosikdav.cred"
}
else {
	$ADCredentialPath = $aad_grp_mgmt_cred
}

Write-Log "--------------------------------------------------------------"
Write-Log "Script file: $($ScriptPath)\$($ScriptName)"
If ([Environment]::UserInteractive) {
    Write-Log "Running interactively" -ForegroundColor DarkBlue -BackgroundColor Green
}
Else {
    Write-Log "Running non-interactively"
}
Write-Log "Log file: $($LogFile)"
Write-Log "ADCredentialPath: $($ADCredentialPath)"
Write-Log "Script start"



$ADCredential = Import-Clixml -Path $ADCredentialPath

#read UJVREZ AD users
$ADUsers = Get-ADUser @GetADUserParams | Select-Object $ADUserProperties
Write-Log "AD users: $($ADUsers.count))"

#read UJVREZ AD groups
$ADGroups = Get-ADGroup @GetADGroupParams | Select-Object $ADGroupProperties
Write-Log "AD groups: $($ADGroups.count))"

#process users
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

#process groups and group members
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
Write-Log "AD group memberships: $($ReportGrpMem.count)"

#get MSAL access token from Entra ID
$Certificate = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$($AppName)" }
$AccessToken = Request-MSALtoken -TenantId $TenantId -ClientId $ClientId -Certificate $Certificate -Scope "https://storage.azure.com/.default"
$UploadParams.AccessToken = $AccessToken

#create CSV content
$csvStringUsr = $ReportUsr | ConvertTo-Csv -NoTypeInformation -Delimiter ',' | Out-String
$csvStringGrpLst = $ReportGrpLst | ConvertTo-Csv -NoTypeInformation -Delimiter ',' | Out-String
$csvStringGrpMem = $ReportGrpMem | ConvertTo-Csv -NoTypeInformation -Delimiter ',' | Out-String

#upload CSV content to Azure Blob Storage
$UploadParams.BlobName = "ADUsr_$Today.csv"
$UploadParams.Content = $csvStringUsr
Invoke-AzureBlobUpload @UploadParams

$UploadParams.BlobName = "ADGrpLst_$Today.csv"
$UploadParams.Content = $csvStringGrpLst
Invoke-AzureBlobUpload @UploadParams

$UploadParams.BlobName = "ADGrpMem_$Today.csv"
$UploadParams.Content = $csvStringGrpMem
Invoke-AzureBlobUpload @UploadParams

Write-Log "Script finish"