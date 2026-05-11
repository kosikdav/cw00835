#######################################################################################################################
# Get-ADUsersGroupsReport-UJVREZ.ps1
#######################################################################################################################

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

. $ScriptPath\include-function-Request-MSALtoken.ps1

<#
$AppName = "CEZ_UJV_T2T_MIGRATION_STORAGE"
$CertYears = 10
$CertPassword = "12345678"
$StartDate = (Get-Date).AddDays(-1)
$EndDate = (Get-Date).AddYears($CertYears)
$Password = ConvertTo-SecureString $CertPassword -AsPlainText -Force

Create-SelfSignedCertificate.ps1 -CommonName $AppName -StartDate $StartDate -EndDate $EndDate -Password $Password -Force
#>

$interactiveRun = [Environment]::UserInteractive

$AppName = "CEZ_UJV_T2T_MIGRATION_STORAGE"
$StorageAccount = "cezt2tstore"
$Container = "ujvrez"

$ClientId   = 'c0b8e48f-aacc-4db1-bcae-8a7341ce436d'
$TenantId   = "b233f9e1-5599-4693-9cef-38858fe25406"

$ADUserSearchBase  = "DC=ad,DC=ujv,DC=cz"
$ADGroupSearchBase  = "DC=ad,DC=ujv,DC=cz"

$ADUserProperties = @(
	'displayName',
	'mail',
	'distinguishedName',
	'samAccountName',
	'userPrincipalName',
	'objectGUID',
	'employeeId',
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

$interactiveRun = [Environment]::UserInteractive

[array]$ReportUsr = @()
[array]$ReportGrpLst = @()
[array]$ReportGrpMem = @()

$UploadParams = @{
	StorageAccount = $StorageAccount
	Container = $container
}

#######################################################################################################################
# function definitions
#######################################################################################################################

function Write-Log {
	[CmdletBinding(SupportsShouldProcess = $true)]
	param (
		[Parameter(Position=0)][string]$String,
		[string][ValidateSet("Info","Warning","Warn","Error","Err")]$MessageType = "Info"
	)
	# main function body ##################################
	$File = $script:LogFile
	$TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
	switch ($MessageType) {
		"Info"		{$LineType = "INFO"}
		"Warning" 	{$LineType = "WARN"}
		"Error" 	{$LineType = "ERR"}
		Default 	{$LineType = "INFO"}
	}
	$LinePrefix = $TimeStamp + " [" + ($LineType.PadRight(4," ")).ToUpper() + "] "
	Add-Content $File -Value ($LinePrefix + $String)
	Write-Host $String
}

function Invoke-AzureBlobUpload {
	[CmdletBinding()]
	param (
		[string]$StorageAccount,
		[string]$Container,
		[string]$BlobName,
		[string]$Content,
		[string]$AccessToken
	)
	# main function body ##################################
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
    Write-Log "Running interactively"
}
Else {
    Write-Log "Running non-interactively"
}
Write-Log "Log file: $($LogFile)"
Write-Log "ADCredentialPath: $($ADCredentialPath)"
Write-Log "Script start"

$ADCredential = Import-Clixml -Path $ADCredentialPath

#read UJVREZ AD users
$GetADUserParams = @{
	Filter = '*'
	SearchBase = $ADUserSearchBase
	Properties = $ADUserProperties
	Credential = $ADCredential
}
$ADUsers = Get-ADUser @GetADUserParams | Select-Object $ADUserProperties
Write-Log "AD users: $($ADUsers.count)"

#read UJVREZ AD groups
$GetADGroupParams = @{
	Filter = '*'
	SearchBase = $ADGroupSearchBase
	Properties = $ADGroupProperties
	Credential = $ADCredential
}
$ADGroups = Get-ADGroup @GetADGroupParams | Select-Object $ADGroupProperties
Write-Log "AD groups: $($ADGroups.count)"

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
	write-host "." -NoNewline
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

	$adUserParams = @{
		Filter     = "MemberOf -eq '$($group.DistinguishedName)'"
		Properties = $ADUserProperties
		Credential = $ADCredential
	}
	$GroupMembers = Get-ADUser @adUserParams
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
write-host
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
$Date = (Get-Date).ToString("yyMMdd")

$UploadParams.BlobName = "ADUsr_$Date.csv"
$UploadParams.Content = $csvStringUsr
Invoke-AzureBlobUpload @UploadParams

$UploadParams.BlobName = "ADGrpLst_$Date.csv"
$UploadParams.Content = $csvStringGrpLst
Invoke-AzureBlobUpload @UploadParams

$UploadParams.BlobName = "ADGrpMem_$Date.csv"
$UploadParams.Content = $csvStringGrpMem
Invoke-AzureBlobUpload @UploadParams

Write-Log "Script finish"
