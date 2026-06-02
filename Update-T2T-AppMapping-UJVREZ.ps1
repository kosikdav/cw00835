#######################################################################################################################
# Update-T2T-AppMapping-UJVREZ.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
	[switch]$TestRun
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1
. $ScriptPath\include-Script-Start-Include.ps1

$LogFolder				= "t2t-ujvrez"
$LogFilePrefix			= "app-mapping"
$OutputFolder			= "t2t-ujvrez"
$OutputFilePrefix		= "app-mapping"
$OutputFileSuffixSRC 	= "src"

$MappingCSV_Apps_FilePath 		= "d:\data\t2t-ujvrez\appMapping.csv"

$Src_mapfile_attr = "UJV_sam"
$Dst_mapfile_attr = "CEZ_sam"
$Src_PN_attr = "extension_93b54ce056df45bd8f5f398753fa17c0_employeeNumber"
$Dst_PN_attr = "employeeNumber"
$Dst_mailAD40 = "msExchExtensionAttribute40"

$DstADUserProperties = @(
	'displayName',
	'mail',
	'distinguishedName',
	'samAccountName',
	'userPrincipalName',
	'extensionAttribute10'
)

$DstADSearchBaseAll  = "OU=uzivatele,DC=cezdata,DC=corp"
$DstADSearchBaseSKC  = "OU=UJV,OU=aplikacni,OU=uzivatele,DC=cezdata,DC=corp"

$Src_AppReg_LOG_READER 			= $AppReg_UJVREZ_LOG_READER
$Src_AppReg_EXO_MGMT 			= $AppReg_UJVREZ_EXO_MGMT   
$Src_T2T_EXO_MIGRATION_GROUP 	= $UJVREZ_T2T_EXO_MIGRATION_GROUP

$Dst_AppReg_LOG_READER 			= $AppReg_CEZDATA_LOG_READER
$Dst_AppReg_EXO_MGMT 			= $AppReg_CEZDATA_EXO_MGMT   

$CommonEntraAttributes = "id,userPrincipalName,displayName,onPremisesSamAccountName,mail"

[array]$MappingReportSRC = @()
[array]$MappingReportDST = @()
[array]$SGUMReport = @()

[hashtable]$mapping_DB = @{}
[hashtable]$Dst_UserDB_per_pn = @{}
[array]$DSTUsersMapped = @()

#######################################################################################################################

$LogFile 	= New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$OutputFileSRC	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixSRC -Ext "csv" -Freq "YMDHMS"
$OutputFileDST	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixDST -Ext "csv" -Freq "YMDHMS"

#######################################################################################################################

if ($InteractiveRun) {
	$ADCredentialPath = "c:\cred\qp_aad_grp_mgmt\qp_aad_grp_mgmt_qskosikdav.cred"
}
else {
	$ADCredentialPath = $aad_grp_mgmt_cred
}

$ADCredential = Import-Clixml -Path $ADCredentialPath

$DstGetADUserParams = @{
	Filter = "ObjectClass -eq 'user'"
	SearchBase = $DstADSearchBaseSKC
	Properties = $DstADUserProperties
	Credential = $ADCredential
}

#######################################################################################################################

. $IncFile_StdLogStartBlock

Write-Log "DstADSearchBaseAll: $DstADSearchBaseAll"
Write-Log "DstADSearchBaseSKC: $DstADSearchBaseSKC"
Write-Log "Src_mapfile_attr: $Src_mapfile_attr"
Write-Log "Dst_mapfile_attr: $Dst_mapfile_attr"
write-log $string_divider

write-log "SRC MAPPING" -ForegroundColor Cyan
Request-MSALToken -AppRegName $Src_AppReg_LOG_READER -TTL 30

#######################################################################################################################
# mapping file
#######################################################################################################################

############################
#apps
[array]$appMapping = Import-CSVtoArray -Path $MappingCSV_Apps_FilePath
write-host "App mailbox mapping: $($appMapping.count)"

write-host "Checking duplicate $($Src_mapfile_attr) in mapping file..." -NoNewline
$duplicateSrcMappings = $appMapping | Group-Object -Property $Src_mapfile_attr | Where-Object { $_.Count -gt 1 }
foreach ($group in $duplicateSrcMappings) {
	write-host "Duplicate $($Src_mapfile_attr): $($group.Name) - Count: $($group.Count)"
	foreach ($mapping in $group.Group) {
		write-host $mapping
	}
	exit
}
write-host "done"

write-host "Checking duplicate $($Dst_mapfile_attr) in mapping file..." -NoNewline
$duplicateDstMappings = $appMapping | Group-Object -Property $Dst_mapfile_attr | Where-Object { $_.Count -gt 1 }
foreach ($group in $duplicateDstMappings) {
	write-host "Duplicate $($Dst_mapfile_attr): $($group.Name) - Count: $($group.Count)"
	foreach ($mapping in $group.Group) {
		write-host $mapping
	}
	exit
}
write-host "done"

$SrcAppUsersWithMapping = $appMapping.$Src_mapfile_attr
$DstAppUsersWithMapping = $appMapping.$Dst_mapfile_attr

foreach ($mapping in $appMapping) {
	$mapping_DB.add($mapping.$Src_mapfile_attr, $mapping.$Dst_mapfile_attr)
}

#######################################################################################################################
# get SRC AAD users
#######################################################################################################################
$UriResource = "users"
$UriSelect = "$CommonEntraAttributes,assignedLicenses"
$UriFilter = "userType eq 'Member'"
$Uri = New-GraphUri -Resource $UriResource -Version "v1.0" -Select $UriSelect -Filter $UriFilter -Top 999
[array]$SrcAADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$Src_AppReg_LOG_READER].AccessToken -ProgressDots -Text "SRC AAD users"

$SrcAADUsers = $SrcAADUsers | Where-Object { $_.onpremisesSamAccountName -in $SrcAppUsersWithMapping }
write-host "SRC AAD users - filtered users without mapping: $($SrcAADUsers.count)"

#######################################################################################################################
# get DST AD users
#######################################################################################################################

#read CEZDATA AD users and filter only those with enabled account
write-host "DST AD users - reading from AD (long running operation)..." -NoNewline
$DstADUsers = Get-ADUser @DstGetADUserParams | Select-Object $DstADUserProperties
write-host "done ($($DstADUsers.count))"

#check duplicates in ext10 before proceeding
write-host "DST AD users - checking duplicate ext10..." -NoNewline
$DstADUsersExt10 = $DstADUsers | Where-Object { $_.extensionAttribute10 -ne $null -and $_.extensionAttribute10 -ne "" -and $_.extensionAttribute10 -notlike "_*" }
$duplicateUsers = $DstADUsersExt10 | Group-Object -Property "extensionAttribute10" | Where-Object { $_.Count -gt 1 }
foreach ($group in $duplicateUsers) {
	write-host "Duplicate ext10: $($group.Name) - Count: $($group.Count)"
	foreach ($user in $group.Group) {
		write-host "  User: $($user.displayName) - UPN: $($user.userPrincipalName)" -ForegroundColor Red
	}
	exit
}
write-host "done"

#OU property to each user object by parsing it from DistinguishedName, we will need it for filtering users by OU and for reporting
write-host "DST AD users - adding OU property..." -NoNewline
foreach ($user in $DstADUsers) {    
	$user | Add-Member -NotePropertyName 'OU' -NotePropertyValue ( $user.DistinguishedName -replace '^CN=[^,]+,' )
}
write-host "done"

write-host $string_divider

$countSRCOK = 0
$countSRCUpdated = 0
$countSRCNotFound = 0
$countSRCTotal = 0

#######################################################################################################################
# SRC mapping
#######################################################################################################################
foreach ($user in $SrcAADUsers) {
	$ReportObject = $null

		$countSRCTotal++
		if ($user.mail) {
			$UJV_mailDomain = $user.mail.Split("@")[1]
		}
		else {
			$UJV_mailDomain = [string]::Empty
		}
		$ReportObject = [PSCustomObject]@{
			Result = $null
			UJV_UPN = $user.userPrincipalName
			UJV_UPNdomain = $user.userPrincipalName.Split("@")[1]
			UJV_DisplayName = $user.displayName
			UJV_mail = $user.mail
			UJV_mailDomain = $UJV_mailDomain
			UJV_samAccountName = $user.onpremisesSamAccountName
			Mapped_SAM = $null
			CEZ_UPN = $null
			CEZ_DisplayName = $null
			CEZ_mail = $null
			CEZ_samAccountName = $null
			OldExt10 = $null
			NewExt10 = $null
		}

		$mapped_SAM = $mapping_DB[$user.onpremisesSamAccountName]
		$ReportObject.Mapped_SAM = $mapped_SAM
		$DstUser = Get-ADUser -Identity $mapped_SAM -Properties $DstADUserProperties -Credential $ADCredential

		if ($DstUser) {
			$ReportObject.CEZ_UPN = $DstUser.userPrincipalName
			$ReportObject.CEZ_DisplayName = $DstUser.displayName
			$ReportObject.CEZ_mail = $DstUser.mailAD40
			$ReportObject.CEZ_samAccountName = $DstUser.SamAccountName

			if ($DstUser.extensionAttribute10 -ne $user.onpremisesSamAccountName) {#
				if ($DstUser.extensionAttribute10) {
					$currentExt10 = $DstUser.extensionAttribute10
				}
				else {
					$currentExt10 = "<empty>"
				}
				try {
					if (-not $TestRun) {
						Set-ADUser -Identity $DstUser.samAccountName -Replace @{extensionAttribute10 = $user.onpremisesSamAccountName} -Credential $ADCredential
					}
					$countSRCUpdated++
					Write-Log "UPDATING:  $($user.userPrincipalName) ($($user.displayName)) PN:$($user.onpremisesSamAccountName) - $($DstUser.userPrincipalName) $($DstUser.SamAccountName) ext10:$($currentExt10) -> $($user.onpremisesSamAccountName)" -ForegroundColor Yellow
					$ReportObject.OldExt10 = $currentExt10
					$ReportObject.NewExt10 = $user.onpremisesSamAccountName
					$ReportObject.Result = "UPDATED"
				}
				catch {
					Write-Log "Error updating user: $($user.userPrincipalName) ($($user.displayName)) - $($_.Exception.Message)" -ForegroundColor Red -MessageType "ERR"
				}
			}
			else {
				#write-Log "OK:        $($user.userPrincipalName) ($($user.displayName)) UJV_PN:$($user.onpremisesSamAccountName) - $($DstUser.userPrincipalName) $($DstUser.SamAccountName) ext10:$($DstUser.ext10)" -ForegroundColor Green
				$ReportObject.OldExt10 = $currentExt10
				$ReportObject.NewExt10 = $currentExt10
				$ReportObject.Result = "OK"
				$countSRCOK++
			}
		}
		else {
			Write-Log "NOTFOUND:  $($user.userPrincipalName) ($($user.displayName)) UJV_PN:$($user.onpremisesSamAccountName) - expected CEZ_PN: $($mapped_pn)" -ForegroundColor Red -MessageType "ERR"
			$ReportObject.Result = "NOTFOUND"
			$countSRCNotFound++
		}
	if ($ReportObject) {
		$MappingReportSRC += $ReportObject
	}
}

write-host $string_divider

#######################################################################################################################

Write-Host
Write-Log "SRC mapping summary:" -ForegroundColor Cyan
Write-Log "----------------------------" -ForegroundColor Cyan
Write-Log "Total:       $countSRCTotal"
Write-Log "Updated:     $($countSRCUpdated)" -foregroundcolor yellow
Write-Log "UpdatedSec:  $($countSRCUpdatedSecondary)" -foregroundcolor yellow
Write-Host
Write-Log "OK:          $($countSRCOK) ($(($countSRCOK/$countSRCTotal*100).ToString("##.##"))%)" -foregroundcolor Green
Write-Log "NotFound:    $($countSRCNotFound) ($(($countSRCNotFound/$countSRCTotal*100).ToString("##.##"))%)" -foregroundcolor red
Write-Log "NoMapping:   $($countSRCNoMapping) ($(($countSRCNoMapping/$countSRCTotal*100).ToString("##.##"))%)" -foregroundcolor darkyellow
Write-Log "MailErr:     $($countSRCMailErr) ($(($countSRCMailErr/$countSRCTotal*100).ToString("##.##"))%)" -foregroundcolor darkcyan
Write-Host

Export-Report -Text "SRC - T2T user mapping report UJVREZ-CEZDATA" -Report $MappingReportSRC -Path $OutputFileSRC -SortProperty "UJV_UPN"

#######################################################################################################################

. $IncFile_StdLogEndBlock