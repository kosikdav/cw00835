#######################################################################################################################
# Set-MaiboxT2T-Properties
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "t2t"
$LogFilePrefix		= "mbx-properties"
$LogFileFreq		= "Y"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Freq $LogFileFreq -Ext "log"

[datetime]$date = (get-date).AddDays(-$DaysBack)

$MbxFilter1 = "(alias -like '*') "
$MbxFilter2 = "-and ((RecipientTypeDetails -eq 'UserMailbox') -or (RecipientTypeDetails -eq 'SharedMailbox') -or (RecipientTypeDetails -eq 'RoomMailbox') -or (RecipientTypeDetails -eq 'EquipmentMailbox')) "
$MbxFilter3 = "-and (WhenMailboxCreated -gt '$($date)')"
$userMbxFilter = $MbxFilter1 + $MbxFilter2 + $MbxFilter3

[array]$userMbxSet = @()
$DB_changed = $false
$ToBeDeletedRecords = @()

function Get-TargetT2TUser {
	[CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$SrcIdentity,
        [Parameter(Mandatory)][string]$SrcAccessToken,
		[Parameter(Mandatory)][string]$DstAccessToken
    )
	# main function body ##################################
	if ($SrcIdentity -eq "6c38bc03-f798-454d-9ef8-ee18cab92105") {
		$User = [pscustomobject]@{
			id = "d352c2a8-eb08-4e57-b71d-4a107e080b4e"
			userPrincipalName = "qxsvecrad2@cez.cz"
			onpremisesSamAccountName = "qxsvecrad2"
		}
		return $User
	}
}
#######################################################################################################################

. $IncFile_StdLogStartBlock

$DstMailRoutingDomain = "cezdata.mail.onmicrosoft.com"

$Src_AppReg_LOG_READER 			= $AppReg_UJVREZ_LOG_READER
$Src_AppReg_EXO_MGMT 			= $AppReg_UJVREZ_EXO_MGMT   
$Src_T2T_EXO_MIGRATION_GROUP 	= $UJVREZ_T2T_EXO_MIGRATION_GROUP

$Dst_AppReg_LOG_READER 			= $AppReg_CEZDATA_LOG_READER
$Dst_AppReg_EXO_MGMT 			= $AppReg_CEZDATA_EXO_MGMT   

Request-MSALToken -AppRegName $Src_AppReg_LOG_READER -TTL 30

#get source T2T migration group
$T2TMigrationGroupName = Get-GroupNameFromGraphById -AccessToken $AuthDB[$Src_AppReg_LOG_READER].AccessToken -Id $Src_T2T_EXO_MIGRATION_GROUP

if ($T2TMigrationGroupName) {
	Write-Host "Source T2T migration group found: $T2TMigrationGroupName " -NoNewline
} else {
	Write-Host "Source T2T migration group with ID $Src_T2T_EXO_MIGRATION_GROUP not found."
	Exit
}

$UriResource = "groups/$($Src_T2T_EXO_MIGRATION_GROUP)/members"
$Uri = New-GraphUri -Resource $UriResource -Version "v1.0"
[array]$T2TMigrationGroupMembers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$Src_AppReg_LOG_READER].AccessToken
write-host "SRC T2T migration group members: $($T2TMigrationGroupMembers.count)"

Connect-EXOService -AppRegName $Src_AppReg_EXO_MGMT  -TTL 120
write-host "SRC mailboxes total: " -NoNewline
[array]$SrcMailboxesAll = Get-EXOMailbox -ResultSize Unlimited -PropertySets All
write-host $SrcMailboxesAll.count
[array]$SrcMailboxes = $SrcMailboxesAll | Where-Object { $_.id -in $T2TMigrationGroupMembers.id }
write-host "SRC mailboxes in T2T migration group: $($SrcMailboxes.count)"
Remove-Variable SrcMailboxesAll
Get-PSSession | Remove-PSSession

Request-MSALToken -AppRegName $Dst_AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriSelect = "id,userPrincipalName,onpremisesSamAccountName"
$UriFilter = "userType eq 'Member' and onpremisesSyncEnabled eq true"
$Uri = New-GraphUri -Resource $UriResource -Version "v1.0" -Select $UriSelect -Filter $UriFilter -Top 999
[array]$DstAADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$Dst_AppReg_LOG_READER].AccessToken -ProgressDots -Text "AAD users"
write-host "DST AAD users: $($DstAADUsers.count)"

Connect-EXOService -AppRegName $Dst_AppReg_EXO_MGMT -TTL 120
$DstEXOMailUsers = Get-EXORecipient -ResultSize Unlimited -PropertySets "Minimum, MailboxMove" -RecipientType MailUser
write-host "DST EXO mailUsers: $($DstEXOMailUsers.count)"
Get-PSSession | Remove-PSSession

$Session = New-PSSession -Name "OnPremExchange" -ConfigurationName "Microsoft.Exchange" -ConnectionUri "http://cw00616exch3.cezdata.corp/PowerShell/" -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking -AllowClobber | Out-Null
write-host "DST AD mailUsers: " -NoNewline
$DstADMailUsers = Get-MailUser -ResultSize Unlimited
write-host $DstADMailUsers.count

foreach ($SrcMailbox in $SrcMailboxes) {
	write-host "------------------------"
	[array]$SRCsmtpAddresses = [array]$SRCx500Addresses = [array]$SRCproxyAddresses = @()
	$ArchiveGuid = $null
	
	if ($SrcMailbox.EmailAddresses) {
        foreach ($EmailAddress in $SrcMailbox.EmailAddresses) {
            if ($EmailAddress -like "smtp:*" -and $EmailAddress -notlike "*onmicrosoft.com") {
                $SRCsmtpAddresses += $EmailAddress
            }
            if ($EmailAddress -like "X500:*") {
                $SRCx500Addresses += $EmailAddress
            }
        }
        $SRCproxyAddresses = $SrcMailbox.EmailAddresses -join ";"
    }
    if ($SrcMailbox.ArchiveGuid -and $SrcMailbox.ArchiveGuid -ne "00000000-0000-0000-0000-000000000000") {
        $ArchiveGuid = $SrcMailbox.ArchiveGuid
    }

	Write-Host "Processing mailbox: $($SrcMailbox.UserPrincipalName)"
	write-host "PrimarySmtpAddress: $($SrcMailbox.PrimarySmtpAddress)"
	write-host "ExchangeObjectId: $($SrcMailbox.ExchangeObjectId)"
	write-host "ExchangeGuid: $($SrcMailbox.ExchangeGuid)"
	write-host "ArchiveStatus: $($SrcMailbox.ArchiveStatus)"
	write-host "ArchiveGuid: $($SrcMailbox.ArchiveGuid)"
	write-host "LegacyExchangeDN: $($SrcMailbox.LegacyExchangeDN)"
	write-host $SRCx500Addresses -ForegroundColor Cyan
	write-host $SRCsmtpAddresses -ForegroundColor Green
	$DstUser = Get-TargetT2TUser -SrcIdentity $SrcMailbox.id -SrcAccessToken $AuthDB[$Src_AppReg_LOG_READER].AccessToken -DstAccessToken $AuthDB[$Src_AppReg_EXO_MGMT].AccessToken
	if ($DstUser) {
		write-host "Target user: $($DstUser.userPrincipalName)"	
		Try {
			$DstMailUser = Get-MailUser -Identity $DstUser.onpremisesSamAccountName -ErrorAction Stop
			write-host "Target user $($DstUser.onpremisesSamAccountName) is type mailUser. Updating properties..."
			write-host $DstMailuser
			
			write-host "Setting PrimarySmtpAddress to $($DstUser.userPrincipalName)"
			Set-MailUser -Identity $DstUser.onpremisesSamAccountName -PrimarySmtpAddress $DstUser.userPrincipalName
			
			write-host "Setting ExternalEmailAddress to $($SrcMailbox.PrimarySmtpAddress)"
			Set-MailUser -Identity $DstUser.onpremisesSamAccountName -ExternalEmailAddress $SrcMailbox.PrimarySmtpAddress
			
			write-host "Adding smtp address $($DstUser.onpremisesSamAccountName)@$DstMailRoutingDomain"
			Set-MailUser -Identity $DstUser.onpremisesSamAccountName -EmailAddresses @{add="smtp:$($DstUser.onpremisesSamAccountName)@$DstMailRoutingDomain"}

			write-host "Adding LegacyExchangeDN $($SrcMailbox.LegacyExchangeDN)"
			Set-MailUser -Identity $DstUser.onpremisesSamAccountName -EmailAddresses @{add="x500:$($SrcMailbox.LegacyExchangeDN)"}

			write-host "Adding x500 addresses $($SRCx500Addresses -join ",")"
			foreach ($x500 in $SRCx500Addresses) {
				Set-MailUser -Identity $DstUser.onpremisesSamAccountName -EmailAddresses @{add="$x500"}
			}
		}
		Catch {
			write-host "Target user: $($DstUser.onpremisesSamAccountName) is not type mailUser"
			Enable-MailUser -Identity $DstUser.onpremisesSamAccountName -PrimarySmtpAddress $DstUser.userPrincipalName -Alias $DstUser.onpremisesSamAccountName -ExternalEmailAddress $SrcMailbox.PrimarySmtpAddress
		}
	}
	Else {
		write-host "No target user found for source mailbox $($SrcMailbox.UserPrincipalName) with id $($SrcMailbox.id). Skipping..." -ForegroundColor Yellow
	}
}

#######################################################################################################################

. $IncFile_StdLogEndBlock
