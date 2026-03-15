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


$Src_AppReg_LOG_READER 			= $AppReg_UJVREZ_LOG_READER
$Src_AppReg_EXO_MGMT 			= $AppReg_UJVREZ_EXO_MGMT   
$Src_T2T_EXO_MIGRATION_GROUP 	= $UJVREZ_T2T_EXO_MIGRATION_GROUP

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
write-host "members: $($T2TMigrationGroupMembers.count)"

Connect-EXOService -AppRegName $Src_AppReg_EXO_MGMT  -TTL 120
[array]$SrcMailboxesAll = Get-EXOMailbox -ResultSize Unlimited -PropertySets All
$SrcMailboxes = $SrcMailboxesAll | Where-Object { $_.id -in $T2TMigrationGroupMembers.id }
Remove-Variable SrcMailboxesAll

$Session = New-PSSession -Name "OnPremExchange" -ConfigurationName "Microsoft.Exchange" -ConnectionUri "http://cw00616exch3.cezdata.corp/PowerShell/" -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking -AllowClobber | Out-Null

foreach ($SrcMailbox in $SrcMailboxes) {
	write-host "------------------------"
	$ArchiveGuid = $x500Addresses = $null
    $smtpAddresses = $x500Addresses = $proxyAddresses = $null

	if ($mailbox.EmailAddresses) {
        foreach ($EmailAddress in $mailbox.EmailAddresses) {
            if ($EmailAddress -like "smtp:*" -and $EmailAddress -notlike "*onmicrosoft.com") {
                $smtpAddresses += $EmailAddress
            }
            if ($EmailAddress -like "X500:*") {
                $x500Addresses += $EmailAddress
            }
        }
        $proxyAddresses = $mailbox.EmailAddresses -join ";"
    }
    if ($mailbox.ArchiveGuid -and $mailbox.ArchiveGuid -ne "00000000-0000-0000-0000-000000000000") {
        $ArchiveGuid = $mailbox.ArchiveGuid
    }

	Write-Host "Processing mailbox: $($SrcMailbox.UserPrincipalName)"
	write-host $SrcMailbox.PrimarySmtpAddress
	write-host $SrcMailbox.ExchangeObjectId
	write-host $SrcMailbox.ExchangeGuid
	write-host $SrcMailbox.ArchiveGuid
	write-host $SrcMailbox.LegacyExchangeDN
	$DstUser = Get-TargetT2TUser -SrcIdentity $SrcMailbox.id -SrcAccessToken $AuthDB[$Src_AppReg_LOG_READER].AccessToken -DstAccessToken $AuthDB[$Src_AppReg_EXO_MGMT].AccessToken
	if ($DstUser) {
		write-host "Target user: $($DstUser.userPrincipalName)"	
		Try {
			$DstMailUser = Get-MailUser -Identity $DstUser.userPrincipalName
		}
		Catch {
			write-host "Target user: $($DstUser.userPrincipalName) is not type mailUser"
			Enable-MailUser -Identity $DstUser.userPrincipalName -PrimarySmtpAddress $DstUser.userPrincipalName -Alias $DstUser.onpremisesSamAccountName -ExternalEmailAddress $SrcMailbox.PrimarySmtpAddress
		}
	}
	Else {
		write-host "No target user found for source mailbox $($SrcMailbox.UserPrincipalName) with id $($SrcMailbox.id). Skipping..." -ForegroundColor Yellow
	}
}

#######################################################################################################################

. $IncFile_StdLogEndBlock
