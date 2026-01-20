#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "mbxmgmt"
$LogFilePrefix		= "mbxmgmt-full"
$LogFileFreq		= "Y"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Freq $LogFileFreq -Ext "log"

#######################################################################################################################
. $IncFile_StdLogStartBlock

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

$Name_SMTPAuthEnabled = Get-GroupNameFromGraphById -id $GroupId_SMTPAuthEnabled -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
$Name_IMAPEnabled = Get-GroupNameFromGraphById -id $GroupId_IMAPEnabled -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
$Name_POPEnabled = Get-GroupNameFromGraphById -id $GroupId_POPEnabled -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
$Name_LHDisabled = Get-GroupNameFromGraphById -id $GroupId_LHDisabled -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken

Write-Log "Litigation hold duration: $($LHDuration) days"

Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 120


#AuditEnabled false
Write-Log "Processing AuditEnabled"
$AuditDisabledASIS = Get-Mailbox -Filter {AuditEnabled -eq $false} -ResultSize Unlimited
Write-Log "AuditDisabledASIS: $(Get-Count -Object $AuditDisabledASIS)"
foreach ($Mailbox in $AuditDisabledASIS) {
	try {
		Set-Mailbox -Identity $Mailbox.PrimarySmtpAddress -AuditEnabled $true -WarningAction SilentlyContinue -ErrorAction Stop
		Write-Log "audit enabled for $($Mailbox.PrimarySmtpAddress)" -ForceOffScreen
	}
	Catch {
		Write-Host " $($_.Exception.Message)" -ForegroundColor Magenta
	}
}

#Audit actions
Write-Log "Processing Audit actions"
$AllEXOMailboxes = Get-EXOMailbox -Filter {RecipientTypeDetails -ne "DiscoveryMailbox"} -ResultSize Unlimited -Properties AuditEnabled,AuditAdmin,AuditDelegate,AuditOwner

foreach ($EXOMailbox in $AllEXOMailboxes) {
	if ($EXOMailbox.AuditEnabled) {
		$MissingAuditAdminActions = Compare-Object -ReferenceObject $AuditAdminActionsAll -DifferenceObject $EXOMailbox.AuditAdmin -PassThru
		$MissingAuditDelegateActions = Compare-Object -ReferenceObject $AuditDelegateActionsAll -DifferenceObject $EXOMailbox.AuditDelegate -PassThru
		$MissingAuditOwnerActions = Compare-Object -ReferenceObject $AuditOwnerActionsAll -DifferenceObject $EXOMailbox.AuditOwner -PassThru
		
		if ($MissingAuditAdminActions) {
			Write-Log "$($EXOMailbox.userPrincipalName) adding AuditAdmin actions: $($MissingAuditAdminActions)"
			$ActionsToAdd = $MissingAuditAdminActions -split ","
			Try {
				Set-Mailbox -Identity $EXOMailbox.userPrincipalName -AuditAdmin @{Add = $ActionsToAdd}
			}
			Catch {
				Write-Log $_.Exception.Message -MessageType "ERR"
			}
		}
		
		if ($MissingAuditDelegateActions) {
			Write-Log "$($EXOMailbox.userPrincipalName) adding AuditDelegate actions: $($MissingAuditDelegateActions)"
			$ActionsToAdd = $MissingAuditDelegateActions -split ","
			Try {
				Set-Mailbox -Identity $EXOMailbox.userPrincipalName -AuditDelegate @{Add = $ActionsToAdd}
			}
			Catch {
				Write-Log $_.Exception.Message -MessageType "ERR"
			}

		}
		
		if ($MissingAuditOwnerActions) {
			Write-Log "$($EXOMailbox.userPrincipalName) adding AuditOwner actions: $($MissingAuditOwnerActions)"
			$ActionsToAdd = $MissingAuditOwnerActions -split ","
			Try {
				Set-Mailbox -Identity $EXOMailbox.userPrincipalName -AuditOwner @{Add = $ActionsToAdd}
			}
			Catch {
				Write-Log $_.Exception.Message -MessageType "ERR"
			}
		}
	}
}

$AADUsers_DB = Import-CSVtoHashDB -Path $DBFileUsersMemStd -KeyName "id"

Write-Host "Getting all mailboxes...." -NoNewline
$AllMailboxesCAS = Get-CasMailbox -ResultSize Unlimited
Write-Host "done ($($AllMailboxesCAS.Count))"

#POP3
Write-Log "Processing POP3"
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
[array]$Members_POPEnabled = Get-GroupMembersFromGraphById -id $GroupId_POPEnabled -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
Write-Log "Members_POPEnabled: $(Get-Count -Object $Members_POPEnabled)"
$POPEnabledASIS = $AllMailboxesCAS | Where-Object {$_.POPEnabled -eq $true}
Write-Log "POPEnabledASIS: $(Get-Count -Object $POPEnabledASIS.Count)"
foreach ($Mailbox in $POPEnabledASIS) {
	if ($Mailbox.ExternalDirectoryObjectId -notin $Members_POPEnabled.id) {
		Write-log "Disabling POP3 for $($Mailbox.PrimarySmtpAddress)"
		Set-CasMailbox -Identity $Mailbox.PrimarySmtpAddress -PopEnabled $false
	}
}
Write-Log "Processing POP3 exception group $($Name_POPEnabled)"
if ($Members_POPEnabled.Count -gt 0) {
	$Counter = 0
	foreach ($member in $Members_POPEnabled) {
		$CASMailbox = Get-CasMailbox -Identity $member.userPrincipalName
		if ($CASMailbox.POPEnabled -eq $false) {
			$Counter++
			Write-log "Enabling POP3 for $($member.userPrincipalName) per exception group"
			Set-CasMailbox -Identity $member.userPrincipalName -POPEnabled $true
		}
	}
	Write-Log "$($Counter) POP3 exception group members required enabling POP3"
}
else {
	Write-Log "No POP3 exception group members"
}


#IMAP
Write-Log "Processing IMAP"
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
[array]$Members_IMAPEnabled = Get-GroupMembersFromGraphById -id $GroupId_IMAPEnabled -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
Write-Log "Members_IMAPEnabled: $(Get-Count -Object $Members_IMAPEnabled)"
$IMAPEnabledASIS = $AllMailboxesCAS | Where-Object {$_.ImapEnabled -eq $true}
Write-Log "IMAPEnabledASIS: $(Get-Count -Object $IMAPEnabledASIS.Count)"
foreach ($Mailbox in $IMAPEnabledASIS) {
	if ($Mailbox.ExternalDirectoryObjectId -notin $Members_IMAPEnabled.id) {
		Write-Log "Disabling IMAP for $($Mailbox.PrimarySmtpAddress)"
		Set-CasMailbox -Identity $Mailbox.PrimarySmtpAddress -ImapEnabled $false
	}
	
}
Write-Log "Processing IMAP exception group $($Name_IMAPEnabled)"
if ($Members_IMAPEnabled.Count -gt 0) {
	$Counter = 0
	foreach ($member in $Members_IMAPEnabled) {
		$CASMailbox = Get-CasMailbox -Identity $member.userPrincipalName
		if ($CASMailbox.ImapEnabled -eq $false) {
			$Counter++
			Write-Log "Enabling IMAP for $($member.userPrincipalName) per exception group"
			Set-CasMailbox -Identity $member.userPrincipalName -ImapEnabled $true
		}
	}
	Write-Log "$($Counter) IMAP exception group members required enabling IMAP"
}
else {
	Write-Log "No IMAP exception group members"
}


#SMTPAuth
Write-Log "Processing SMTPAuth"
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
[array]$Members_SMTPAuthEnabled = Get-GroupMembersFromGraphById -id $GroupId_SMTPAuthEnabled -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
Write-Log "Members_SMTPAuthEnabled: $(Get-Count -Object $Members_SMTPAuthEnabled)"
$SMTPAuthEnabledASIS = $AllMailboxesCAS | Where-Object {$_.SmtpClientAuthenticationDisabled -eq $false}
Write-Log "SMTPAuthEnabledASIS: $(Get-Count -Object $SMTPAuthEnabledASIS.Count)"
foreach ($Mailbox in $SMTPAuthEnabledASIS) {
	if ($Mailbox.ExternalDirectoryObjectId -notin $Members_SMTPAuthEnabled.id) {
		Write-Log "Disabling SMTP Auth for $($Mailbox.PrimarySmtpAddress)"
		Set-CasMailbox -Identity $Mailbox.PrimarySmtpAddress -SmtpClientAuthenticationDisabled $true
	}
}
if ($Members_SMTPAuthEnabled.Count -gt 0) {
	$Counter = 0
	foreach ($member in $Members_SMTPAuthEnabled) {
		$CASMailbox = Get-CasMailbox -Identity $member.userPrincipalName
		if ($CASMailbox.SmtpClientAuthenticationDisabled -eq $true) {
			$Counter++
			Write-Log "Enabling SMTP Auth for $($Mailbox.PrimarySmtpAddress) per exception group"
			Set-CASMailbox -Identity $member.userPrincipalName -SmtpClientAuthenticationDisabled $false
		}
	}
	Write-Log "$($Counter) SMTPAuth exception group members required enabling SMTPAuth"
}
else {
	Write-Log "No SMTPAuth exception group members"
}


#Litigation hold
Write-Log "Processing litigation hold"
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 120
[array]$Members_LHDisabled = Get-GroupMembersFromGraphById -id $GroupId_LHDisabled -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
Write-Log "Members_LHDisabled: $(Get-Count -Object $Members_LHDisabled)"
$LHDisabledASIS = Get-Mailbox -Filter {LitigationHoldEnabled -eq $false} -ResultSize Unlimited
Write-Log "LHDisabledASIS: $(Get-Count -Object $LHDisabledASIS.Count)"
foreach ($Mailbox in $LHDisabledASIS) {
	if ($Mailbox.ExternalDirectoryObjectId -notin $Members_LHDisabled.id) {
		if ($Mailbox.ExternalDirectoryObjectId -and $AADUsers_DB.ContainsKey($Mailbox.ExternalDirectoryObjectId)) {
			$AADUser = $AADUsers_DB[$Mailbox.ExternalDirectoryObjectId]
		}
		else {	
			#write-host "not in AADUsers_DB"
			Continue
		}

		if ($Mailbox.RecipientTypeDetails -ne "UserMailbox") {
			#write-host "not UserMailbox"
			Continue
		}

		if (-not ($AADUser.department -and $AADUser.companyName -and $AADUser.onPremisesSyncEnabled)) {
			#write-host "missing department, companyName or onPremisesSyncEnabled"
			Continue
		}

		$onpremSAM = $AADUser.onPremisesSamAccountName
		if 	(($onpremSAM -like "q*") -and ($onpremSAM -notlike "qt*") -and ($onpremSAM -notlike "qk*")) {
			#write-host "onpremSAM not qt* or qk*"
			Continue
		}

		if ($AADUser.onPremisesDistinguishedName -like "*OU=Groupware,OU=aplikacni,OU=uzivatele,DC=cezdata,DC=corp") {
			#write-host "onPremisesDistinguishedName OU=Groupware,OU=aplikacni,OU=uzivatele,DC=cezdata,DC=corp"
			Continue
		}

		Try {
			Set-Mailbox -Identity $Mailbox.PrimarySmtpAddress -LitigationHoldEnabled $true -LitigationHoldDuration $LHDuration -WarningAction SilentlyContinue -ErrorAction Stop
			Write-Log "litigation hold enabled for $($Mailbox.PrimarySmtpAddress)"
		}
		Catch {
			if ($_.Exception.Message.Contains("Exchange Online license doesn't permit")) {
				#Write-Host "ERROR - NO LICENSE FOR LH" -ForegroundColor Red
			}
			Else {
				Write-Host "  $($_.Exception.Message)" -ForegroundColor Magenta
			}
		}
	}
	else {
		#write-host "has exception from LH" -ForegroundColor Cyan
	}
	
}
if ($Members_LHDisabled.Count -gt 0) {
	$Counter = 0
	foreach ($member in $Members_LHDisabled) {
		$Mailbox = Get-Mailbox -Identity $member.userPrincipalName
		if ($Mailbox.LitigationHoldEnabled -eq $true) {
			$Counter++
			Write-Log "disabling litigation hold for $($member.userPrincipalName) per exception group"
			Set-Mailbox -Identity $member.userPrincipalName -LitigationHoldEnabled $false
		}
	}
	Write-Log "$($Counter) LH exception group members required disabling LH"
}
else {
	Write-Log "No LH exception group members"
}

#Litigation hold duration
Write-Log "Processing litigation hold duration"
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$AllEXOMailboxes = Get-EXOMailbox -Properties LitigationHoldEnabled,LitigationHoldDuration -ResultSize Unlimited
foreach ($Mailbox in $AllEXOMailboxes) {
	if ($Mailbox.LitigationHoldEnabled -eq $true) {
		try {
			$CurrentLHDuration = [int]$Mailbox.LitigationHoldDuration.substring(0,$Mailbox.LitigationHoldDuration.IndexOf("."))
		}
		Catch {
			$CurrentLHDuration = -1
		}
		if ($CurrentLHDuration -ne $LHDuration) {
			Try {
				Set-Mailbox -Identity $Mailbox.PrimarySmtpAddress -LitigationHoldDuration $LHDuration -WarningAction SilentlyContinue -ErrorAction Stop
				Write-Log "litigation hold duration set to $LHDuration for $($Mailbox.PrimarySmtpAddress)"
			}
			Catch {
				if ($_.Exception.Message.Contains("Exchange Online license doesn't permit")) {
					#Write-Host "ERROR - NO LICENSE FOR LH" -ForegroundColor Red
				}
				Else {
					Write-Host $_.Exception.Message -ForegroundColor Red
				}
			}
		}
	}
}

Get-PSSession | Remove-PSSession

#######################################################################################################################

. $IncFile_StdLogEndBlock
