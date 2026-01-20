#######################################################################################################################
# Set-MaiboxProperties
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "mbxmgmt"
$LogFilePrefix		= "mbxmgmt"
$LogFileFreq		= "Y"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Freq $LogFileFreq -Ext "log"

$DaysBack = 30

[datetime]$date = (get-date).AddDays(-$DaysBack)

$MbxFilter1 = "(alias -like '*') "
$MbxFilter2 = "-and ((RecipientTypeDetails -eq 'UserMailbox') -or (RecipientTypeDetails -eq 'SharedMailbox') -or (RecipientTypeDetails -eq 'RoomMailbox') -or (RecipientTypeDetails -eq 'EquipmentMailbox')) "
$MbxFilter3 = "-and (WhenMailboxCreated -gt '$($date)')"
$userMbxFilter = $MbxFilter1 + $MbxFilter2 + $MbxFilter3

[array]$userMbxSet = @()
$DB_changed = $false
$ToBeDeletedRecords = @()

#######################################################################################################################

. $IncFile_StdLogStartBlock

Write-Log "Litigation hold duration: $($LHDuration) days"
Write-Log "Cutoff date: $($date)"
Write-Log "Mailbox filter: $($userMbxFilter)" -ForceOffScreen

# load DB mailbox-mgmt-db from file or initialize empty
if (test-path $DBFileEXOMboxMgmt) {
    Try {
        $EXOMboxMgmt_DB = Import-Clixml -Path $DBFileEXOMboxMgmt
        Write-Log "DB file $($DBFileEXOMboxMgmt) imported successfully, $($EXOMboxMgmt_DB.count) records found"
    } 
    Catch {
        Write-Log "Error importing $($DBFileEXOMboxMgmt), creating empty DB" -MessageType "Error"
        [hashtable]$EXOMboxMgmt_DB = @{}
        $DB_changed = $true
    }
}
else {
    Write-Log "DB file $($DBFileEXOMboxMgmt) not found, creating empty DB" -MessageType "Error"
    [hashtable]$EXOMboxMgmt_DB = @{}
    $DB_changed = $true
}

Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 120
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

write-log "Members_SMTPAuthEnabled".PadRight(30,".") -NoNewline
[array]$Members_SMTPAuthEnabled = Get-GroupMembersFromGraphById -id $GroupId_SMTPAuthEnabled -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
$Name_SMTPAuthEnabled = Get-GroupNameFromGraphById -id $GroupId_SMTPAuthEnabled -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
write-log "($($Members_SMTPAuthEnabled.Count))" -ForegroundColor Cyan -NoLinePrefix
foreach ($member in $Members_SMTPAuthEnabled) {
	Set-CasMailbox -Identity $member.userPrincipalName -SmtpClientAuthenticationDisabled $false
}

write-log "Members_IMAPEnabled".PadRight(30,".") -NoNewline
[array]$Members_IMAPEnabled = Get-GroupMembersFromGraphById -id $GroupId_IMAPEnabled -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
$Name_IMAPEnabled = Get-GroupNameFromGraphById -id $GroupId_IMAPEnabled -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
write-log "($($Members_IMAPEnabled.Count))" -ForegroundColor Cyan -NoLinePrefix
foreach ($member in $Members_IMAPEnabled) {
	Set-CasMailbox -Identity $member.userPrincipalName -ImapEnabled $true
}

write-log "Members_POPEnabled".PadRight(30,".") -NoNewline
[array]$Members_POPEnabled = Get-GroupMembersFromGraphById -id $GroupId_POPEnabled -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
$Name_POPEnabled = Get-GroupNameFromGraphById -id $GroupId_POPEnabled -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
write-log "($($Members_POPEnabled.Count))" -ForegroundColor Cyan -NoLinePrefix
if ($Members_POPEnabled.Count -gt 0) {
	foreach ($member in $Members_POPEnabled) {
		Set-CasMailbox -Identity $member.userPrincipalName -POPEnabled $true
	}
}

write-log "Members_LHDisabled".PadRight(30,".") -NoNewline
[array]$Members_LHDisabled = Get-GroupMembersFromGraphById -id $GroupId_LHDisabled -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
$Name_LHDisabled = Get-GroupNameFromGraphById -id $GroupId_LHDisabled -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
write-log "($($Members_LHDisabled.Count))" -ForegroundColor Cyan -NoLinePrefix
if ($Members_LHDisabled.Count -gt 0) {
	foreach ($member in $Members_LHDisabled) {
		Set-Mailbox -Identity $member.userPrincipalName -LitigationHoldEnabled $false -WarningAction SilentlyContinue
	}
}

#retrieve members of O365Poradna (d9e1b8c5-3dab-416e-9e3d-c332ce177c7b)
write-log "Members_O365Poradna".PadRight(30,".") -NoNewline
[array]$Members_O365Poradna = Get-GroupMembersFromGraphById -id $GroupId_O365Poradna -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
$Name_O365Poradna = Get-GroupNameFromGraphById -id $GroupId_O365Poradna -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
write-log "($($Members_O365Poradna.Count))" -ForegroundColor Cyan -NoLinePrefix

#retrieve mailboxes 
Write-Log "Looking for user mailboxes created after $($date)..." -NoNewline
[array]$userMbxSet = Get-Mailbox -ResultSize Unlimited -Filter $userMbxFilter
Write-Log "($($userMbxSet.Count))" -ForegroundColor Cyan -NoLinePrefix

$IgnoredMailboxesCount = 0
$ProcessedMailboxesCount = 0

If ($userMbxSet.Count -gt 0) {

	ForEach ($userMbx in $userMbxSet) {
		if ($EXOMboxMgmt_DB.ContainsKey($userMbx.ExchangeGuid)) {
			Write-Host "skipping $($userMbx.userPrincipalName)" -ForegroundColor Yellow -BackgroundColor DarkGray
			$IgnoredMailboxesCount++
			continue
		}
		$MbxProcessedCorrectly = $true
		Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
		Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

		#$WarningPreference = 'SilentlyContinue'
		$upn,$rtd,$als,$whnCre,$onpremSAM,$onpremDN = $null
		$isStdUser = $false

		$rtd 		= $userMbx.RecipientTypeDetails
		$upn 		= $userMbx.userprincipalname
		$als 		= $userMbx.alias
		$whnCre 	= $userMbx.WhenMailboxCreated
		$Properties = "id,userprincipalname,displayName,department,companyName,onPremisesSyncEnabled,onPremisesSamAccountName,onPremisesDistinguishedName"
		$User = Get-UserFromGraphById -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -Id $upn -Properties $Properties -Version "v1.0"
		$onpremSAM = $user.onPremisesSamAccountName
		$onpremDN = $user.onPremisesDistinguishedName

		if (($rtd -eq "UserMailbox") -and $User.department -and $User.companyName -and $User.onPremisesSyncEnabled) {
			$isStdUser = $true
		}

		if 	(($onpremSAM -like "q*") -and ($onpremSAM -notlike "qt*") -and ($onpremSAM -notlike "qk*")) {
			$isStdUser = $false
		}

		if ($onpremDN -like "*OU=Groupware,OU=aplikacni,OU=uzivatele,DC=cezdata,DC=corp" ) {
			$isStdUser = $false
		}
		Write-Log " "
		Write-Log "$($userMbx.PrimarySmtpAddress.ToUpper()) (mbx created:$($whnCre), UPN:$($upn) KPJM:$($onpremSAM) stduser:$($isStdUser))"

		###############################################################################
		# enable Audit
		Try {
			Write-Log "$($upn) Set-Mailbox -AuditEnabled `$true ... " -NoNewline
			Set-Mailbox -Identity $upn -AuditEnabled $true -WarningAction SilentlyContinue
			Write-Log "success" -ForegroundColor Green -NoLinePrefix
		}
		Catch {
			$MbxProcessedCorrectly = $false
			Write-Log "ERROR" -ForegroundColor Red -NoLinePrefix
			Write-Log "  $($_.Exception.Message)" -MessageType "ERROR"
		}
		
		###############################################################################
		# set audit actions 
		Try {
			Write-Log "$($upn) SetMailbox -AuditAdmin ... -AuditDelegate ...  -AuditOwner  ... " -NoNewline
			$AuditAdmin = $AuditAdminActionsAll -split ","
			$AuditDelegate = $AuditDelegateActionsAll -split ","
			$AuditOwner = $AuditOwnerActionsAll -split ","
			Set-Mailbox -Identity $upn -AuditAdmin $AuditAdmin -AuditDelegate $AuditDelegate -AuditOwner $AuditOwner -WarningAction SilentlyContinue
			Write-Log "success" -ForegroundColor Green -NoLinePrefix
		}
		Catch {
			$MbxProcessedCorrectly = $false
			Write-Log "ERROR" -ForegroundColor Red -NoLinePrefix
			Write-Log "  $($_.Exception.Message)" -MessageType "ERROR"
		}

		###############################################################################
		# set MaxReceiveSize and MaxSendSize
		Try {
			Write-Log "$($upn) Set-Mailbox -MaxReceiveSize 150MB ... " -NoNewline
			Set-Mailbox -Identity $upn -MaxReceiveSize $MaxReceiveSize -WarningAction SilentlyContinue
			Write-Log "success" -ForegroundColor Green -NoLinePrefix
		}
		Catch {
			$MbxProcessedCorrectly = $false
			Write-Log "ERROR" -ForegroundColor Red -NoLinePrefix
			Write-Log "  $($_.Exception.Message)" -MessageType "ERROR"
		}
		Try {
			Write-Log "$($upn) Set-Mailbox -MaxSendSize 150MB ... " -NoNewline
			Set-Mailbox -Identity $upn -MaxSendSize $MaxSendSize -WarningAction SilentlyContinue
			Write-Log "success" -ForegroundColor Green -NoLinePrefix
		}
		Catch {
			$MbxProcessedCorrectly = $false
			Write-Log "ERROR" -ForegroundColor Red -NoLinePrefix
			Write-Log "  $($_.Exception.Message.Split(":")[1])" -MessageType "ERROR"
		}
		
		###############################################################################
		# SmtpAuth
		if (-not($Members_SMTPAuthEnabled.id -Contains $User.id)) {
			Try {
				Write-Log "$($upn) Set-CASMailbox -SmtpClientAuthenticationDisabled `$true ... " -NoNewline
				Set-CASMailbox -Identity $upn -SmtpClientAuthenticationDisabled $true -WarningAction SilentlyContinue
				Write-Log "success" -ForegroundColor Green -NoLinePrefix
			}
			Catch {
				$MbxProcessedCorrectly = $false
				Write-Log "ERROR" -ForegroundColor Red -NoLinePrefix
				Write-Log "  $($_.Exception.Message)" -MessageType "ERROR"
			}
		}
		
		###############################################################################
		# IMAP
		if (-not($Members_IMAPEnabled.id -Contains $User.id)) {
			Try {
				Write-Log "$($upn) Set-CASMailbox -ImapEnabled `$false ... " -NoNewline
				Set-CASMailbox -Identity $upn -ImapEnabled $false -WarningAction SilentlyContinue
				Write-Log "success" -ForegroundColor Green -NoLinePrefix
			}
			Catch {
				$MbxProcessedCorrectly = $false
				Write-Log "ERROR" -ForegroundColor Red -NoLinePrefix
				Write-Log "  $($_.Exception.Message)" -MessageType "ERROR"
			}
		}
		
		###############################################################################
		#POP
		if (-not($Members_POPEnabled.id -Contains $User.id)) {
			Try {
				Write-Log "$($upn) Set-CASMailbox  -PopEnabled `$false ... " -NoNewline
				Set-CASMailbox -Identity $upn -PopEnabled $false -WarningAction SilentlyContinue
				Write-Log "success" -ForegroundColor Green -NoLinePrefix
			}
			Catch {
				$MbxProcessedCorrectly = $false
				Write-Log "ERROR" -ForegroundColor Red -NoLinePrefix
				Write-Log "  $($_.Exception.Message)" -MessageType "ERROR"
			}
		}
		
		###############################################################################
		# set myanalytics to opt-out
		Try {
			Write-Log "$($upn) Set-MyAnalyticsFeatureConfig  -PrivacyMode opt-out -Feature all -IsEnabled `$false ... " -NoNewline
			Set-MyAnalyticsFeatureConfig -Identity $upn -PrivacyMode opt-out -Feature all -IsEnabled $false -WarningAction SilentlyContinue | Out-Null
			Write-Log "success" -ForegroundColor Green -NoLinePrefix
		}
		Catch {
			$MbxProcessedCorrectly = $false
			Write-Log "ERROR" -ForegroundColor Red -NoLinePrefix
			Write-Log "  $($_.Exception.Message)" -MessageType "ERROR"
		}

		###############################################################################
		# Viva Insights
		Try {
			Write-Log "$($upn) Set-VivaInsightsSettings -Feature headspace -Enabled `$false ... " -NoNewline
			Set-VivaInsightsSettings -Identity $upn -Feature headspace -Enabled $false -WarningAction SilentlyContinue | Out-Null
			Write-Log "success" -ForegroundColor Green -NoLinePrefix
		}
		Catch {
			$MbxProcessedCorrectly = $false
			Write-Log "ERROR" -ForegroundColor Red -NoLinePrefix
			Write-Log "  $($_.Exception.Message)" -MessageType "ERROR"
		}

		###############################################################################
		# set litigation hold
		if ($Members_LHDisabled.userPrincipalName -Contains($upn)) {
			Write-Log "$($upn) member of $($Name_LHDisabled), skipping litigation hold"
		}
		else {
			if ($isStdUser) {
				If (-not $userMbx.LitigationHoldEnabled) {
					Try {
						Write-Log "$($upn) Set-Mailbox -LitigationHoldEnabled `$true -LitigationHoldDuration $($LHDuration) ... " -NoNewline
						Set-Mailbox -Identity $upn -LitigationHoldEnabled $true -LitigationHoldDuration  $LHDuration -WarningAction SilentlyContinue -ErrorAction Stop
						Write-Log "success" -ForegroundColor Green -NoLinePrefix
					}
					Catch {
						if ($_.Exception.Message.Contains("Exchange Online license doesn't permit")) {
							Write-Log "ERROR - NO LICENSE FOR LH" -ForegroundColor Red -NoLinePrefix
						}
						Else {
							$MbxProcessedCorrectly = $false
							Write-Log "ERROR" -ForegroundColor Red -NoLinePrefix
							Write-Log "  $($_.Exception.Message)" -MessageType "ERROR"
						}
					}
				} 
				Else {
					$LHD = $userMbx.LitigationHoldDuration.Substring(0,$userMbx.LitigationHoldDuration.IndexOf("."))
					If ($LHD -ne $LHDuration) {
						Try {
							Write-Log "$($upn) Set-Mailbox -LitigationHoldDuration $($LHDuration) (was $($LHD)) ... " -NoNewline
							Set-Mailbox -Identity $upn -LitigationHoldDuration $LHDuration -WarningAction SilentlyContinue -ErrorAction Stop
							Write-Log "success" -ForegroundColor Green -NoLinePrefix
						}
						Catch {
							$MbxProcessedCorrectly = $false
							Write-Log "ERROR" -ForegroundColor Red -NoLinePrefix
							Write-Log "  $($_.Exception.Message)" -MessageType "ERROR"
						}
					}
				}
			}
			else {
				Write-Log "$($upn) not a standard user, skipping litigation hold"	
			}
		}
		
		###############################################################################
		# add user to TMS_ICTS_support_cez@cez.cloud (Office 365 poradna)
		if ($isStdUser) {
			if ($Members_O365Poradna.userPrincipalName -Contains($upn)) {
				Write-Log "$($upn) already member of $($Name_O365Poradna), skipping"
			}
			else {
				Write-Log "$($upn) add user to TMS_ICTS_support_cez@cez.cloud ... " -NoNewline
				$result = Add-GraphGroupMemberById -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -GroupId $GroupId_O365Poradna -userId $user.Id
				if (-not($result.StartsWith("ERROR"))) {
					Write-Log "success" -ForegroundColor Green -NoLinePrefix
				}
				else {
					Write-Log "ERROR" -ForegroundColor Red -NoLinePrefix
					Write-Log $result -MessageType "ERROR"
				}
			}
		}
		Else {
			Write-Log "$($upn) not a standard user, skipping adding to teams"	
		}
		
		
		###############################################################################
		# set mailbox regional configuration
		$regConf = Get-MailboxRegionalConfiguration -Identity $userMbx.DistinguishedName	
		if ((-not $regConf.TimeZone) -or (-not $regConf.Language)) {
			Try {
				Write-Log "$($upn) Set-MailboxRegionalConfiguration (Language,TimeFormat,DateFormat,TimeZone) ... " -NoNewline
				Set-MailboxRegionalConfiguration -Identity $upn -Language cs-CZ -TimeFormat "H:mm" -DateFormat "dd.MM.yyyy" -TimeZone "Central Europe Standard Time" -WarningAction SilentlyContinue
				Write-Log "success" -ForegroundColor Green -NoLinePrefix
			}
			Catch {
				$MbxProcessedCorrectly = $false
				Write-Log "ERROR" -ForegroundColor Red -NoLinePrefix
				Write-Log "  $($_.Exception.Message)" -MessageType "ERROR"
			}
		} 	
		else {
			Write-Log "$($upn) mailbox regional configuration already set to CZ, skipping"	
		}

		###############################################################################
		# mark mailbox as processed
		if ($MbxProcessedCorrectly) {
			$ProcessedMailboxesCount++
			$MailboxRecord = [PSCustomObject]@{
				userPrincipalName 	= $userMbx.userprincipalname
				ExchangeGuid 		= $userMbx.ExchangeGuid
				WhenMailboxCreated 	= $userMbx.WhenMailboxCreated
				processedDate       = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
			}
			$EXOMboxMgmt_DB.Add($userMbx.ExchangeGuid, $MailboxRecord)
    		$DB_changed = $true
		}
	}
}

Write-Log "Ignored mailboxes: $($IgnoredMailboxesCount)"
Write-Log "Processed mailboxes: $($ProcessedMailboxesCount)"

#find expired blobs in DB
[datetime]$date = (get-date).AddDays(-$DaysBack-2)
foreach ($ExchangeGuid in $EXOMboxMgmt_DB.Keys) {
    if ($EXOMboxMgmt_DB[$ExchangeGuid].processedDate -lt $date) {
        $ToBeDeletedRecords += $ExchangeGuid
        write-host "Expired record: $($EXOMboxMgmt_DB[$EXOMboxMgmt_DB].userPrincipalName) processed: $($EXOMboxMgmt_DB[$EXOMboxMgmt_DB].processedDate) " -ForegroundColor Red
    }
}
Write-Log "Expired records in DB: $($ToBeDeletedRecords.Count)"

#delete expired blobs from DB
if (ToBeDeletedRecords.Count -gt 0) {
	Write-Log "Deleting expired records from DB..."
	foreach ($ExchangeGuid in $ToBeDeletedRecords) {
		#$EXOMboxMgmt_DB.Remove($ExchangeGuid)
		$DB_changed = $true
	}
}


#saving DB XML if needed
if (($EXOMboxMgmt_DB.count -gt 0) -and ($DB_changed)){
    Try {
        $EXOMboxMgmt_DB | Export-Clixml -Path $DBFileEXOMboxMgmt
        Write-Log "DB file $($DBFileEXOMboxMgmt) exported successfully, $($EXOMboxMgmt_DB.count) records saved"
    }
    Catch {
        Write-Log "Error exporting $($DBFileEXOMboxMgmt)" -MessageType "Error"
    }
}

Get-PSSession | Remove-PSSession

#######################################################################################################################

. $IncFile_StdLogEndBlock
