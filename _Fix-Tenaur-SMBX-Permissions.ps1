$EnableOnScreenLogging = $true
$ManualAuth = $false
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
$Stopwatch =  [system.diagnostics.stopwatch]::StartNew()
$daysBackOffset = 0

. $ScriptPath\include-Functions-Init.ps1
. $ScriptPath\include-Root-Vars.ps1

##################################################################################################

$LogFolder			= "mbxmgmt"
$LogFilePrefix		= "mbxmgmt"
$LogFileFreq		= "YMD"

[datetime]$date = (get-date).AddDays(-5)
#[string]$userMbxFilter	= "(alias -like '*') -and (RecipientTypeDetails -eq 'UserMailbox') -and ((userprincipalname -notlike 'qp*') -and (userprincipalname -notlike 'qs*') -and (userprincipalname -notlike 'qr*')) -and ((alias -notlike 'qp*') -and (alias -notlike 'qs*') -and (alias -notlike 'qr*')) -and (WhenMailboxCreated -gt '$date')"
$userMbxFilter	= "(alias -like 'qptnr*') -and ((RecipientTypeDetails -eq 'UserMailbox') -or (RecipientTypeDetails -eq 'SharedMailbox'))"
$LHDuration 	= 1825
$MaxReceiveSize = "150MB"
$MaxSendSize 	= "150MB"

#CEZ_EXO_Allow_SMTP_Client_Auth
$GroupId_SMTPAuthEnabled = "7d0dd440-f4e6-40a4-ab14-cecc74040a42"
#CEZ_EXO_Allow_IMAP_Client_access
$GroupId_IMAPEnabled = "751798f5-5e7e-47b0-836c-69b0c5bfbdd3"
#CEZ_EXO_Allow_POP3_Access
$GroupId_POPEnabled = "9c7c879d-308b-44e5-8224-5dabe58656ba"
##################################################################################################

$LogPath            = Join-Path $root_log_folder $LogFolder
$LogFileName        = $LogFilePrefix + "-" + (GetTimestamp($LogFileFreq)) + ".log"

. $ScriptPath\include-Functions-Common.ps1
. $ScriptPath\include-appreg-CEZ_EXO_MBX_MGMT.ps1
. $ScriptPath\include-GetMSALToken.ps1

LogWrite -LogString "/--------------------------------------------------------------------------------"
LogWrite -LogString "Script start: $($ScriptName)"
LogWrite -LogString "AAD app name: $($AppName)"
LogWrite -LogString "Client Id: $($ClientId)"
LogWrite -LogString "litigation hold duration: $LHDuration"
LogWrite -LogString "cutoff date: $date"
LogWrite -LogString "mailbox filter: $userMbxFilter"

. $ScriptPath\include-ConnectEXO.ps1

[array]$userMbxSet = @()

#retrieve mailboxes 
LogWrite -LogString "Looking for Tenaur shared mailboxes"
$qptnrMailboxes = Get-Mailbox -ResultSize Unlimited -Filter $userMbxFilter
LogWrite -LogString "$($userMbxSet.Count) mailboxes returned by query"
foreach ($mailbox in $qptnrMailboxes) {
    Write-Host $mailbox.mail
}
exit

If ($userMbxSet.Count -gt 0) {

	ForEach ($userMbx in $userMbxSet) {
		LogWrite -LogString " "
		$upn,$rtd,$als,$whnCre,$onpremSAM,$onpremDN = $null
		$isStdUser = $false

		$rtd 		= $userMbx.RecipientTypeDetails
		$upn 		= $userMbx.userprincipalname
		$als 		= $userMbx.alias
		$whnCre 	= $userMbx.WhenMailboxCreated

		$Uri = "https://graph.microsoft.com/beta/users/$($upn)"
		try{
			[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
			$Query = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -Uri $Uri -Method GET -ContentType "application/json"
			$User = $Query
			$onpremSAM 	= $User.onPremisesSamAccountName
			$onpremDN 	= $User.onPremisesDistinguishedName		}
		Catch {LogWrite -LogString "  $($_.Exception.Message)" -MessageType "ERROR"}
		
		if (($rtd -eq "UserMailbox") -and ($null -ne $User.department) -and ($null -ne $User.companyName) -and ($User.onPremisesSyncEnabled -eq $true)) {
			$isStdUser = $true
		}

		if 	(($onpremSAM -like "q*") -and ($onpremSAM -notlike "qt*") -and ($onpremSAM -notlike "qk*")) {
			$isStdUser = $false
		}

		if ($onpremDN -like "*OU=Groupware,OU=aplikacni,OU=uzivatele,DC=cezdata,DC=corp" ) {
			$isStdUser = $false
		}
		LogWrite -LogString " "
		LogWrite -LogString "$($userMbx.PrimarySmtpAddress.ToUpper()) (mbx created:$($whnCre), UPN:$($upn) KPJM:$($onpremSAM) stduser:$($isStdUser))"
		
		#enable Audit
		Try {
			LogWrite -LogString "  >> Set-Mailbox -Identity $($upn) -AuditEnabled `$true"
			Set-Mailbox -Identity $upn -AuditEnabled $true -WarningAction Stop
		}
		Catch {LogWrite -LogString "  $($_.Exception.Message.Split(":")[1])" -MessageType "ERROR"}
			
		#set default audit options
		Try {
			LogWrite -LogString "  >> Set-Mailbox -Identity $($upn) -DefaultAuditSet Admin,Delegate,Owner"
			Set-Mailbox -Identity $upn -DefaultAuditSet Admin,Delegate,Owner -WarningAction Stop
		}
		Catch {LogWrite -LogString "  $($_.Exception.Message.Split(":")[1])" -MessageType "ERROR"}

		#set MaxReceiveSize and MaxSendSize
		Try {
			LogWrite -LogString "  >> Set-Mailbox -Identity $($upn) -MaxReceiveSize 150MB"
			Set-Mailbox -Identity $upn -MaxReceiveSize $MaxReceiveSize -WarningAction Stop
		}
		Catch {LogWrite -LogString "  $($_.Exception.Message.Split(":")[1])" -MessageType "ERROR"}
		Try {
			LogWrite -LogString "  >> Set-Mailbox -Identity $($upn) -MaxSendSize 150MB"
			Set-Mailbox -Identity $upn -MaxSendSize $MaxSendSize -WarningAction Stop
		}
		Catch {LogWrite -LogString "  $($_.Exception.Message.Split(":")[1])" -MessageType "ERROR"}
		
		#SmtpAuth
		if ($Members_SMTPAuthEnabled.id -Contains $User.id) {
			#enable SmtpAuth if user is member of CEZ_EXO_Allow_SMTP_Client_Auth
			Try {
				LogWrite -LogString "  >> $($upn) is member of CEZ_EXO_Allow_SMTP_Client_Auth - allowing SMTPClientAuthentication"
				Set-CasMailbox -Identity $upn -SmtpClientAuthenticationDisabled $false -WarningAction Stop
			}
			Catch {LogWrite -LogString "  $($_.Exception.Message.Split(":")[1])" -MessageType "ERROR"}
		}
		else {
			#disable SmtpAuth for all other users
			Try {
				LogWrite -LogString "  >> Set-CasMailbox -Identity $($upn) -SmtpClientAuthenticationDisabled `$true"
				Set-CasMailbox -Identity $upn -SmtpClientAuthenticationDisabled $true -WarningAction Stop
			}
			Catch {LogWrite -LogString "  $($_.Exception.Message.Split(":")[1])" -MessageType "ERROR"}
		}
		
		# IMAP
		if ($Members_IMAPEnabled.id -Contains $User.id) {
			Try {
				#enable IMAP if user is member of CEZ_EXO_Allow_IMAP_Client_access
				LogWrite -LogString "  >> $($upn) is member of CEZ_EXO_Allow_IMAP_Client_access - allowing IMAP"
				Set-CasMailbox -Identity $upn -ImapEnabled $true -WarningAction Stop
			}
			Catch {LogWrite -LogString "  $($_.Exception.Message.Split(":")[1])" -MessageType "ERROR"}
		}
		else {
			#disable IMAP for all other users
			Try {
				LogWrite -LogString "  >> Set-CasMailbox -Identity $($upn) -ImapEnabled `$false"
				Set-CasMailbox -Identity $upn -ImapEnabled $false -WarningAction Stop
			}
			Catch {LogWrite -LogString "  $($_.Exception.Message.Split(":")[1])" -MessageType "ERROR"}
		}
		
		#disable POP
		if ($Members_POPEnabled.id -Contains $User.id) {
			Try {
				#enable SmtpAuth if user is member of CEZ_EXO_Allow_POP3_Access
				LogWrite -LogString "  >> $($upn) is member of CEZ_EXO_Allow_POP3_Access - allowing POP"
				Set-CasMailbox -Identity $upn -PopEnabled $true -WarningAction Stop
			}
			Catch {LogWrite -LogString "  $($_.Exception.Message.Split(":")[1])" -MessageType "ERROR"}
		}
		else {
			Try {
				#disable POP for all other users
				LogWrite -LogString "  >> Set-CasMailbox -Identity $($upn) -PopEnabled `$false"
				Set-CasMailbox -Identity $upn -PopEnabled $false -WarningAction Stop
			}
			Catch {LogWrite -LogString "  $($_.Exception.Message.Split(":")[1])" -MessageType "ERROR"}
		}
		
		#set myanalytics to opt-out
		Try {
			LogWrite -LogString "  >> Set-MyAnalyticsFeatureConfig -Identity $($upn) -PrivacyMode opt-out -Feature all -IsEnabled `$false"
			Set-MyAnalyticsFeatureConfig -Identity $upn -PrivacyMode opt-out -Feature all -IsEnabled $false -WarningAction Stop
		}
		Catch {LogWrite -LogString "  $($_.Exception.Message.Split(":")[1])" -MessageType "ERROR"}

		#set Cortana briefing mail to disabled
		Try {
			LogWrite -LogString "  >> Set-UserBriefingConfig -Identity $($upn) -Enabled `$false"
			Set-UserBriefingConfig -Identity $upn -Enabled $false -WarningAction Stop
		}
		Catch {LogWrite -LogString "  $($_.Exception.Message.Split(":")[1])" -MessageType "ERROR"}

		#set litigation hold
		if ($isStdUser) {
			If ($userMbx.LitigationHoldEnabled -eq $false) {
				Try {
					LogWrite -LogString "  >> Set-Mailbox -Identity $($upn) -LitigationHoldEnabled `$true -LHDuration $($LHDuration)"
					Set-Mailbox -Identity $upn -LitigationHoldEnabled $true -LHDuration $LHDuration -WarningAction Stop -ErrorAction Stop
				}
				Catch {LogWrite -LogString "  $($_.Exception.Message.Split(":")[1])" -MessageType "ERROR"}
			} 
			ElseIf (($userMbx.LitigationHoldEnabled -eq $true) -and ($userMbx.LitigationHoldDuration.Substring(0,$userMbx.LitigationHoldDuration.IndexOf(".")) -ne $LHDuration)) {
				Try {
					LogWrite -LogString "  >> Set-Mailbox -Identity $($upn) -LHDuration $($LHDuration)"
					Set-Mailbox -Identity $upn -LHDuration $LHDuration -WarningAction Stop -ErrorAction Stop
				}
				Catch {LogWrite -LogString "  $($_.Exception.Message.Split(":")[1])" -MessageType "ERROR"}
			}
		}
		else {
			LogWrite -LogString "  $($upn) not a standard user, skipping litigation hold"	
		}
		
		# add user to TMS_ICTS_support_cez@cez.cloud (Office 365 poradna)
		if ($isStdUser) {
			if ($Members_O365Poradna.userPrincipalName -Contains($upn)) {
				LogWrite -LogString "  $($upn) already member of $($GroupId_O365Poradna), skipping"
			}
			else {
				try {
					$params = @{"@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($User.Id)"}
					LogWrite -LogString "  >> New-MgGroupMemberByRef -GroupId $($GroupId_O365Poradna) -BodyParameter `@`{`"@odata.id`" = `"https://graph.microsoft.com/v1.0/directoryObjects/$($User.Id)`"`}"
					New-MgGroupMemberByRef -GroupId $GroupId_O365Poradna -BodyParameter $params -WarningAction Stop -ErrorAction Stop
				}
				Catch {LogWrite -LogString "  $($_.Exception.Message.Split(":")[0,1])" -MessageType "ERROR"}
			}
		}
		Else {
			LogWrite -LogString "  $($upn) not a standard user, skipping adding to teams"	
		}
		
		if ($userMbxSet_NoRegConf.Identity -Contains $als) {
			Try {
				LogWrite -LogString "  >> Set-MailboxRegionalConfiguration -Identity $($upn) -Language 1029 -TimeFormat `"H:mm`" -DateFormat `"dd.MM.yyyy`" -TimeZone `"Central Europe Standard Time`" -ErrorAction stop"
				Set-MailboxRegionalConfiguration -Identity $upn -Language 1029 -TimeFormat "H:mm" -DateFormat "dd.MM.yyyy" -TimeZone "Central Europe Standard Time" -WarningAction Stop -ErrorAction Stop
				LogWrite -LogString "  $($upn) mailbox regional configuration set to CZ"
			}
			Catch {LogWrite -LogString "  $($_.Exception.Message)" -MessageType "ERROR"}
		}
		else {
			LogWrite -LogString "  Mailbox regional configuration set to CZ already, skipping"	
		}
	}

	#retrieve identities with default regional settings (users with null settings in TimeZone or Language)
	#LogWrite -LogString "-----------------------------------------------------------"
	#LogWrite -LogString "--- mailbox regional configuration ------------------------"
	#LogWrite -LogString "-----------------------------------------------------------"
	#LogWrite -LogString "$($userMbxSet_NoRegConf.Count) of the returned mailboxes have empty regional configuration"

	<#
	If ($userMbxSet_NoRegConf.Count -gt 0) {
		ForEach ($userMbx_NoRegConf in $userMbxSet_NoRegConf) {
			write-host $userMbx_NoRegConf | fl
			Try {
				LogWrite -LogString "  >> Set-MailboxRegionalConfiguration -Identity $($userMbx_NoRegConf.identity) -Language 1029 -TimeFormat `"H:mm`" -DateFormat `"dd.MM.yyyy`" -TimeZone `"Central Europe Standard Time`" -ErrorAction stop"
				Set-MailboxRegionalConfiguration -Identity $userMbx_NoRegConf.identity -Language 1029 -TimeFormat "H:mm" -DateFormat "dd.MM.yyyy" -TimeZone "Central Europe Standard Time" -WarningAction Stop -ErrorAction Stop
				LogWrite -LogString " $($userMbx_NoRegConf.'identity') mailbox regional configuration set to CZ"
			}
			Catch {LogWrite -LogString "  $($_.Exception.Message)" -MessageType "ERROR"}
		}
	}
	#>

}

Get-PSSession | Remove-PSSession
LogWrite -LogString "EXO disconnected"
LogWrite -LogString "Run time: $($Stopwatch.Elapsed)"
LogWrite -LogString "Script end"
LogWrite -LogString "\--------------------------------------------------------------------------------"
LogWrite -LogString " "
