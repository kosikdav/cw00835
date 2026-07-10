#######################################################################################################################
# Update-T2T-SMTPAddresses.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
	[string]$SourceFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1
. $ScriptPath\include\include-functions-T2T.ps1

if ($SourceMailbox -and $SourceFile) {
	Write-Host "Please specify either SourceMailbox or SourceFile parameter, not both." -ForegroundColor Red
	Exit
}

#######################################################################################################################

$LogFolder			= "t2t"
$LogFilePrefix		= "update-t2t-smtpaddresses"
$LogFileFreq		= "Y"

#######################################################################################################################

. $ScriptPath\include-Script-Start-Include.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Freq $LogFileFreq -Ext "log"

$UJVREZdomains = @(
    "cvrez.cz",
    "egp.cz",
    "engineeringpraha.cz",
    "icvr.cz",
    "nqsafe.cz",
	"radiomedic.cz",
    "skodapraha.cz",
    "ujv.cz",
    "vzuplzen.cz"
)

#######################################################################################################################

. $IncFile_StdLogStartBlock

$OldMailboxes = Import-CSVToArray -Path $SourceFile

$ExchangeSession = New-PSSession -Name "OnPremExchange" -ConfigurationName "Microsoft.Exchange" -ConnectionUri "http://cw00616exch3.cezdata.corp/PowerShell/" -Authentication Kerberos
Import-PSSession $ExchangeSession -DisableNameChecking -AllowClobber

########################################################################################################################
# MAIN PROCESSING LOOP
########################################################################################################################
foreach ($OldMailbox in $OldMailboxes) {
	$MissingSMTPAddresses = $null
	$ExchangeGuid = $OldMailbox.ExchangeGuid
	write-host ("-" * 120) -ForegroundColor DarkGray
	write-host "$($OldMailbox.PrimarySMTPAddress) ($($ExchangeGuid))" -ForegroundColor Cyan

	[array]$OldSMTPAddresses = $OldMailbox.smtpAddresses -split ";" | Where-Object { $_.split("@")[1] -in $UJVREZdomains }
	Try {
		$NewMailbox = Get-RemoteMailbox -Identity $ExchangeGuid -ErrorAction Stop
	}
	Catch {
		continue
	}
	$CurrentPrimarySMTPAddress = $NewMailbox.PrimarySMTPAddress
	$CurrentSMTPAddresses = $NewMailbox.EmailAddresses | Where-Object { $_ -like "smtp:*" }
	$CurrentSMTPAddresses = $CurrentSMTPAddresses -replace "^smtp:", ""
	if ($CurrentPrimarySMTPAddress -ne $OldMailbox.PrimarySMTPAddress) {
		write-log "$($ExchangeGuid)  changing primary SMTP to: $($OldMailbox.PrimarySMTPAddress)" -ForegroundColor Green
		Set-RemoteMailbox -Identity $ExchangeGuid -PrimarySmtpAddress $OldMailbox.PrimarySMTPAddress -ErrorAction continue
	}
	
	$MissingSMTPAddresses = $OldSMTPAddresses | Where-Object { $_ -notin $CurrentSMTPAddresses -and $_ -ne $OldMailbox.PrimarySMTPAddress }
	if ($MissingSMTPAddresses) {
		foreach ($MissingSMTP in $MissingSMTPAddresses) {
			write-log "$($ExchangeGuid)  adding missing SMTP address: $MissingSMTP" -ForegroundColor Green
			Set-RemoteMailbox -Identity $ExchangeGuid -EmailAddresses @{Add="smtp:"+$MissingSMTP} -ErrorAction Continue
		}
	}
}

#######################################################################################################################

. $IncFile_StdLogEndBlock
