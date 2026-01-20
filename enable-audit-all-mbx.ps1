$EnableOnScreenLogging = $true
$ScriptName = $MyInvocation.MyCommand.Name
$Stopwatch  =  [system.diagnostics.stopwatch]::StartNew()

# Import code for function "GetLogFileName"
. d:\scripts\include-function-GetLogFileName.ps1
# Import code for root folders vars
. d:\scripts\include-root-vars.ps1

$LogPath 		= $root_log_folder + "mbxmgmt\"
$LogFilePrefix	= "enable-audit-all-"
$LogFileName    = GetLogFileName("YMD")

[datetime]$date 		= (get-date).AddDays(-5)
#[string]$userMbxFilter	= "(alias -like '*') -and (RecipientTypeDetails -eq 'UserMailbox') -and ((userprincipalname -notlike 'qp*') -and (userprincipalname -notlike 'qs*') -and (userprincipalname -notlike 'qr*')) -and ((alias -notlike 'qp*') -and (alias -notlike 'qs*') -and (alias -notlike 'qr*')) -and (WhenMailboxCreated -gt '$date')"
[array]$userMbxSet = @()

# Import code for AAD app reg CEZ_EXO_MBX_MGMT
. d:\scripts\include-appreg-CEZ_EXO_MBX_MGMT.ps1

# Import code for function "LogWrite"
. d:\scripts\include-function-logwrite.ps1

Import-Module ExchangeOnlineManagement
Import-Module MSAL.PS

###############################################################################################################################################
# MAIN SCRIPT BODY
###############################################################################################################################################
LogWrite -LogString "/----------------------------------------------------------------------------"
LogWrite -LogString "AAD service principal id: $ClientId"
LogWrite -LogString "/----------------------------------------------------------------------------"

$Certificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"

for ($first = 0; $first -lt 26; $first++) {
	for ($second = 0; $second -lt 26; $second++) {
		Connect-ExchangeOnline -CertificateThumbPrint $ThumbPrint -AppID $ClientId -Organization $TenantName
		for ($third = 0; $third -lt 26; $third++) {
			$abc = [char](65+$first)+[char](65+$second)+[char](65+$third)
			LogWrite -LogString "/--------------------------------------------------------"
			$userMbxFilter	= "(alias -like '$($abc)*')"
			LogWrite -LogString "$($ABC) - $($userMbxFilter)"
			$userMbxSet = Get-Mailbox -ResultSize Unlimited -Filter $userMbxFilter

			ForEach ($userMbx in $userMbxSet) {
				Try {
					LogWrite -LogString "  >> Set-Mailbox -Identity $($userMbx.identity) -AuditEnabled `$true"
					Set-Mailbox -Identity $userMbx.identity -AuditEnabled $true -WarningAction Stop
				}
				Catch {LogWrite -LogString "  $($_.Exception.Message.Split(":")[1])" -MessageType "ERROR"}
					
				Try {
					LogWrite -LogString "  >> Set-Mailbox -Identity $($userMbx.identity) -DefaultAuditSet Admin,Delegate,Owner"
					Set-Mailbox -Identity $userMbx.identity -DefaultAuditSet Admin,Delegate,Owner -WarningAction Stop
				}
				Catch {LogWrite -LogString "  $($_.Exception.Message.Split(":")[1])" -MessageType "ERROR"}	
			}

			LogWrite -LogString "\--------------------------------------------------------"
		}
		Get-PSSession | Remove-PSSession
	}
}

LogWrite -LogString "Run time: $($Stopwatch.Elapsed)"
LogWrite -LogString "Script end"
LogWrite -LogString "\----------------------------------------------------------------------------"
LogWrite -LogString " "
