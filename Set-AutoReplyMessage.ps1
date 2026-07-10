#######################################################################################################################
# Set-MaiboxT2T-Properties
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
	[string]$SourceFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1
. $ScriptPath\include\include-functions-T2T.ps1

#######################################################################################################################

$LogFolder			= "exo\autoreply"
$LogFilePrefix		= "set-autoreply"
$LogFileFreq		= "Y"

function Get-HTMLBody {
    param (
        [string]$OldAddress,
        [string]$NewAddress,
        [string]$DisplayName
    )

$Body = 
@"
    <html>
    <body style="font-family: Arial, sans-serif; font-size: 11pt; color: #000000;">
    <p>
    Dobrý den,
    </p>
    <p>
    adresa <b>$OldAddress</b> již není aktivní.</p>
    <p>Váš e-mail byl přesměrován na moji novou adresu <a href="mailto:$NewAddress"><b>$NewAddress</b></a>,
    <br>kam jsem se přesunul spolu s celým úsekem Zelená energetika ČEZ ESCO od 1. července 2026.
    </p>
    <p>
    Tuto adresu také prosím používejte pro budoucí komunikaci se mnou - můžeme tak spolu i nadále v dresu ČEZ řešit témata jako
    <br>fotovoltaické elektrárny, bateriová úložiště, elektromobilita či nově také systémy chlazení a vytápění!
    </p>
    <p>
    Děkuji za pochopení
    </p>
    <p>
    $DisplayName
    </p>
    </body>
    </html>
"@

return $Body
}

#######################################################################################################################

. $ScriptPath\include-Script-Start-Include.ps1

Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 120

$UserData = Import-CsvToArray -Path $SourceFile

if ($UserData) {
    foreach ($user in $UserData) {
        Write-Log  "Setting autoreply for $($user.mailbox)" -NoNewLine
        $htmlBody = Get-HTMLBody -OldAddress $user.old_address -NewAddress $user.new_address -DisplayName $user.display_name
        Try {
        Set-MailboxAutoReplyConfiguration -Identity $user.mailbox -AutoReplyState Enabled -InternalMessage $htmlBody -ExternalMessage $htmlBody -ExternalAudience All
            Write-Interactive " OK" -ForegroundColor Green
        }
        Catch {
            Write-Log "Error setting autoreply for $($user.mailbox): $($_.Exception.Message)"
        }
    }
}
