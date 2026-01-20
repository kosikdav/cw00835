#######################################################################################################################
# New-AAD-Guest
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    [string]$SourceFile,
    [string]$Mail,
    [string]$Owner
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

. $ScriptPath\include-Script-StdStartBlock.ps1

if (-not $SourceFile) {
    if (-not $Mail) {
        Write-Host "ERROR: Missing parameter -Mail" -ForegroundColor Red
        exit
    }
    if (-not $Owner) {
        Write-Host "ERROR: Missing parameter -Owner" -ForegroundColor Red
        exit
    }
}


#######################################################################################################################

$LogFolder			= "new-aad-guest"
$LogFilePrefix		= "new-aad-guest"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$Line1 = "Byli jste pozvani ke spolupraci v ramci Microsft 365 prostredi Skupiny CEZ na zadost uzivatele $($Owner)."
$Line2 = "Pozvanku prijmete kliknutim na odkaz nize."
$Line3 = "Pokud budete vyzvani k volbe typu uctu, zvolte prosim pracovni/skolni ucet, nikoli osobni."
$DefaultInvitationMessage = $Line1 , $Line2 , $Line3 -join " "
function New-B2BUser {
    param (
        [Parameter(Mandatory=$true)][string]$Mail,
        [Parameter(Mandatory=$true)][string]$Owner,
        [string]$InvitationMessage,
        [Parameter(Mandatory=$true)][string]$AccessToken
    )
    $UriResource = "users"
    $UriFilter = "proxyAddresses/any(c:c eq '$($Mail)')+and+UserType+eq+'Guest'"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Filter $UriFilter
    $ExistingGuest = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -ContentType $ContentTypeJSON
    if (-not $ExistingGuest) {
        if (-not $InvitationMessage) {
            $InvitationMessage = $DefaultInvitationMessage
        }
        $UriResource = "invitations"
        $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
        $GraphBodyCreate = @{
            invitedUserEmailAddress = $Mail;
            inviteRedirectUrl = "https://myapps.microsoft.com";
            sendInvitationMessage = $true;
            invitedUserType = "Guest";
            invitedUserMessageInfo = @{
                customizedMessageBody = $InvitationMessage
            }
        } | ConvertTo-Json
        $Headers = @{Authorization = "Bearer $($AccessToken)"}
        
        Try {
            $resultCREATE = Invoke-RestMethod -Headers $Headers -Uri $Uri -Body $GraphBodyCreate -Method "POST" -ContentType $ContentTypeJSON
            $guestId = $resultCREATE.invitedUser.id
            Write-Log "Invitation sent to $($Mail), userid $($guestId)" -ForegroundColor "Cyan"
            start-sleep -Seconds 5
            
            $stampString = $Owner.Trim().ToLower() + ";" + (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
            $GraphBodyEmployeeType = @{
                employeeType = $stampString
            } | ConvertTo-Json
            $UriResource = "users/$($guestId)"
            $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
            Try {
                $ResultUPDATE = Invoke-RestMethod -Headers $Headers -Uri $Uri -Body $GraphBodyEmployeeType -Method "PATCH" -ContentType $ContentTypeJSON
                Write-Log "SUCCESS employeeType: $($Guest.mail) `"$($stampString)`"" -ForegroundColor "Green"
            }
            Catch {
                $ErrorMessagePATCH = $_.ErrorDetails.Message | Out-String
                Write-Log "ERR PATCH employeeType: $($Guest.Mail) `"$($stampString)`"" -MessageType Error
                Write-Log $($ErrorMessagePATCH) -MessageType Error
            }
        }
        Catch {
            $ErrorMessage = $_.ErrorDetails.Message | Out-String
            Write-Log "ERR create invite for $($Mail)" -MessageType Error
            Write-Log $($ErrorMessage) -MessageType Error
            return
        }
    }
    else {
        Write-Log "B2B user $($Mail) already exists" -ForegroundColor "Yellow"
        continue
    }
}
#######################################################################################################################

. $IncFile_StdLogStartBlock

##############################################################################################
# read Guests from Graph 

if ($SourceFile) {
    #process source file
    $SourceList = Import-CSVtoArray -Path $SourceFile
    if ($SourceList.Count -gt 0) {
        if ($SourceList[0].Mail -and $SourceList[0].Owner) {
            foreach ($ListRecord in $SourceList) {
                Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
                New-B2BUser -Mail $ListRecord.Mail -Owner $ListRecord.Owner -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken
            }
        }
        else {
            Write-Log "ERROR: source file does not contain required columns" -ForegroundColor "Red"
            Exit
        }
    }
    else {
        Write-Log "ERROR: source file empty" -ForegroundColor "Red"
        Exit
    }
}
else {
    #process single user
    New-B2BUser -Mail $Mail -Owner $Owner -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken
}

#######################################################################################################################

. $IncFile_StdLogEndBlock
