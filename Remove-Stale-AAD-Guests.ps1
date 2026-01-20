#######################################################################################################################
# Remove-AAD-Guest
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "aad-guest-mgmt"
$LogFilePrefix		= "remove-stale-aad-guests"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$CutoffDateGuests = (Get-Date (Get-Date).AddDays(-$InactivityLimitGuests) -Format u).Replace(' ','T')
$CutoffDatePendingInvites = (Get-Date (Get-Date).AddDays(-$InactivityLimitPendingInvites) -Format u).Replace(' ','T')
$GuestRemovalReport = @()
$PendingGuestRemovalReport = @()
$currentDate = (Get-Date).ToString("yyyy-MM-dd")

#######################################################################################################################

. $IncFile_StdLogStartBlock

##############################################################################################
# read Guests from Graph 

Write-Log "Inactivity limit for guests: $InactivityLimitGuests days (cutoff date: $CutoffDateGuests)"
Write-Log "Inactivity limit for pending invites: $InactivityLimitPendingInvites days (cutoff date: $CutoffDatePendingInvites)"
write-log $string_divider
Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30

$UriResource = "users"
$UriFilter = "UserType+eq+'Guest'"
$UriSelect = "id,userPrincipalName,displayName,mail,accountEnabled,createdDateTime,employeeHireDate,userType,employeeType,onPremisesExtensionAttributes"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Filter $UriFilter -Select $UriSelect -Top 999
$Guests = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -ContentType $ContentTypeJSON -text "AAD Guests" -ProgressDots

$UriResource = "users"
$UriFilter = "ExternalUserState+eq+'PendingAcceptance'"
$UriSelect = "id,mail,createdDateTime,ExternalUserState"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Filter $UriFilter -Select $UriSelect -Top 999
$PendingInvites = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -ContentType $ContentTypeJSON -text "AAD Guests pending acceptance" -ProgressDots

Write-Log "Guests: $($Guests.count)"
Write-Log "Pending invites: $($PendingInvites.count)"

$counterGuests = 0
ForEach ($Guest in $Guests) {
    $ext15 = $LastAuditTrace = $null
	$ext15 = $Guest.onPremisesExtensionAttributes.extensionAttribute15
	if ($ext15 -and $ext15.StartsWith("XTSync_")) {
		Continue
	}
    Write-Host "$($Guest.mail) " -NoNewline
    if ($Guest.employeeHireDate) {
		Try {
            $LastAuditTrace = [datetime]::Parse($Guest.employeeHireDate)
            $DaysSinceLAT = (New-TimeSpan -Start $LastAuditTrace -End $CurrentDate).Days
            Write-Host "- lastAuditTrace: $($LastAuditTrace) ($($DaysSinceLAT))"
        }
        Catch {
            Write-Host
            Continue
        }
	}
    if ($LastAuditTrace -and $LastAuditTrace -lt $CutoffDateGuests) {
        Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
        Remove-B2BUser -Identity $Guest.id -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -Silent:$true
        $counterGuests++
        Write-Log "$($Guest.userPrincipalName) ($($Guest.id)) guest removed - lastAuditTrace: $($LastAuditTrace) ($($DaysSinceLAT))" -ForegroundColor Yellow
    }
}
write-log $string_divider
write-log "Total guests removed: $counterGuests" -ForegroundColor Yellow

$counterInvites = 0
foreach ($Guest in $PendingInvites) {
    Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
    if ($Guest.createdDateTime) {
        Try {
            $CreatedDateTime = [datetime]::Parse($Guest.createdDateTime)
            $DaysSinceInvite = (New-TimeSpan -Start $CreatedDateTime -End $CurrentDate).Days
        }
        Catch {
            Continue
        }
    }
    if ($CreatedDateTime -and $CreatedDateTime -lt $CutoffDatePendingInvites) {
        Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
        Remove-B2BUser -Identity $Guest.id -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -Silent:$true
        $counterInvites++
        Write-Log "$($Guest.mail) ($($Guest.id)) pending invite removed - createdDateTime: $($CreatedDateTime) ($($DaysSinceInvite))" -ForegroundColor Yellow
    }
}
write-log "Total pending invites removed: $counterInvites" -ForegroundColor Yellow
write-log $string_divider

#######################################################################################################################

. $IncFile_StdLogEndBlock
