#######################################################################################################################
# Copy-GuestGroupMembership.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
	[Parameter(Mandatory = $true)][string]$SourceUser,
	[Parameter(Mandatory = $true)][string]$TargetUser
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder				= "copy-guest-group-membership"
$LogFilePrefix			= "copy-guest-group-membership"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile 		= New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

function Get-YesNoKeyboardInput {
    param (
        [Parameter(Mandatory=$true)][string]$Prompt
    )
    Write-Host "$($Prompt) [Y/N]" -ForegroundColor Yellow
    :prompt 
    while ($true) {
        switch ([console]::ReadKey($true).Key) {        
            { $_ -eq [System.ConsoleKey]::Y } { Return $true }        
            { $_ -eq [System.ConsoleKey]::N } { Return $false }        
            default { Write-Host "Only 'Y' or 'N' allowed!" }    
        }
    }
}

##################################################################################################

. $IncFile_StdLogStartBlock

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

$UriResource = "users/$($SourceUser)"
$UriSelect = "id,userPrincipalName,mail,accountEnabled,onPremisesExtensionAttributes,userType"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
$SourceGuest = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
if (-not ($SourceGuest)) {
    write-host "Source user not found"
    Exit
}
else {
	if ($SourceGuest.UserType -ne "Guest") {
		write-host "Source user is not a guest user"
		Exit
	}
}

write-host $string_divider
write-host "Source user: " -NoNewline
write-host "$($SourceGuest.mail) " -ForegroundColor Yellow -NoNewline
write-host "($($SourceGuest.userPrincipalName) $($SourceGuest.id)) " -ForegroundColor DarkGray

$UriResource = "users/$($TargetUser)"
$UriSelect = "id,userPrincipalName,mail,accountEnabled,onPremisesExtensionAttributes,userType"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
$TargetGuest = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
if (-not ($TargetGuest)) {
	write-host "Target user not found"
	Exit
}
else {
	if ($TargetGuest.UserType -ne "Guest") {
		write-host "Target user is not a guest user"
		Exit
	}
}
write-host "Target user: " -NoNewline
write-host "$($TargetGuest.mail) " -ForegroundColor Yellow -NoNewline
write-host "($($TargetGuest.userPrincipalName) $($TargetGuest.id))" -ForegroundColor DarkGray
write-host $string_divider

$UriResource = "users/$($SourceUser)/memberOf"
$UriFilter = "groupTypes/any(c:c eq 'Unified') and not groupTypes/any(c:c eq 'DynamicMembership')"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Filter $UriFilter -Count
$SourceMemberships = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ConsistencyLevel "eventual"
if ($SourceMemberships) {
	write-host "Source guest memberships found: $($SourceMemberships.Count)"
	foreach ($Membership in $SourceMemberships) {
		write-host "$($Membership.displayName) " -ForegroundColor Cyan -NoNewline
		write-host "($($Membership.id))" -ForegroundColor DarkGray
	}
}
else {
	write-host "Source user not member of any group" -ForegroundColor Red
	Exit
}

$UriResource = "users/$($TargetUser)/memberOf"
$UriFilter = "groupTypes/any(c:c eq 'Unified') and not groupTypes/any(c:c eq 'DynamicMembership')"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Filter $UriFilter -Count
$TargetCurrentMemberships = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ConsistencyLevel "eventual"

if ($TargetCurrentMemberships) {
	$TargetMissingMemberships = $SourceMemberships | Where-Object { $_.id -notin $TargetCurrentMemberships.id }
	if ($TargetMissingMemberships) {
		write-host "Target user is missing following group memberships of source user: $($TargetMissingMemberships.Count)" -ForegroundColor Yellow
		foreach ($Membership in $TargetMissingMemberships) {
			write-host "$($Membership.displayName) " -ForegroundColor Cyan -NoNewline
			write-host "($($Membership.id))" -ForegroundColor DarkGray
		}
	}
	else {
		write-host "Target user already has all group memberships of source user" -ForegroundColor Green
		Exit
	}
}

if ((Get-YesNoKeyboardInput "Continue with adding memberships?")) {
	Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
	foreach ($Membership in $TargetMissingMemberships) {
		$result = Add-GraphGroupMemberById -GroupId $Membership.id -userId $TargetGuest.id -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken
		if (-not($result.StartsWith("ERROR"))) {
			write-log "Adding $($TargetGuest.id) to group $($Group.displayName) ($($Group.membershipRule))"
		}
		else {
			write-log "ERROR: $result" -MessageType "ERR"
		}
	}
}


#######################################################################################################################

. $IncFile_StdLogEndBlock
