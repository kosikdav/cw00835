
# Set-AAD-Synced-Users-Type.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "aad-guest-mgmt"
$LogFilePrefix		= "aad-ext-usr-type"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

#######################################################################################################################

. $IncFile_StdLogStartBlock

$guestsFixed = 0

##############################################################################################
# read ext member users from Graph 
Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
$UriResource = "users"
$UriFilter = "userType eq 'Member' and externalUserState eq 'Accepted'"
$UriSelect = "id,userPrincipalName,userType,displayName,externalUserState,onPremisesSyncEnabled,onPremisesExtensionAttributes"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Filter $UriFilter -Select $UriSelect
[array]$AllExtMedmbers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -ContentType $ContentTypeJSON 
if (-not $AllExtMedmbers) {
    Exit
}
foreach ($user in $AllExtMedmbers) {
    if ((-not $user.id) -or ($user.onPremisesSyncEnabled -eq $true)) {
        Continue
    }
    if (($null -ne $user.onPremisesExtensionAttributes.extensionAttribute15) -and ($user.onPremisesExtensionAttributes.extensionAttribute15 -ne "")) {
        $userString = "$($user.id) $($user.userPrincipalName) ($($user.displayName)) synced from: $($user.onPremisesExtensionAttributes.extensionAttribute15)"
    }
    else {
        $userString = "$($user.id) $($user.userPrincipalName) ($($user.displayName))"
    }
    $UriResource = "users/$($user.id)"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
    $Body = @{
        "userType" = "Guest"
    } | ConvertTo-Json
    Try {
        $Response = Invoke-WebRequest -Uri $Uri -Method "PATCH" -Body $Body -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -ContentType $ContentTypeJSON -UseBasicParsing
        Write-Log "Updated $($userString)"
        $guestsFixed++
    }
    Catch {
        Write-Log "Error updating $userString" -MessageType "ERR"
        Write-Log $_.Exception.Message -MessageType "ERR"
    }
}   
Write-Log "Total users updated: $($guestsFixed)"

#######################################################################################################################

. $IncFile_StdLogEndBlock
