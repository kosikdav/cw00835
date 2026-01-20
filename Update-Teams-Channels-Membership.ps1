#######################################################################################################################
# Update-Teams-Channels-Membership.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder          = "update-teams"
$LogFilePrefix      = "update-teams-channels-membership"
$LogFileFreq        = "YMD"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Suffix $LogFileSuffix -Ext "log"

$TMS_KomunitaCPR_Id = "a2bded8b-7bc4-48f2-819f-ed3b47f34d02"

#Channels to process - these are the channel names that will be processed (prefix search) and corresponding user location city (prefix search)
$ChannelsToProcess = @{
    "Komunita Ostrava" = "Ostrava"
    "Komunita Plze" = "Plze"
    "Komunita Praha" = "Praha"
    "Komunita Kol" = "Kol"
}

$CustomCityMapping_DB = @{
    "zdenek.dubnicky@cez.cz" = "Praha"
    "petr.fronek@cez.cz" = "Praha"
    "radek.bojda@cez.cz" = "Praha"
}

#User whitelist - these users will not be removed from any channel
$UserWhitelist = @(
    # Jan Kavalir
    #"446aca97-f738-4955-9ffc-7a5d293675f9"
)

#######################################################################################################################

. $IncFile_StdLogStartBlock

Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

##########################################################################################

Write-Log $String_Divider
Write-Log "Processing `"$(Get-GroupNameFromGraphById -id $TMS_KomunitaCPR_Id -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken)`" ($($TMS_KomunitaCPR_Id))" -ForegroundColor Cyan

$UriResource = "users"
$UriSelect = "id,userPrincipalName,mail,city"
$UriFilter = "accountEnabled+eq+true+and+userType+eq+'Member'"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect -Filter $UriFilter -Top 999
[array]$AllUsers = Get-GraphOutputREST -Uri $Uri -ContentType $ContentTypeJSON -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -Text "all AAD users" -ProgressDots $true

$UriResource = "teams/$($TMS_KomunitaCPR_Id)/members"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
[array]$Members = Get-GraphOutputREST -Uri $Uri -ContentType $ContentTypeJSON -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -Text "$($TMS_KomunitaCPR_Id) team members" -ProgressDots $true
write-log "Team members (/teams/<id>/members): $($Members.count)" -ForegroundColor Cyan

[array]$TeamMembers = $AllUsers | Where-Object { $Members.userId -contains $_.id }
foreach ($TeamMember in $TeamMembers) {
    if ($CustomCityMapping_DB.ContainsKey($TeamMember.userPrincipalName)) {
        write-host "Custom city mapping for user: $($TeamMember.userPrincipalName) - City: $($CustomCityMapping_DB[$TeamMember.userPrincipalName])"
        $TeamMember.city = $CustomCityMapping_DB[$TeamMember.userPrincipalName]
    }
}
$TeamMembers = $TeamMembers | Where-Object { $null -ne $_.city }

$UriResource = "teams/$($TMS_KomunitaCPR_Id)/channels"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
[array]$Channels = Get-GraphOutputREST -Uri $Uri -ContentType $ContentTypeJSON -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
write-host "Channels (/teams/<id>/channels): $($Channels.count)" -ForegroundColor Cyan
foreach ($Channel in $Channels) {
    write-host "Channel: $($Channel.displayName) - ID: $($Channel.id)"
}

$AddCounter = 0
$RemoveCounter = 0

foreach ($Channel in $Channels) {
    foreach ($Key in $ChannelsToProcess.Keys) {
        #write-host $Key
        if ($Channel.displayName -like "$key*") {
            write-log "Processing channel: $($Channel.displayName) - Key: `"$($ChannelsToProcess[$Key])`""
            $UriResource = "teams/$($TMS_KomunitaCPR_Id)/channels/$($Channel.id)/members"
            $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
            [array]$ChannelMembers = Get-GraphOutputREST -Uri $Uri -ContentType $ContentTypeJSON -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
            [array]$ChannelOwners = $ChannelMembers | Where-Object { $_.roles -contains "owner" }
            [array]$ASISMembers = $AllUsers | Where-Object { $ChannelMembers.userId -contains $_.id }
            [array]$TOBEMembers = $TeamMembers | Where-Object { $_.city.startsWith($ChannelsToProcess[$Key]) }
            [array]$MembersToAdd = $TOBEMembers | Where-Object { $ASISMembers.id -notcontains $_.id }
            [array]$MembersToRemove = $ASISMembers | Where-Object { ($TOBEMembers.id -notcontains $_.id) -and (-not ($ChannelOwners.userId -contains $_.id)) -and (-not ($UserWhitelist -contains $_.id))}
            
            if ($MembersToAdd) {
                foreach ($Member in $MembersToAdd) {
                    $UriResource = "teams/$($TMS_KomunitaCPR_Id)/channels/$($Channel.id)/members"
                    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
                    $Body = @{
                        "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                        roles = @()
                        "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$($Member.id)')"
                    } | ConvertTo-Json
                    
                    Try {
                        $Response = Invoke-RestMethod -Uri $Uri -Method "POST" -Body $Body -ContentType $ContentTypeJSON -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders
                        write-log "$($Member.mail) $($Member.displayName) ($($Member.Id)) added to $($Channel.displayName)"
                        $AddCounter++
                    }
                    Catch {
                        write-log $_.Exception.Message -MessageType "Error"
                    }
                }
            }

            if ($MembersToRemove) {
                foreach ($Member in $ChannelMembers) {
                    if ($MembersToRemove.id -contains $Member.userId) {

                        $UriResource = "teams/$($TMS_KomunitaCPR_Id)/channels/$($Channel.id)/members/$($Member.id)"
                        $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
                        Try {
                            $Response = Invoke-RestMethod -Uri $Uri -Method "DELETE" -ContentType $ContentTypeJSON -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders
                            write-log "$($Member.email) $($Member.displayName) ($($Member.Id)) removed from $($Channel.displayName)" -ForegroundColor Red
                            $RemoveCounter++
                        }
                        Catch {
                            write-log $_.Exception.Message -MessageType "Error"
                        }
                    }
                }
            }
        }
    }
}

write-log "$($AddCounter) members added to channels"
write-log "$($RemoveCounter) members removed from channels"

#######################################################################################################################

. $IncFile_StdLogEndBlock

