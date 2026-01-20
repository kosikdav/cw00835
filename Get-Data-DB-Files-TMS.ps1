#######################################################################################################################
# Get-Data-DB-Files-TMS
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    [string]$workloads
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "db"
$LogFilePrefix		= "get-data-db-files-tms"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

[array]$DBReportTeams               = @()
[array]$DBReportTeamsChannelsOwners = @()
[hashtable]$TeamsUser_DB = @{}
[int]$ThrottlingDelayPerGroupInMsec = 500
[int]$ThrottlingDelayPerChannelInMsec = 300

$TTL = 10

#######################################################################################################################
. $IncFile_StdLogStartBlock
Write-Log "TTL: $($TTL) minutes"
Write-Log "ThrottlingDelayPerGroupInMsec: $($ThrottlingDelayPerGroupInMsec)"
Write-Log "ThrottlingDelayPerChannelInMsec: $($ThrottlingDelayPerChannelInMsec)"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL

$UriResource = "users"
$UriSelect = "id,userPrincipalName"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Select $UriSelect
[array]$AADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON

foreach ($User in $AADUsers) {
    $UserRecord = [pscustomobject]@{
        Id  						= $User.id
        UserPrincipalName 			= $User.userPrincipalName
    }
    $TeamsUser_DB.Add($User.id,$UserRecord)
}

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
$UriResource = "groups"
$UriFilter = "resourceProvisioningOptions/Any(x:x+eq+'Team')"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Filter $UriFilter
[array]$TeamGroups = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "Team groups" -ProgressDots

$TeamCounter = 0
foreach ($TeamGroup in $TeamGroups) {
    $TeamCounter++
    write-host "$('{0:d5}' -f $TeamCounter)/$('{0:d5}' -f $TeamGroups.Count)" -ForegroundColor Yellow -BackgroundColor Green -NoNewline
    write-host " $($TeamGroup.displayName)" -ForegroundColor Yellow
    $TeamFilesFolderURL = $MIPLabels = $TeamOwnerString = [string]::Empty
    $DynamicMembership = $false
    $UriResource = "groups/$($TeamGroup.id)"
    $UriSelect = "assignedLabels,allowExternalSenders,autoSubscribeNewMembers"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
    Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
    $TeamGroupExt = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -AppRegName $AppReg_LOG_READER -ContentType $ContentTypeJSON
    #MIP label
    if ($TeamGroupExt.assignedLabels) {
        $MIPLabels = $TeamGroupExt.assignedLabels.displayName.ToLower() -join ";"
    }
    #dynamic membership?
    if ($TeamGroup.GroupTypes.Contains("DynamicMembership")) {
        $DynamicMembership = $true
    }
    #get team instance
    $UriResource = "teams/$($TeamGroup.id)"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
    Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
    $Team = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -AppRegName $AppReg_LOG_READER -ContentType $ContentTypeJSON

    #team filesfolderURL
    $UriResource = "groups/$($TeamGroup.id)/sites/root/weburl"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
    Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
    $TeamFilesFolderURL = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -AppRegName $AppReg_LOG_READER -ContentType $ContentTypeJSON
    $TeamFilesFolderURL = ($TeamFilesFolderURL.Trim().ToLower()).TrimEnd("/")
    
    #team owners
    $UriResource = "groups/$($TeamGroup.id)/owners"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
    Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
    [array]$TeamOwners = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -AppRegName $AppReg_LOG_READER -ContentType $ContentTypeJSON
    if ($TeamOwners.Count -gt 0) {
        $TeamOwnerString = $TeamOwners.UserPrincipalName.ToLower() -join ";"
    }

    #add team to report DB TeamsChannelsOwners
    $DBReportTeamsChannelsOwners += [pscustomobject]@{
        TeamId              = $TeamGroup.id;
        TeamName		    = $TeamGroup.displayName;
        Mail		        = $TeamGroup.mail;
        Level               = "Team";
        InternalId          = $Team.internalId;
        TeamChannelId       = $TeamGroup.id + "_" + "team";
        ChannelName		    = "n/a";
        ChannelMail         = "n/a";
        FilesFolderUrl      = $TeamFilesFolderURL;

        CreatedDateTime     = $TeamGroup.createdDateTime;
        Sensitivity         = $MIPLabels;
        DynamicMembership   = $DynamicMembership;
        TeamVisibility      = $Team.visibility;
        isArchived          = $Team.isArchived;
        AllowExtSenders     = $TeamGroupExt.allowExternalSenders;
        AutoSubscribe       = $TeamGroupExt.autoSubscribeNewMembers;
        Owners              = $TeamOwnerString;
        TeamOwnerCount      = $TeamOwners.Count;
        TeamOwners          = $Team.summary.ownersCount;
        TeamMembers         = $Team.summary.membersCount;
        TeamGuests          = $Team.summary.guestsCount;

        ChannelId               = "n/a";
        ChannelDisplayName      = "n/a";
        ChannelMembershipType   = "n/a";
        ChannelOwnerCount       = "n/a";
        ChannelMemberCount      = "n/a"
    }
    #add team to report DB Teams
    $DBReportTeams += [pscustomobject]@{
        TeamId              = $TeamGroup.id;
        TeamName		    = $TeamGroup.displayName;
        Mail    		    = $TeamGroup.mail;
        InternalId          = $Team.internalId;
        FilesFolderUrl      = $TeamFilesFolderURL;
        
        CreatedDateTime     = $TeamGroup.createdDateTime;
        Sensitivity         = $MIPLabels;
        DynamicMembership   = $DynamicMembership;
        TeamVisibility      = $Team.visibility;
        isArchived          = $Team.isArchived;
        AllowExtSenders     = $TeamGroupExt.allowExternalSenders;
        AutoSubscribe       = $TeamGroupExt.autoSubscribeNewMembers;
        Owners              = $TeamOwnerString;
        TeamOwnerCount      = $TeamOwners.Count;
        TeamOwners          = $Team.summary.ownersCount;
        TeamMembers         = $Team.summary.membersCount;
        TeamGuests          = $Team.summary.guestsCount
    }

    <#
    $UriResource = "teams/$($TeamGroup.id)/incomingChannels"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
    Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
    $incomingChannels = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -includeUnknownEnumMembers
    if ($incomingChannels.Count -gt 0) {
        $incomingChannelIds = $incomingChannels.id
    }
    #>

    $UriResource = "teams/$($TeamGroup.id)/allChannels"
    $UriFilter = "membershipType+ne+'standard'"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Filter $UriFilter
    Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
    [array]$AllChannels = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -AppRegName $AppReg_LOG_READER -ContentType $ContentTypeJSON -includeUnknownEnumMembers
    if ($AllChannels.Count -gt 0) {
        ForEach ($Channel in $AllChannels) {
            if (-not $Channel.id) {
                continue
            }
            $AllChannelMembers = $AllChannelOwners = @()
            #channel filesfolderURL
            $ChannelFilesFolderURL = "no_URL"
            if ($Channel.filesFolderWebUrl) {
                $ChannelFilesFolderURL = $Channel.filesFolderWebUrl.Trim().ToLower()
                if ($ChannelFilesFolderURL.StartsWith($RootSPOURL)) {
                    $SiteString = ($ChannelFilesFolderURL.Replace($RootSPOURL+"/","")).Split("/")[0]
                    $ChannelFilesFolderURL = $RootSPOURL + "/" + $SiteString
                }
            }

            #channel owners
            $ChannelOwners = "no_owners"
            $UriResource = "teams/$($Team.id)/channels/$($Channel.id)/members"
            $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
            Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
            [array]$AllChannelMembers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -AppRegName $AppReg_LOG_READER -ContentType $ContentTypeJSON
            if ($AllChannelMembers.Count -gt 0) {
                foreach ($ChannelMember in $AllChannelMembers) {
                    if (-not $ChannelMember.userid) {
                        continue
                    }
                    $CurrentAADUser = $null
                    try {
                        $CurrentAADUser = $TeamsUser_DB[$ChannelMember.userid]
                        if ($CurrentAADUser.accountEnabled -and ($ChannelMember.Roles)) {
                            if ($ChannelMember.Roles -Contains("owner")) {
                                $AllChannelOwners += $ChannelMember
                                $ChannelOwners = $ChannelOwners + $CurrentAADUser.userPrincipalName.ToLower() + ";"
                            }
                        }
                    }
                    Catch {
                        Write-Log "Channel:$($Channel.displayName) --- member:$($ChannelMember.email) id:$($ChannelMember.userid) not found in TeamsUser_DB" -MessageType "ERR"
                    }
                }
                $ChannelOwners = $ChannelOwners.Trim(";")
            }

            #channel membership type
            $ChannelMembershipType = ($Channel.membershipType).ToLower()
            If ($Channel.id -in $IncomingChannelIds) {
                $ChannelMembershipType = "shared - incoming"
            }

            #add channel to report DB TeamsChannelsOwners
            $DBReportTeamsChannelsOwners += [pscustomobject]@{
                TeamId				= $TeamGroup.id;
                TeamName		    = $TeamGroup.displayName;
                Mail                = $TeamGroup.mail;
                Level               = "Channel";
                InternalId          = $Team.internalId;
                TeamChannelId       = $TeamGroup.id + "_" + $Channel.id;
                ChannelName		    = $Channel.displayName;
                ChannelMail		    = $Channel.email;
                FilesFolderUrl      = $ChannelFilesFolderURL;

                CreatedDateTime     = $Channel.createdDateTime;
                Sensitivity         = $MIPLabels;
                DynamicMembership   = $DynamicMembership;
                TeamVisibility      = $TeamGroup.visibility.ToLower();
                isArchived          = $Team.isArchived;
                AllowExtSenders     = $TeamGroupExt.allowExternalSenders;
                AutoSubscribe       = $TeamGroupExt.autoSubscribeNewMembers;
                Owners              = $ChannelOwners;
                TeamOwnerCount      = $TeamOwners.Count;
                TeamOwners          = $Team.summary.ownersCount;
                TeamMembers         = $Team.summary.membersCount;
                TeamGuests          = $Team.summary.guestsCount;

                ChannelId               = $Channel.id;
                ChannelDisplayName      = $Channel.displayName;
                ChannelMembershipType   = $ChannelMembershipType;
                ChannelOwnerCount       = $AllChannelOwners.Count;
                ChannelMemberCount      = $AllChannelMembers.Count
            }
            Start-Sleep -Milliseconds $ThrottlingDelayPerChannelInMsec
        }
    }
    Start-Sleep -Milliseconds $ThrottlingDelayPerGroupInMsec
}

Export-Report "DBReportTeams" -Report $DBReportTeams -Path $DBFileTeams -SortProperty "TeamName"
Export-Report "DBReportTeamsChannelsOwners" -Report $DBReportTeamsChannelsOwners -Path $DBFileTeamsChannelsOwners -SortProperty "TeamName"

. $IncFile_StdLogEndBlock
