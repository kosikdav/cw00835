#######################################################################################################################
# Get-Data-DB-Files-GRP
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
$LogFilePrefix		= "get-data-db-files-grp"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

[array]$DBReportGroupsAllMin        = $null
[array]$DBReportGroupsM365          = $null

[int]$ThrottlingDelayPerGroupInMsec = 200

$TTL = 30

#######################################################################################################################
. $IncFile_StdLogStartBlock
Write-Log "TTL: $($TTL) minutes"
Write-Log "ThrottlingDelayPerGroupInMsec: $($ThrottlingDelayPerGroupInMsec)"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
$UriResource = "groups"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999
[array]$AllAADGroups = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ProgressDots
Write-Log "Total groups: $($AllAADGroups.Count)"

foreach ($Group in $AllAADGroups) {
    Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
    $GroupOwnerString = [string]::Empty
    $GroupIsUnified = $GroupIsTeam = $false
    $MemberCount = $OwnerCount = "n/a"
    $resourceBehaviorOptions = $Group | Select-Object -ExpandProperty resourceBehaviorOptions
    if ($Group.securityEnabled) {
        $groupType = "Security"
    }
    else {
        $groupType = "Distribution"
    }
    if (($Group | Select-Object -ExpandProperty GroupTypes) -Contains "Unified") {
        $groupType = "Unified"
        $GroupIsUnified = $true
        if (($Group | Select-Object -ExpandProperty GroupTypes) -Contains "DynamicMembership") {
            $groupType = "UnifiedDynamicMembership"
        }    
    }
    $provisionedAs = $null
    if (($Group | Select-Object -ExpandProperty resourceProvisioningOptions) -Contains "Team") {
        $ProvisionedAs = "Team"
        $GroupIsTeam = $true
    }

    $UriResource = "groups/$($Group.id)/members/`$count"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
    $MemberCount = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ConsistencyLevel "eventual"

    $DBReportGroupsAllMin += [pscustomobject]@{
        id						= $Group.id
        displayName 			= $Group.displayName
        createdDateTime         = $Group.createdDateTime
        securityEnabled			= $Group.securityEnabled
        mailEnabled				= $Group.mailEnabled
        mail                    = $Group.mail
        onPremisesSyncEnabled   = $Group.onPremisesSyncEnabled
        MemberCount             = $MemberCount
    }
    
    if ($groupIsUnified -and (-not $groupIsTeam)) {
        $UriResource = "groups/$($Group.id)/owners"
        $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
        $GroupOwners = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
        $OwnerCount = $GroupOwners.Count
        if ($GroupOwners.Count -gt 0) {
            $GroupOwnerString = $GroupOwners.UserPrincipalName.ToLower() -join ";"
        }
    
        $DBReportGroupsM365 += [pscustomobject]@{
            id                              = $Group.id
            isArchived						= $Group.isArchived
            InternalId						= $Group.internalId
            DisplayName						= $Group.displayName
            createdDateTime					= $Group.createdDateTime
            createdByAppId					= $Group.createdbyAppId
            Mail							= $Group.Mail
            MailNickname					= $Group.MailNickname
            ProvisionedAs	 				= $ProvisionedAs
            Visibility						= $Group.Visibility
            GroupType						= $GroupType
            MemberCount                     = $MemberCount
            OwnerCount                      = $OwnerCount
            Owners                          = $GroupOwnerString

            ResourceBehaviorOptions			= $resourceBehaviorOptions
            ResourceProvisioningOptions		= $Group | Select-Object -ExpandProperty resourceProvisioningOptions

            AllowOnlyMembersToPost 			= ($resourceBehaviorOptions -Contains "AllowOnlyMembersToPost")
            HideGroupInOutlook 				= ($resourceBehaviorOptions -Contains "HideGroupInOutlook")
            SubscribeNewGroupMembers 		= ($resourceBehaviorOptions -Contains "SubscribeNewGroupMembers")
            WelcomeEmailDisabled 			= ($resourceBehaviorOptions -Contains "WelcomeEmailDisabled")
        }       
    }
    Start-Sleep -Milliseconds $ThrottlingDelayPerGroupInMsec
}

Export-Report "DBReportGroupsAllMin" -Report $DBReportGroupsAllMin -Path $DBFileGroupsAllMin -SortProperty "displayName"
Export-Report "DBReportGroupsM365" -Report $DBReportGroupsM365 -Path $DBFileGroupsM365 -SortProperty "mail"

. $IncFile_StdLogEndBlock
