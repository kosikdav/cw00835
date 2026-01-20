#######################################################################################################################
# Get-EXO-Role-Assignments
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

$OutputFolder       = "exo\reports"
$OutputFilePrefix	= "exo"
$OutputFileSuffix	= "role-assignments"

. $ScriptPath\include-Script-StdIncBlock.ps1

[hashtable]$AADSPDB_ByObjectId = @{}
[hashtable]$EXOMgmtScopeDB = @{}

$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffix -Ext "csv"

[System.Collections.ArrayList]$EXORAReport = @()

#######################################################################################################################

. $ScriptPath\include-Script-StartLog-Generic.ps1

Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 120
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

$UriResource = "servicePrincipals"
$UriSelect = "id,displayName,appId"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
$AADServicePrincipals = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
foreach ($ServicePrincipal in $AADServicePrincipals) {
    $Record = [PSCustomObject]@{
        id = $ServicePrincipal.id
        DisplayName = $ServicePrincipal.DisplayName
        AppId = $ServicePrincipal.AppId
    }
    $AADSPDB_ByObjectId.Add($ServicePrincipal.Id, $Record)
}

#############################################################################################################################
#EXO service principal

$EXOMgmtScopes = Get-ManagementScope
$EXOMgmtRoleAssignments = Get-ManagementRoleAssignment

foreach ($EXOMgmtScope in $EXOMgmtScopes) {
    $Group = $null
    if ($EXOMgmtScope.RecipientFilter) {
        $GroupDN = ($EXOMgmtScope.RecipientFilter.TrimStart("MemberOfGroup -eq '")).TrimEnd("'")
        $Group = Get-DistributionGroup -Identity $GroupDN
    }
    $MgmtScopeObject = [PSCustomObject]@{
        Id = $EXOMgmtScope.Id
        Name = $EXOMgmtScope.Name
        Filter = $EXOMgmtScope.Filter
        RecipientFilter = $EXOMgmtScope.RecipientFilter
        FilterGroupName = $Group.DisplayName
        FilterGroupAADId = $Group.ExternalDirectoryObjectId
    }
    $EXOMgmtScopeDB.Add($EXOMgmtScope.Id, $MgmtScopeObject)
}

foreach ($EXOMgmtRoleAssignment in $EXOMgmtRoleAssignments) {
    $AADAppRecord = $MgmtScopeRecord = $FilterGroupMembers = $null
    
    if ($EXOMgmtRoleAssignment.RoleAssigneeType -eq "ServicePrincipal" -and $AADSPDB_ByObjectId.ContainsKey($EXOMgmtRoleAssignment.App)) {
        $AADAppRecord = $AADSPDB_ByObjectId[$EXOMgmtRoleAssignment.App]
    }
    if ($EXOMgmtRoleAssignment.CustomResourceScope -and $EXOMgmtScopeDB.ContainsKey($EXOMgmtRoleAssignment.CustomResourceScope)) {
        $MgmtScopeRecord = $EXOMgmtScopeDB[$EXOMgmtRoleAssignment.CustomResourceScope]
        $FilterGroupMembers = Get-GroupMembersFromGraphById -id $MgmtScopeRecord.FilterGroupAADId -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
    }
    $RAObject = [PSCustomObject]@{
        Id = $EXOMgmtRoleAssignment.Id
        #Identity = $EXOMgmtRoleAssignment.Identity
        #Guid = $EXOMgmtRoleAssignment.Guid
        ExchangeObjectId = $EXOMgmtRoleAssignment.ExchangeObjectId
        Enabled = $EXOMgmtRoleAssignment.Enabled
        IsValid = $EXOMgmtRoleAssignment.IsValid
        WhenChangedUTC = $EXOMgmtRoleAssignment.WhenChangedUTC
        WhenCreatedUTC = $EXOMgmtRoleAssignment.WhenCreatedUTC     
        RoleAssigneeType = $EXOMgmtRoleAssignment.RoleAssigneeType
        Name = $EXOMgmtRoleAssignment.Name
        RoleAssignee = $EXOMgmtRoleAssignment.RoleAssignee
        #RoleAssigneeName = $EXOMgmtRoleAssignment.RoleAssigneeName
        #DataObject = $EXOMgmtRoleAssignment.DataObject
        #User = $EXOMgmtRoleAssignment.User  
        EffectiveUserName = $EXOMgmtRoleAssignment.EffectiveUserName                  
        App = $EXOMgmtRoleAssignment.App
        AppName = $AADAppRecord.DisplayName
        AppId = $AADAppRecord.AppId
        AssignmentMethod = $EXOMgmtRoleAssignment.AssignmentMethod
        Role = $EXOMgmtRoleAssignment.Role                
        RoleAssignmentDelegationType = $EXOMgmtRoleAssignment.RoleAssignmentDelegationType
        CustomResourceScope = $EXOMgmtRoleAssignment.CustomResourceScope
        ScopeFilter = $MgmtScopeRecord.Filter
        FilterGroupName = $MgmtScopeRecord.FilterGroupName
        FilterGroupAADId = $MgmtScopeRecord.FilterGroupAADId
        FilterGroupMembers = $FilterGroupMembers.Count
        RecipientReadScope = $EXOMgmtRoleAssignment.RecipientReadScope
        RecipientWriteScope = $EXOMgmtRoleAssignment.RecipientWriteScope
        ConfigReadScope = $EXOMgmtRoleAssignment.ConfigReadScope
        ConfigWriteScope = $EXOMgmtRoleAssignment.ConfigWriteScope
        CustomRecipientWriteScope = $EXOMgmtRoleAssignment.CustomRecipientWriteScope
        CustomConfigWriteScope = $EXOMgmtRoleAssignment.CustomConfigWriteScope                 
    }
    $EXORAReport += $RAObject
}

#######################################################################################################################

Export-Report -Text "EXO role assignments" -Report $EXORAReport -SortProperty "Name" -Path $OutputFile

. $ScriptPath\include-Script-EndLog-generic.ps1
