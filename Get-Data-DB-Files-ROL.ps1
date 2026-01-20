#######################################################################################################################
# Get-Data-DB-Files-ROL
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
$LogFilePrefix		= "get-data-db-files-rol"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

[array]$DBReportResourceActions     = $null
[array]$DBReportOAuthScopes         = $null
[array]$DBReportAppRoles            = $null
[array]$DBReportResourcePerms       = $null
[array]$DBReportAADAdminRoles       = $null

[int]$TTL = 30

#######################################################################################################################
. $IncFile_StdLogStartBlock

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
$UriResource  = "roleManagement/directory/resourceNamespaces/microsoft.directory/resourceActions"
$UriSelect = "id,actionVerb,description,isPrivileged,name,resourceScopeId"
$Uri = New-GraphUri -Version "beta" -Resource $UriResource -Select $UriSelect
[array]$ResourceActions = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
foreach ($ResourceAction in $ResourceActions) {
    $Namespace = $Entity = $PropertySet = $Action = [string]::Empty
    $Namespace = $ResourceAction.Name.Split("/")[0]
    $Entity = $ResourceAction.Name.Split("/")[1]
    if ($ResourceAction.Name -match "(?<=^[^/]+/[^/]+/).*(?=/[^/]*$)") {
        $PropertySet = $Matches[0]
    }
    if ($ResourceAction.Name -match "[^/]+$") {
        $Action = $Matches[0]
    }
    $DBReportResourceActions += [pscustomobject]@{
        Id				= $ResourceAction.Id
        Name			= $ResourceAction.name
        Namespace		= $Namespace
        Entity			= $Entity
        PropertySet		= $PropertySet
        Action			= $Action
        ActionVerb		= $ResourceAction.actionVerb
        Description		= $ResourceAction.description
        IsPrivileged 	= $ResourceAction.isPrivileged
        ResourceScopeId = $ResourceAction.resourceScopeId
    }
}

##############################################

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
$UriResource  = "servicePrincipals(appId='00000003-0000-0000-c000-000000000000')"
$UriSelect = "id,appId,displayName,appRoles,oauth2PermissionScopes,resourceSpecificApplicationPermissions"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
$SPZero = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
foreach ($Scope in $SPZero.oauth2PermissionScopes) {
    $DBReportOAuthScopes += [pscustomobject]@{
        id          = $Scope.id
        isEnabled   = $Scope.isEnabled
        type        = $Scope.type
        value       = $Scope.value
        adminConsentDisplayName = $Scope.adminConsentDisplayName
        adminConsentDescription = $Scope.adminConsentDescription
        userConsentDisplayName = $Scope.userConsentDisplayName
        userConsentDescription = $Scope.userConsentDescription
    }
}
foreach ($Role in $SPZero.appRoles) {
    $DBReportAppRoles += [pscustomobject]@{
        id          = $Role.id
        allowedMemberTypes = $Role.allowedMemberTypes -join ","
        isEnabled   = $Role.isEnabled
        origin      = $Role.origin
        value       = $Role.value
        displayName = $Role.displayName
        description = $Role.description
    }
}
foreach ($Permission in $SPZero.resourceSpecificApplicationPermissions) {
    $DBReportResourcePerms += [pscustomobject]@{
        id          = $Permission.id
        isEnabled   = $Permission.isEnabled
        value       = $Permission.value
        displayName = $Permission.displayName
        description = $Permission.description
    }
}

################################################

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL $TTL
$UriResource  = "roleManagement/directory/roleDefinitions"
$UriSelect = "id,description,displayName,isBuiltIn,isEnabled,resourceScopes,rolePermissions,templateId,version"
$Uri = New-GraphUri -Version "beta" -Resource $UriResource -Select $UriSelect
[array]$RoleDefinitions = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
foreach ($RoleDefinition in $RoleDefinitions) {
    $ExcludedResourceActions = $Condition = [string]::Empty
    if ($RoleDefinition.rolePermissions.excludedResourceActions) {
        foreach ($Item in $RoleDefinition.rolePermissions.excludedResourceActions) {
            if ($Item.Length -ge 1) {
                $ExcludedResourceActions += $Item + ","
            }
        }
        $ExcludedResourceActions = $ExcludedResourceActions.TrimEnd(",")
    }
    if ($RoleDefinition.rolePermissions.Condition.Count) {
        foreach ($Item in $RoleDefinition.rolePermissions.Condition) {
            if ($Item.Length -ge 1) {
                $Condition += $Item + ","
            }
        }
        $Condition = $Condition.TrimEnd(",")
    }

    if ($RoleDefinition.rolePermissions.allowedResourceActions) {
        foreach ($AllowedResourceAction in $RoleDefinition.rolePermissions.allowedResourceActions) {
            $Namespace = $Entity = $PropertySet = $Action = [string]::Empty
            $Namespace = $AllowedResourceAction.Split("/")[0]
            $Entity = $AllowedResourceAction.Split("/")[1]
            if ($AllowedResourceAction -match "(?<=^[^/]+/[^/]+/).*(?=/[^/]*$)") {
                $PropertySet = $Matches[0]
            }
            if ($AllowedResourceAction -match "[^/]+$") {
                $Action = $Matches[0]
            }
            $DBReportAADAdminRoles += [pscustomobject]@{
                Id			= $RoleDefinition.Id
                DisplayName	= $RoleDefinition.DisplayName
                IsBuilIn 	= $RoleDefinition.IsBuiltIn
                IsEnabled	= $RoleDefinition.IsEnabled
                isPrivileged = $RoleDefinition.isPrivileged
                TemplateId 	= $RoleDefinition.templateId
                Version 	= $RoleDefinition.version
                Namespace	= $Namespace
                Entity		= $Entity
                PropertySet	= $PropertySet
                Action		= $Action
                Permission	= $AllowedResourceAction
                ExcludedResourceActions = $ExcludedResourceActions
                Condition	= $Condition
            }
        }
    }
    else {
        $DBReportAADAdminRoles += [pscustomobject]@{
            Id			= $RoleDefinition.Id
            DisplayName	= $RoleDefinition.DisplayName
            IsBuilIn 	= $RoleDefinition.IsBuiltIn
            IsEnabled	= $RoleDefinition.IsEnabled
            isPrivileged = $RoleDefinition.isPrivileged
            TemplateId 	= $RoleDefinition.templateId
            Version 	= $RoleDefinition.version
            Namespace	= "n/a"
            Entity		= "n/a"
            PropertySet	= "n/a"
            Action		= "n/a"
            Permission	= "none"
            ExcludedResourceActions = $ExcludedResourceActions
            Condition	= $Condition
        }
    }
}

Export-Report -Text "DBReportResourceActions" -Report $DBReportResourceActions -Path $DBFileAADResourceActions -SortProperty "name"
Export-Report -Text "DBReportOAuthScopes" -Report $DBReportOAuthScopes -Path $DBFileAADOAuthScopes -SortProperty "value"
Export-Report -Text "DBReportAppRoles" -Report $DBReportAppRoles -Path $DBFileAADAppRoles -SortProperty "value"
Export-Report -Text "DBReportResourcePerms" -Report $DBReportResourcePerms -Path $DBFileAADResourcePerms -SortProperty "value"
Export-Report -Text "DBReportAADAdminRoles" -Report $DBReportAADAdminRoles -Path $DBFileAADAdmRoles -SortProperty "displayName"

. $IncFile_StdLogEndBlock
