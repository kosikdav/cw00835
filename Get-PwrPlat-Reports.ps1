param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################
$LogFolder			= "exports"
$LogFilePrefix		= "aad-groups-reports"
$OutputFolder		= "power-platform\reports"
$OutputFilePrefix	= "pwrplat"

$OutputFileSuffixPwrEnv     = "environments"
$OutputFileSuffixPwrApps    = "pwrapps"
$OutputFileSuffixFlows	    = "flows"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputFileEnvironments = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixPwrEnv -Ext "csv"
$OutputFilePwrApps      = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixPwrApps -Ext "csv"
$OutputFileFlows        = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixFlows -Ext "csv"

#######################################################################################################################
. $IncFile_StdLogStartBlock

. $IncFile_AppReg_POWERPLAT_MGMT

try {
    Add-PowerAppsAccount -Endpoint $PwrEndpoint -TenantID $TenantId -ApplicationId $ClientId -CertificateThumbprint $CertficateThumbprint
}
Catch {
    write-log "Error: $_"
    write-log "Error: $_.Exception.Message"
}

[array]$ReportPwrEnvironments = @()
[array]$ReportPwrApps = @()
[array]$ReportFlows = @()

[array]$PwrEnvironments = Get-AdminPowerAppEnvironment

foreach ($env in $PwrEnvironments) {
    Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 120

    [array]$Apps = Get-AdminPowerApp -EnvironmentName $env.EnvironmentName
    [array]$Flows = Get-AdminFlow -EnvironmentName $env.EnvironmentName -IncludeDeleted $true -IncludeEUDBNonCompliantFlows $true

    $ReportPwrEnvironments += [PSCustomObject]@{
        EnvironmentName = $env.EnvironmentName
        DisplayName = $env.DisplayName
        Description = $env.Description
        IsDefault = $env.IsDefault
        Location = $env.Location
        CreatedTime = $env.CreatedTime
        CreatedBy = $env.CreatedBy.UserPrincipalName
        CreatedById = $env.CreatedBy.Id
        CreatedByType = $env.CreatedBy.Type
        LastModifiedTime = $env.LastModifiedTime
        LastModifiedBy = $env.LastModifiedBy
        CreationType = $env.CreationType 
        EnvironmentType = $env.EnvironmentType
        CommonDataServiceDatabaseProvisioningState = $env.CommonDataServiceDatabaseProvisioningState
        CommonDataServiceDatabaseType   = $env.CommonDataServiceDatabaseType
        InternalCds = $env.InternalCds
        OrganizationId = $env.OrganizationId
        SecurityGroupId = $env.SecurityGroupId
        RetentionPeriod = $env.RetentionPeriod
        AppsCount = $Apps.Count
        FlowsCount = $Flows.Count
    }

    foreach ($app in $Apps) {
        $ReportPwrApps += [PSCustomObject]@{
            EnvironmentName = $env.EnvironmentName
            EnvironmentDisplayName = $env.DisplayName
            #description = $app.internal.properties.description
            AppName = $app.AppName
            DisplayName = $app.DisplayName
            appVersion = $app.internal.properties.appVersion
            Owner = $app.owner.UserPrincipalName
            OwnerId = $app.owner.Id
            createdBy = $app.internal.properties.createdBy.userPrincipalName
            lastModifiedBy = $app.internal.properties.lastModifiedBy.userPrincipalName
            #additionalAuthors = $app.internal.properties.additionalAuthors

            CreatedTime = $app.CreatedTime
            LastModifiedTime = $app.LastModifiedTime
            
            AppId = $app.AppId
            AppTemplateId = $app.AppTemplateId
            Internalid = $app.Internal.id
            #type = $app.Internal.type
            logicalName = $app.Internal.logicalName
            appLocation = $app.Internal.appLocation
            isAppComponentLibrary = $app.Internal.isAppComponentLibrary
            appType = $app.Internal.appType
            primaryDeviceWidth = $app.internal.tags.primaryDeviceWidth
            primaryDeviceHeigh = $app.internal.tags.primaryDeviceHeight
            supportsPortrait = $app.internal.tags.supportsPortrait
            supportsLandscape = $app.internal.tags.supportsLandscape
            primaryFormFactor = $app.internal.tags.primaryFormFactor
            showStatusBar = $app.internal.tags.showStatusBar
            publisherVersion = $app.internal.tags.publisherVersion
            minimumRequiredApiVersion = $app.internal.tags.minimumRequiredApiVersion
            hasComponent = $app.internal.tags.hasComponent
            hasUnlockedComponent = $app.internal.tags.hasUnlockedComponent
            isUnifiedRootApp = $app.internal.tags.isUnifiedRootApp
            sienaVersion = $app.internal.tags.sienaVersion

            sharedGroupsCount = $app.internal.properties.sharedGroupsCount
            sharedUsersCount = $app.internal.properties.sharedUsersCount
            appOpenUri = $app.internal.properties.appOpenUri
            userAppMetadatafavorite = $app.internal.properties.userAppMetadata.favorite
            userAppMetadataincludeInAppsList = $app.internal.properties.userAppMetadata.includeInAppsList
            isFeaturedApp = $app.internal.properties.isFeaturedApp
            bypassConsent = $app.internal.properties.bypassConsent
            isHeroApp = $app.internal.properties.isHeroApp
            almMode = $app.internal.properties.almMode
            performanceOptimizationEnabled = $app.internal.properties.performanceOptimizationEnabled
            canConsumeAppPass = $app.internal.properties.canConsumeAppPass
            enableModernRuntimeMode = $app.internal.properties.enableModernRuntimeMode
            isTeamsOnly = $app.internal.properties.executionRestrictions.isTeamsOnly
            DLPstatus = $app.internal.properties.executionRestrictions.dataLossPreventionEvaluationResult.status
            DLPLastEvaluationDate = $app.internal.properties.executionRestrictions.dataLossPreventionEvaluationResult.lastEvaluationDate
            appPlanClassification = $app.internal.properties.appPlanClassification
            usesPremiumApi = $app.internal.properties.usesPremiumApi
            usesOnlyGrandfatheredPremiumApis = $app.internal.properties.usesOnlyGrandfatheredPremiumApis
            usesCustomApi = $app.internal.properties.usesCustomApi
            usesOnPremiseGateway = $app.internal.properties.usesOnPremiseGateway
            usesPcfExternalServiceUsage = $app.internal.properties.usesPcfExternalServiceUsage
            isCustomizable = $app.internal.properties.isCustomizable
            chatPaneCopilotEnabled = $app.internal.properties.chatPaneCopilotEnabled
            draftingCopilotEnabled = $app.internal.properties.draftingCopilotEnabled
            canvasGalleryFilteringCopilotEnabled = $app.internal.properties.canvasGalleryFilteringCopilotEnabled
            appVersionSource = $app.internal.properties.appVersionSource
        }
    }

    foreach ($flow in $Flows) {
        $createdByUPN = [string]::Empty
        if ($flow.CreatedBy.userId)        {
            $createdByUPN = (Get-GraphUserById -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -Id $flow.CreatedBy.userId).UserPrincipalName
        }

        $ReportFlows += [PSCustomObject]@{
            FlowName = $flow.FlowName
            Enabled = $flow.Enabled
            DisplayName = $flow.DisplayName
            UserType = $flow.UserType
            CreatedTime = $flow.CreatedTime
            LastModifiedTime = $flow.LastModifiedTime
            CreatedById = $flow.CreatedBy.userId
            CreatedBy = $createdByUPN
            WorkflowEntityId = $flow.WorkflowEntityId
            id = $flow.Internal.id
            apiId = $flow.Internal.Properties.piId
            state = $flow.Internal.Properties.state
            flowSuspensionReason = $flow.Internal.Properties.flowSuspensionReason
            flowFailureAlertSubscribed = $flow.Internal.Properties.flowFailureAlertSubscribed
            isManaged = $flow.Internal.Properties.isManaged
            isConsequential = $flow.Internal.Properties.flowOpenAiData.isConsequential
            isConsequentialFlagOverwritten = $flow.Internal.Properties.flowOpenAiData.isConsequentialFlagOverwritten
        }
    }   
}

#######################################################################################################################
write-host "done"

Export-Report -Text "Environments" -Report $ReportPwrEnvironments -Path $OutputFileEnvironments -SortProperty "displayName"
Export-Report -Text "Apps" -Report $ReportPwrApps -Path $OutputFilePwrApps -SortProperty "displayName"
Export-Report -Text "Flows" -Report $ReportFlows -Path $OutputFileFlows -SortProperty "displayName"

. $IncFile_StdLogEndBlock
