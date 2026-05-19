
#######################################################################################################################
# Get-MIP-Backup
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Init.ps1

#######################################################################################################################

$LogFolder 					= "mip-backup"
$LogFilePrefix				= "mip-backup"

$OutputFolder 				= "mip-backup"
$OutputFilePrefix			= "mip-backup"

$OutputFileSuffixLabels	    = "labels"
$OutputFileSuffixPolicies   = "policies"

#######################################################################################################################

. $ScriptPath\include-Script-Start-Include.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$OutputXMLFileLabels = New-OutputFile -RootFolder $OLF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixLabels -Ext "xml"
$OutputXMLFilePolicies = New-OutputFile -RootFolder $OLF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixPolicies -Ext "xml"
$OutputFileLabels = New-OutputFile -RootFolder $OLF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixLabels -Ext "csv"
$OutputFilePolicies = New-OutputFile -RootFolder $OLF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $OutputFileSuffixPolicies -Ext "csv"

[array]$ReportLabels = @()
[array]$ReportPolicies = @()

#######################################################################################################################

. $IncFile_StdLogStartBlock

. "$ScriptPath\cezdata\include-appreg-CEZDATA_PURVIEW_MGMT.ps1"
Connect-IPPSSession -AppId $ClientId -CertificateThumbprint $Thumbprint -Organization $TenantName

$Labels = Get-Label
write-host "MIP labels found: " $Labels.count -foregroundColor Yellow
foreach($Label in $Labels){
    $LabelObject = [pscustomobject]@{
        Name = $Label.Name
        DisplayName = $Label.DisplayName
        Priority = $Label.Priority
        Identity = $Label.Identity
        ContentType = $Label.ContentType
        Workload = $Label.Workload
        Policy = $Label.Policy
        AbacEnabled = $Label.AbacEnabled
        CanonicalLabelId = $Label.CanonicalLabelId
        Capabilities = $Label.Capabilities
        ColumnAssetCondition = $Label.ColumnAssetCondition
        Comment = $Label.Comment
        Conditions = $Label.Conditions
        Id = $Label.Id
        ContentTypeRemovalViolations = $Label.ContentTypeRemovalViolations
        CreatedBy = $Label.CreatedBy
        DefaultContentLabel = $Label.DefaultContentLabel
        Disabled = $Label.Disabled
        DistinguishedName = $Label.DistinguishedName
        ExchangeObjectId = $Label.ExchangeObjectId
        ExchangeVersion = $Label.ExchangeVersion
        ExternalIdentity = $Label.ExternalIdentity
        Guid = $Label.Guid
        ImmutableId = $Label.ImmutableId
        InformationProtectionAttributeRequired = $Label.InformationProtectionAttributeRequired
        InheritToChildItems = $Label.InheritToChildItems
        IsLabelGroup = $Label.IsLabelGroup
        IsParent = $Label.IsParent
        IsValid = $Label.IsValid
        LabelActions = $Label.LabelActions
        LastModifiedBy = $Label.LastModifiedBy
        LocaleSettings = $Label.LocaleSettings
        Mode = $Label.Mode
        MTOOwnerLabelId = $Label.MTOOwnerLabelId
        MTOOwnerTenantId = $Label.MTOOwnerTenantId
        ObjectCategory = $Label.ObjectCategory
        ObjectClass = $Label.ObjectClass
        ObjectState = $Label.ObjectState
        ObjectVersion = $Label.ObjectVersion
        OrganizationalUnitRoot = $Label.OrganizationalUnitRoot
        OrganizationId = $Label.OrganizationId
        OriginatingServer = $Label.OriginatingServer
        ParentId = $Label.ParentId
        ParentLabelDisplayName = $Label.ParentLabelDisplayName
        ReadOnly = $Label.ReadOnly
        SchematizedDataCondition = $Label.SchematizedDataCondition
        Settings = $Label.Settings
        Tooltip = $Label.Tooltip
        WhenChanged = $Label.WhenChanged
        WhenCreated = $Label.WhenCreated
    }
    $ReportLabels += $LabelObject
}

$Policies = Get-LabelPolicy
write-host "MIP policies found: " $Policies.count -foregroundColor Yellow
foreach($Policy in $Policies){
    write-host "Policy: " $Policy.Name -foregroundColor Cyan
    $PolicyObject = [pscustomobject]@{
        Name = $Policy.Name
        Priority = $Policy.Priority
        Enabled = $Policy.Enabled
        Id = $Policy.Id
        Workload = $Policy.Workload
        Type = $Policy.Type          
        CreationTimeUtc = $Policy.CreationTimeUtc
        DistinguishedName = $Policy.DistinguishedName
        DistributionResults = $Policy.DistributionResults
        DistributionStatus = $Policy.DistributionStatus
        DistributionSyncStatus = $Policy.DistributionSyncStatus
        Comment = $Policy.Comment                         
        CreatedBy = $Policy.CreatedBy 
        EndpointDlpAdaptiveScopes = $Policy.EndpointDlpAdaptiveScopes
        EndpointDlpAdaptiveScopesException = $Policy.EndpointDlpAdaptiveScopesException
        ErrorMetadata = $Policy.ErrorMetadata
        ExchangeAdaptiveScopes = $Policy.ExchangeAdaptiveScopes
        ExchangeAdaptiveScopesException = $Policy.ExchangeAdaptiveScopesException
        ExchangeLocation = $Policy.ExchangeLocation
        ExchangeLocationException = $Policy.ExchangeLocationException
        ExchangeObjectId = $Policy.ExchangeObjectId
        ExchangeVersion = $Policy.ExchangeVersion
        ExternalIdentity = $Policy.ExternalIdentity
        ForceValidate = $Policy.ForceValidate
        GlobalListType = $Policy.GlobalListType
        Guid = $Policy.Guid
        
        Identity = $Policy.Identity
        IsValid = $Policy.IsValid
        Labels = $Policy.Labels
        LastModifiedBy = $Policy.LastModifiedBy
        LastStatusUpdateTime = $Policy.LastStatusUpdateTime
        Locations = $Policy.Locations
        Mode = $Policy.Mode
        ModernGroupLocation = $Policy.ModernGroupLocation
        ModernGroupLocationException = $Policy.ModernGroupLocationException
        ModificationTimeUtc = $Policy.ModificationTimeUtc
        
        ObjectCategory = $Policy.ObjectCategory
        ObjectClass = $Policy.ObjectClass
        ObjectState = $Policy.ObjectState
        ObjectVersion = $Policy.ObjectVersion
        OneDriveAdaptiveScopes = $Policy.OneDriveAdaptiveScopes
        OneDriveAdaptiveScopesException = $Policy.OneDriveAdaptiveScopesException
        OneDriveLocation = $Policy.OneDriveLocation
        OneDriveLocationException = $Policy.OneDriveLocationException
        OrganizationalUnitRoot = $Policy.OrganizationalUnitRoot
        OrganizationId = $Policy.OrganizationId
        OriginatingServer = $Policy.OriginatingServer
        PolicyConstraints = $Policy.PolicyConstraints
        PolicyRBACScopes = $Policy.PolicyRBACScopes
        PolicyRulesMetaData = $Policy.PolicyRulesMetaData
        PolicySettingsBlob = $Policy.PolicySettingsBlob
        
        PublicFolderLocation = $Policy.PublicFolderLocation
        ReadOnly = $Policy.ReadOnly
        ScopedLabels = $Policy.ScopedLabels
        Settings = $Policy.Settings
        SharePointAdaptiveScopes = $Policy.SharePointAdaptiveScopes
        SharePointAdaptiveScopesException = $Policy.SharePointAdaptiveScopesException
        SharePointLocation = $Policy.SharePointLocation
        SharePointLocationException = $Policy.SharePointLocationException
        SkypeLocation = $Policy.SkypeLocation
        SkypeLocationException = $Policy.SkypeLocationException
        TeamsAdaptiveScopes = $Policy.TeamsAdaptiveScopes
        TeamsAdaptiveScopesException = $Policy.TeamsAdaptiveScopesException
        
        UPELabelRules = $Policy.UPELabelRules
        UserAdministrativeUnitMembershipMap = $Policy.UserAdministrativeUnitMembershipMap
        WhenChanged = $Policy.WhenChanged
        WhenCreated = $Policy.WhenCreated
    }
    $ReportPolicies += $PolicyObject
}

Export-Clixml -InputObject $Labels -Path $OutputXMLFileLabels
Export-Clixml -InputObject $Policies -Path $OutputXMLFilePolicies

Export-Report -Text "MIP labels report" -Report $ReportLabels -Path $OutputFileLabels -SortProperty "Priority"
Export-Report -Text "MIP policies report" -Report $ReportPolicies -Path $OutputFilePolicies -SortProperty "Priority"

#######################################################################################################################

. $IncFile_StdLogEndBlock
