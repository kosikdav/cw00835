param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    [string]$workloads
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "db"
$LogFilePrefix		= "get-data-db-files-pwr"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"


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

[array]$PwrEnvironments = Get-AdminPowerAppEnvironment

foreach ($env in $PwrEnvironments) {

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
}

#######################################################################################################################

Export-Report -Text "DBFilePwrEnvironments" -Report $ReportPwrEnvironments -Path $DBFilePwrEnvironments -SortProperty "displayName"

. $IncFile_StdLogEndBlock
