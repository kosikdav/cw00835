$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1

$LogFolder			    = "exports"
$LogFilePrefix		    = "pwrplatmgmt"

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

write-log "Start script $ScriptName"


. d:\scripts\cezdata\include-appreg-CEZDATA_POWERPLAT_MGMT.ps1

try {
    Add-PowerAppsAccount -Endpoint $PwrEndpoint -TenantID $TenantId -ApplicationId $ClientId -CertificateThumbprint $CertficateThumbprint
}
Catch {
    write-log "Error: $_"
    write-log "Error: $_.Exception.Message"
}

$envReport = @()
$environments = Get-AdminPowerAppEnvironment
foreach ($env in $environments) {
    $envReport += [PSCustomObject]@{
        EnvironmentName = $env.EnvironmentName
        DisplayName = $env.DisplayName
        Description = $env.Description
        IsDefault = $env.IsDefault
        Location = $env.Location
        CreatedTime = $env.CreatedTime
        CreatedById = $env.CreatedBy.Id
        CreatedByDisplayName = $env.CreatedBy.DisplayName
        CreatedByUserPrincipalName = $env.CreatedBy.UserPrincipalName
        CreatedByEmail = $env.CreatedBy.Email
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
    }
}
$envReport | Export-Csv -Path "$ROF\$LogFilePrefix-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv" -NoTypeInformation -Encoding UTF8
