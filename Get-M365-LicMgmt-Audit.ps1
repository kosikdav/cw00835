#######################################################################################################################
# Get-M365-Licmgmt-Audit
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "lic-mgmt-M365"
$LogFilePrefix		= "m365-licmgm-audit"

$OutputFolder		= "lic-mgmt-m365"
$OutputFilePrefix	= "license-update-errors"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -FileDateYesterday -Freq "Y" -Ext "csv"

[array]$AuditLogEventReport = @()

$start  = $(Get-Date).AddDays(-7).ToString("yyyy-MM-dd") + "T00:00:00Z"

#######################################################################################################################

. $IncFile_StdLogStartBlock
write-log "Start time: $($start)"

$SKU_DB = Import-CSVtoHashDB -Path $DBFileLicensingInfoSKUs -KeyName "skuId"

If (Test-Path -Path $OutputFile) {
    $ExistingEvents_DB = Import-CSVtoHashDB -Path $OutputFile -KeyName "id"
}
else {
    $ExistingEvents_DB = @{}
}

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "auditLogs/directoryAudits"
$UriFilter1 = "loggedByService+eq+'Core Directory'"
$UriFilter2 = "category+eq+'UserManagement'"
$UriFilter3 = "activityDisplayName+eq+'Change user license'"
$UriFilter4 = "Result+eq+'Failure'"
$UriFilter5 = "activityDateTime+ge+$($start)"
$UriFilter = $UriFilter1 , $UriFilter2 , $UriFilter3 , $UriFilter4 , $UriFilter5 -join $UriAnd
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Filter $UriFilter
$AuditLogEventsGrp = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON

Write-Log "Audit log events (Update user license) found: $($AuditLogEventsGrp.Count)"

foreach ($AuditLogEvent in $AuditLogEventsGrp) {
    if ($ExistingEvents_DB.ContainsKey($AuditLogEvent.id)) {
        continue
    }
    else {
        $ErrorReason = $AuditLogEvent.additionalDetails.value
        if ($ErrorReason.Contains("Not enough licenses are available to complete this operation")) {
            $SKUId = $ErrorReason.Substring($ErrorReason.IndexOf("for SKU [") + 9, 36)
            $ErrorSKU = $SKU_DB[$SKUId].skuDisplayName
        }
        else {
            $ErrorSKU = [string]::Empty
        }
        $ReportObject = [pscustomobject]@{
            ActivityDateTime    = $AuditLogEvent.activityDateTime
            Result              = $AuditLogEvent.result
            User                = $AuditLogEvent.targetResources.userPrincipalName
            UserId              = $AuditLogEvent.targetResources.id
            Error               = $AuditLogEvent.additionalDetails.key 
            ErrorSKU            = $ErrorSKU
            ErrorReason         = $AuditLogEvent.additionalDetails.value
            Id                  = $AuditLogEvent.id
            CorellationId       = $AuditLogEvent.correlationId
            LoggedByService     = $AuditLogEvent.loggedByService
            Category            = $AuditLogEvent.category
            OperationType       = $AuditLogEvent.operationType
            ActivityDisplayName = $AuditLogEvent.activityDisplayName

        }
        $AuditLogEventReport += $ReportObject
    }
}

Export-Report "AAD groups audit report" -Report $AuditLogEventReport -Path $OutputFile -SortProperty "ActivityDateTime" -Append:$True

#######################################################################################################################

. $IncFile_StdLogEndBlock
