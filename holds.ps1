$targetUrl = "https://cezdata-my.sharepoint.com/personal/jindrich_houzva_cez_cz"

$cases = Get-ComplianceCase 
foreach ($case in $cases) {
        write-host $case.Name -ForegroundColor Green
        #write-host $case.CreatedDateTime -ForegroundColor Green
    <#
    write-host $case.Name -ForegroundColor Green
    write-host $case.CaseType
    write-host $case.ClosedBy    
    write-host $case.ClosedDateTime
    write-host $case.ClosingStatus 
    write-host $case.CreatedDateTime
    write-host $case.Description
    write-host $case.ExternalId
    write-host $case.Identity
    write-host $case.IsValid
    write-host $case.LastAccessTime
    write-host $case.LastModifiedBy
    write-host $case.LastModifiedDateTime
    write-host $case.ObjectState
    write-host $case.RecentItemId
    write-host $case.SecondaryCaseType
    write-host $case.SourceCaseType
    write-host $case.Sources
    write-host $case.Status
    write-host $case.TenantId
    #>
    if ($case.CreatedDateTime -gt (Get-Date).AddYears(-2)) {
        write-host "Case older than 2 years" -ForegroundColor Red 
    }
    #$CaseHoldPolicy = Get-CaseHoldPolicy -Case $case.Identity -DistributionDetail
    
}

