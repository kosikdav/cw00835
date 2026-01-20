
$AzSubReportFile = "d:\exports\cezdata\az-subs\az.csv" 

[array]$AzSubReport = @()

$subscriptions = Get-AzSubscription
foreach ($subscription in $subscriptions) {
    
    $Tag_Owner = $null
    $Tag_Department = $null
    $Tag_Environment = $null
    $Tag_Cost_Center = $null
    $Tag_Cost_Management = $null
    
    $Tags = $subscription.Tags
    if ($Tags.ContainsKey("Owner")) {
        $Tag_Owner = $Tags["Owner"].Trim()
    }
    if ($Tags.ContainsKey("Department")) {
        $Tag_Department = $Tags["Department"].Trim()
    }
    if ($Tags.ContainsKey("Environment")) {
        $Tag_Environment = $Tags["Environment"].Trim()
    }
    if ($Tags.ContainsKey("Cost Center")) {
        $Tag_Cost_Center = $Tags["Cost Center"].Trim()
    }
    if ($Tags.ContainsKey("Cost Management")) {
        $Tag_Cost_Management = $Tags["Cost Management"].Trim()
    }

    $Record = [PSCustomObject]@{
        Id = $subscription.id;
        Name = $subscription.name;
        State = $subscription.state;
        TenantId = $subscription.tenantId;
        Tag_Owner = $Tag_Owner;
        Tag_Department = $Tag_Department;
        Tag_Environment = $Tag_Environment;
        Tag_Cost_Center = $Tag_Cost_Center;
        Tag_Cost_Management = $Tag_Cost_Management
    }
    $AzSubReport += $Record
}

$AzSubReport | Export-Csv -Path $AzSubReportFile -NoTypeInformation -Encoding utf8