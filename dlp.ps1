$TenantId = "b233f9e1-5599-4693-9cef-38858fe25406"
$PwrEndPoint  = "prod"
#Add-PowerAppsAccount -TenantID $TenantId -Endpoint $PwrEndPoint
$policiesData = Get-DlpPolicy
$policies = $policiesData.value
write-host "Policies: $($policies.count)"
foreach ($policy in $policies) {
    write-host "Policy: $($policy.displayName)"
    write-host $policy.connectorGroups.count
    write-host $policy.environmentType
    write-host $policy.environments.count
}

$environments = Get-AdminPowerAppEnvironment
write-host "Environments: $($environments.count)"