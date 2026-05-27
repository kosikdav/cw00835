#######################################################################################################################
# Set-MaiboxProperties
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1
. $ScriptPath\include-Script-StdIncBlock.ps1

Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 120
$targetTenantId = "3687fd79-edff-4560-9dda-317079330262"
$appId = "bc272630-cf1e-43b5-998f-e576a39541c6"
$scope = "AIRPLUS_CROSS_TENANT_MIGRATION"
$orgrelname = "AIRPLUS_T2T_EXO_MIGRATION"
New-OrganizationRelationship $orgrelname -Enabled:$true -MailboxMoveEnabled:$true -MailboxMoveCapability RemoteOutbound -DomainNames $targetTenantId -OAuthApplicationId $appId -MailboxMovePublishedScopes $scope

