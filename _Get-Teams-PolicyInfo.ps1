
$TeamsUsers = Get-CsOnlineUser | Select-Object DisplayName,Identity,UserPrincipalName,SipAddress,Enabled,WindowsEmailAddress,LineURI,HostedVoiceMail,OnPremEnterpriseVoiceEnabled,OnPremLineURI,SipProxyAddress,OnlineDialinConferencingPolicy,TeamsUpgradeEffectiveMode,TeamsUpgradePolicy,HostingProvider

$UserInfoReport = @()
$ProgressCount = 0

Foreach ($User in $TeamsUsers) {
	$Progresscount++
	$ProgressPct =  [int](($ProgressCount/$TeamsUsers.Count)*100)
	Write-Progress -Activity "Building teams policy statistics" -Status "$($ProgressPct)% complete:" -PercentComplete $ProgressPct
    
    Write-Host "Querying policy information for: " $User.Identity -ForegroundColor Green
    $UserPolicies = Get-CsUserPolicyAssignment -Identity $User.Identity

    $TenantDefaultString        =  "Tenant Default"     
    $TeamsMeetingPolicy         = $TenantDefaultString    
    $TeamsMessagingPolicy       = $TenantDefaultString    
    $TeamsAppSetupPolicy        = $TenantDefaultString    
    $TeamsAppPermissionsPolicy  = $TenantDefaultString    
    $TeamsEncryptionPolicy      = $TenantDefaultString    
    $TeamsUpdatePolicy          = $TenantDefaultString    
    $TeamsChannelsPolicy        = $TenantDefaultString   
    $TeamsFeedbackPolicy        = $TenantDefaultString    
    $TeamsLiveEventsPolicy      = $TenantDefaultString    
    
    If ($User.TeamsMeetingPolicy) {$TeamsMeetingPolicy = $User.TeamsMeetingPolicy}    
    If ($User.TeamsMessagingPolicy) {$TeamsMessagingPolicy = $User.TeamsMessagingPolicy}   
    If ($User.TeamsAppSetupPolicy) {$TeamsAppSetupPolicy = $User.TeamsAppSetupPolicy}
    If ($User.TeamsAppPermissionPolicy) {$TeamsAppPermissionsPolicy = $User.TeamsAppPermissionPolicy}
    If ($User.TeamsEnhancedEncryptionPolicy) {$TeamsEncryptionPolicy = $User.TeamsEnhancedEncryptionPolicy}
    If ($User.TeamsUpdateManagementPolicy) {$TeamsUpdatePolicy = $User.TeamsUpdateManagementPolicy}
    If ($User.TeamsChannelsPolicy) {$TeamsChannelsPolicy = $User.TeamsChannelsPolicy}
    If ($User.TeamsFeedbackPolicy) {$TeamsFeedbackPolicy = $User.TeamsFeedbackPolicy}
    If ($User.TeamsMeetingBroadcastPolicy) {$TeamsLiveEventsPolicy = $User.TeamsMeetingBroadcastPolicy}
    
    $TeamsChannelsPolicyAssignments = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsChannelsPolicy"}).PolicySource.AssignmentType
    <#
    foreach ($assignment in $TeamsChannelsPolicyAssignments) {
        write-host $assignment
    }
    #>

    $UserInfoReport += [PSCustomObject]@{
        ObjectId                        = $User.Identity;
        User                            = $User.DisplayName;
        UPN                             = $User.UserPrincipalName;
        SipAddress                      = $User.SipAddress;
	    Enabled                         = $User.Enabled;
        MessagingPolicy1                = $TeamsMessagingPolicy;
        MeetingPolicy1                   = $TeamsMeetingPolicy;
        AppSetupPolicy1                  = $TeamsAppSetupPolicy;
        AppPermissionsPolicy1            = $TeamsAppPermissionsPolicy;
        EnhancedEncryptionPolicy1        = $TeamsEncryptionPolicy;
        UpdatePolicy1                    = $TeamsUpdatePolicy;
        ChannelsPolicy1                  = $TeamsChannelsPolicy;
        FeedbackPolicy1                  = $TeamsFeedbackPolicy;
        LiveEventsPolicy1                = $TeamsLiveEventsPolicy;
	    InterpretedUserType             = $User.InterpretedUserType;
        
        HostedVoiceMail                 = $User.HostedVoiceMail;
        OnPremEnterpriseVoiceEnabled    = $User.OnPremEnterpriseVoiceEnabled;
        SipProxyAddress                 = $User.SipProxyAddress;
        TeamsUpgradeEffectiveMode       = $User.TeamsUpgradeEffectiveMode;
	    HostingProvider                 = $User.HostingProvider;
        
        LocationProfile                         = ($UserPolicies | Where-Object {$_.PolicyType -eq "LocationProfile"}).PolicyName;

        TeamsUpgradePolicy                      = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsUpgradePolicy"}).PolicyName;
        TeamsUpgradePolicySource                = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsUpgradePolicy"}).PolicySource.AssignmentType;
        
        ExternalAccessPolicy                    = ($UserPolicies | Where-Object {$_.PolicyType -eq "ExternalAccessPolicy"}).PolicyName;
        ExternalAccessPolicySource              = ($UserPolicies | Where-Object {$_.PolicyType -eq "ExternalAccessPolicy"}).PolicySource.AssignmentType;

        HostedVoicemailPolicy                   = ($UserPolicies | Where-Object {$_.PolicyType -eq "HostedVoicemailPolicy"}).PolicyName;
        HostedVoicemailPolicySource             = ($UserPolicies | Where-Object {$_.PolicyType -eq "HostedVoicemailPolicy"}).PolicySource.AssignmentType;

        MeetingPolicy2                         = ($UserPolicies | Where-Object {$_.PolicyType -eq "MeetingPolicy"}).PolicyName;
        MeetingPolicy2Source                     = ($UserPolicies | Where-Object {$_.PolicyType -eq "MeetingPolicy"}).PolicySource.AssignmentType;

        MobilityPolicy                          = ($UserPolicies | Where-Object {$_.PolicyType -eq "MobilityPolicy"}).PolicyName;
        MobilityPolicySource                    = ($UserPolicies | Where-Object {$_.PolicyType -eq "MobilityPolicy"}).PolicySource.AssignmentType;

        TeamsVideoInteropServicePolicy          = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsVideoInteropServicePolicy"}).PolicyName;
        TeamsVideoInteropServicePolicySource    = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsVideoInteropServicePolicy"}).PolicySource.AssignmentType;

        OnlineDialinConferencingPolicy          = ($UserPolicies | Where-Object {$_.PolicyType -eq "OnlineDialinConferencingPolicy"}).PolicyName;
        OnlineDialinConferencingPolicySource    = ($UserPolicies | Where-Object {$_.PolicyType -eq "OnlineDialinConferencingPolicy"}).PolicySource.AssignmentType;

        TeamsMessagingPolicy                    = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsMessagingPolicy"}).PolicyName;
        TeamsMessagingPolicySource              = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsMessagingPolicy"}).PolicySource.AssignmentType
        
        TeamsMeetingBroadcastPolicy             = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsMeetingPolicy"}).PolicyName;
        TeamsMeetingBroadcastPolicySource       = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsMeetingPolicy"}).PolicySource.AssignmentType;
        
        TeamsAppPermissionPolicy                = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsAppPermissionPolicy"}).PolicyName;
        TeamsAppPermissionPolicySource          = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsAppPermissionPolicy"}).PolicySource.AssignmentType;
        
        TeamsAppSetupPolicy                     = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsAppSetupPolicy"}).PolicyName;
        TeamsAppSetupPolicySource               = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsAppSetupPolicy"}).PolicySource.AssignmentType;
        
        TeamsChannelsPolicy                     = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsChannelsPolicy"}).PolicyName;
        TeamsChannelsPolicySource               = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsChannelsPolicy"}).PolicySource.AssignmentType;

        TenantDialPlan                          = ($UserPolicies | Where-Object {$_.PolicyType -eq "TenantDialPlan"}).PolicyName;
        TenantDialPlanSource                    = ($UserPolicies | Where-Object {$_.PolicyType -eq "TenantDialPlan"}).PolicySource.AssignmentType;
        
        TeamsCallingPolicy                      = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsCallingPolicy"}).PolicyName;
        TeamsCallingPolicySource                = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsCallingPolicy"}).PolicySource.AssignmentType;
        
        TeamsCallParkPolicy                     = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsCallParkPolicy"}).PolicyName;
        TeamsCallParkPolicySource               = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsCallParkPolicy"}).PolicySource.AssignmentType;
        
        CallerIdPolicy                          = ($UserPolicies | Where-Object {$_.PolicyType -eq "CallerIdPolicy"}).PolicyName;
        CallerIdPolicySource                    = ($UserPolicies | Where-Object {$_.PolicyType -eq "CallerIdPolicy"}).PolicySource.AssignmentType;
        
        TeamsEmergencyCallingPolicy             = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsEmergencyCallingPolicy"}).PolicyName;
        TeamsEmergencyCallingPolicySource       = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsEmergencyCallingPolicy"}).PolicySource.AssignmentType;
        
        TeamsEmergencyCallRoutingPolicy         = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsEmergencyCallRoutingPolicy"}).PolicyName;
        TeamsEmergencyCallRoutingPolicySource   = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsEmergencyCallRoutingPolicy"}).PolicySource.AssignmentType;

        VoicePolicy                             = ($UserPolicies | Where-Object {$_.PolicyType -eq "VoicePolicy"}).PolicyName;
        VoicePolicySource                       = ($UserPolicies | Where-Object {$_.PolicyType -eq "VoicePolicy"}).PolicySource.AssignmentType
    }
}

$UserInfoReport | Export-Csv "d:\exports\teams\TeamsUserInfoReport.csv" -NoTypeInformation -Encoding UTF8 -Delimiter ","

