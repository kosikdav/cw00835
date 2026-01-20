#######################################################################################################################
# Update-CybeReady-Config
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    [bool]$UpdateTransportRules = $True,
    [bool]$UpdateSafeLinksPolicies = $True,
    [bool]$UpdatePhishSimPolicy = $True,
    [bool]$UpdateTABList = $False,
    [int]$TABListTTL = 365
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder          = "cybeready"
$LogFilePrefix      = "config-update"
$LogFileFreq        = "YMD"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Suffix $LogFileSuffix -Ext "log"

$folder = "d:\scripts\cezdata\"

$PhishSim_Landing_Domains_File  = $folder + "phishsim-landing-domains.txt"
$PhishSim_SenderDomains_File    = $folder + "phishsim-sender-domains.txt"
$PhishSim_SenderIPs_File        = $folder + "phishsim-sender-IPs.txt"

$SafeLinks_Domains_File         = $folder + "safelinks-domains.txt"
$SafeLinks_URLs_File            = $folder + "safelinks-URLs.txt"

$TABL_Senders_Allow_File        = $folder + "TABL-senders-allow.txt"
$TABL_URLs_Allow_File           = $folder + "TABL-URLs-Allow.txt"

$TransportRules = @("ce9eaa74-1923-44cd-af9c-a0a1500a02ca", "a061a4d0-84e4-47f5-b591-58f5a61e7dc2")
$SafeLinksPolicies = @("97957906-280f-455a-b39f-042e45f24e37")

Function Get-PolicyConfigFileEntries {
    param(
        [string]$FilePath
    )
    if (Test-Path -Path $FilePath) {
        $entries = Get-Content -Path $FilePath
        $entries = $entries | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        $entries = $entries | ForEach-Object { $_.Trim() }
        return $entries | Sort-Object -Unique
    } else {
        write-log "File not found: $FilePath" -ForegroundColor Red
        return @()
    }
}

#######################################################################################################################

. $IncFile_StdLogStartBlock

$PhishSimLandingDomains = Get-PolicyConfigFileEntries -FilePath $PhishSim_Landing_Domains_File
$PhishSimSenderDomains  = Get-PolicyConfigFileEntries -FilePath $PhishSim_SenderDomains_File
$PhishSimSenderIPs      = Get-PolicyConfigFileEntries -FilePath $PhishSim_SenderIPs_File

$SafeLinksDomains       = Get-PolicyConfigFileEntries -FilePath $SafeLinks_Domains_File
$SafeLinksURLs          = Get-PolicyConfigFileEntries -FilePath $SafeLinks_URLs_File

$TABLSendersAllow       = Get-PolicyConfigFileEntries -FilePath $TABL_Senders_Allow_File
$TABLURLsAllow          = Get-PolicyConfigFileEntries -FilePath $TABL_URLs_Allow_File

Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 30

#########################################################################################
# Process Transport Rules

if ($UpdateTransportRules -and $TransportRules) {
    write-log $String_divider
    $AllEntries = @()
    $AllEntries = $PhishSimLandingDomains

    foreach ($TransportRuleId in $TransportRules) {
        $TransportRule = Get-TransportRule -Identity $TransportRuleId
        write-log "Processing Exchange transport rule: $($TransportRule.Identity) ($TransportRuleId)"
        $SubjectOrBodyContainsWords = $TransportRule.SubjectOrBodyContainsWords
        $diff = Compare-Object -ReferenceObject ($SubjectOrBodyContainsWords | Sort-Object) -DifferenceObject ($AllEntries | Sort-Object)
        if ($diff) {
            write-log "Updating transport rule: $($TransportRule.Identity) ($TransportRuleId)"
            $SubjectOrBodyContainsWords.Clear()
            foreach ($entry in $AllEntries) {
                $SubjectOrBodyContainsWords.Add($entry)
            }
            Set-TransportRule -Identity $TransportRuleId -SubjectOrBodyContainsWords $SubjectOrBodyContainsWords | Out-Null
        } else {
            write-log "No changes needed for transport rule: $($TransportRule.Identity) ($TransportRuleId)"
        }
    }
}

#########################################################################################
# Process SafeLinks Policies

if ($UpdateSafeLinksPolicies -and $SafeLinksPolicies) {
    write-log $String_divider
    $AllEntries = @()
    foreach ($Domain in $SafeLinksDomains) {
        $AllEntries += $Domain + "/*"
        $AllEntries += "*." + $Domain + "/*"
    }
    foreach ($Domain in $PhishSimLandingDomains) {
        $AllEntries += $Domain + "/*"
    }
    foreach ($URL in $SafeLinksURLs) {
        $AllEntries += $URL
    }
    $AllEntries = $AllEntries | Sort-Object -Unique

    foreach ($SafeLinksPolicyId in $SafeLinksPolicies) {
        $SafeLinksPolicy = Get-SafeLinksPolicy -Identity $SafeLinksPolicyId
        write-log "Processing SafeLinks policy: $($SafeLinksPolicy.Identity) ($SafeLinksPolicyId)"
        $diff = Compare-Object -ReferenceObject ($SafeLinksPolicy.DoNotRewriteUrls | Sort-Object) -DifferenceObject ($AllEntries | Sort-Object)
        if ($diff) {
            write-log "Updating SafeLinks policy: $($SafeLinksPolicy.Identity) ($SafeLinksPolicyId)"
            Set-SafeLinksPolicy -Identity $SafeLinksPolicyId -DoNotRewriteUrls $AllEntries
        } 
        else {
            write-log "No changes needed for SafeLinks policy: $($SafeLinksPolicy.Identity) ($SafeLinksPolicyId)"
        }   
    }
}

#########################################################################################
# Process Phishing Simulation Policy

if ($UpdatePhishSimPolicy -and ($PhishSimSenderIPs -or $SenderDomains)) {
    write-log $String_divider
    $PhishSimOverridePolicy = Get-PhishSimOverridePolicy
    $PhishSimOverrideRule = Get-ExoPhishSimOverrideRule
    if ($PhishSimOverridePolicy -and $PhishSimOverrideRule) {
        write-log "Processing phishing simulation override policy: $($PhishSimOverridePolicy.Identity)"
        write-log "Processing phishing simulation override rule: $($PhishSimOverrideRule.Identity)"

        $CurrentSenderIPRanges = $PhishSimOverrideRule.SenderIPRanges
        $CurrentSenderDomains = $PhishSimOverrideRule.Domains

        $diffIP = Compare-Object -ReferenceObject ($CurrentSenderIPRanges | Sort-Object) -DifferenceObject ($PhishSimSenderIPs | Sort-Object)
        $diffDomain = Compare-Object -ReferenceObject ($CurrentSenderDomains | Sort-Object) -DifferenceObject ($PhishSimSenderDomains | Sort-Object)

        if ($diffIP -or $diffDomain) {
            Write-log "Removing existing phishing simulation override rule: $($PhishSimOverrideRule.Identity)"
            Remove-ExoPhishSimOverrideRule -Identity $PhishSimOverrideRule.id -Confirm:$false
            Write-log "Creating new phishing simulation override rule with updated settings"
            New-ExoPhishSimOverrideRule -Policy $PhishSimOverridePolicy.id -SenderIPRanges $PhishSimSenderIPs -Domains $PhishSimSenderDomains
        } 
        else {
            write-log "No changes needed for phishing simulation override rule."
        }
    }
}

#########################################################################################
# Process Tenant Allow/Block list
if ($UpdateTABList) {
    write-log $String_divider
    write-log "Processing Tenant Allow/Block list"

    $CurrentTABLSendersAllow = Get-TenantAllowBlockListItems -ListType "Sender" -Allow
    $CurrentTABLURLsAllow = Get-TenantAllowBlockListItems -ListType "Url" -ListSubType "AdvancedDelivery" -Allow

    $AllSenderEntries = $AllURLEntries = @()
    
    $AllSenderEntries += $PhishSimSenderDomains
    $AllSenderEntries += $TABLSendersAllow

    foreach ($Domain in $PhishSimLandingDomains) {
        $AllURLEntries += "~" + $Domain
    }

    $AllURLEntries = $AllURLEntries | Sort-Object -Unique

    <#
    write-host "Current TABL Senders Allow: $($CurrentTABLSendersAllow.Count)"
    write-host ($CurrentTABLSendersAllow.Value | Sort-Object) -ForegroundColor Cyan
    write-host $AllSenderEntries -ForegroundColor Yellow
    write-host "Current TABL URLs Allow: $($CurrentTABLURLsAllow.Count)"
    write-host ($CurrentTABLURLsAllow.Value | Sort-Object) -ForegroundColor Cyan
    write-host $AllURLEntries -ForegroundColor Yellow
    #>

    Try {
        $diffSend = Compare-Object -ReferenceObject ($CurrentTABLSendersAllow.Value | Sort-Object) -DifferenceObject ($AllSenderEntries | Sort-Object)
    }
    Catch {
        if (($null -eq $CurrentTABLSendersAllow) -and $AllSenderEntries ) {
            $diffSend = $True
        }
    }
    Try {
        $diffURL = Compare-Object -ReferenceObject ($CurrentTABLURLsAllow.Value | Sort-Object) -DifferenceObject ($AllURLEntries | Sort-Object)
    }
    Catch {
        if (($null -eq $CurrentTABLURLsAllow) -and $AllURLEntries ) {
            $diffURL = $True
        }
    }

    $ExpirationDate = ((Get-Date).AddDays($TABListTTL)).ToString("yyyy-MM-ddTHH:mm:ssZ")
    write-host $ExpirationDate

    if ($diffSend) {
        write-log "Updating Tenant Allow/Block list for senders"
        foreach ($entry in $CurrentTABLSendersAllow) {
            Remove-TenantAllowBlockListItems -ListType "Sender" -Ids $entry.Identity | Out-Null
        }
        New-TenantAllowBlockListItems -Allow -ListType "Sender" -Entries $AllSenderEntries -ExpirationDate $ExpirationDate | Out-Null
    }      
    else {
        write-log "No changes needed for Tenant Allow/Block list for senders"
    }
    
    if ($diffURL) {
        write-log "Updating Tenant Allow/Block list for URLs"
        foreach ($entry in $CurrentTABLURLsAllow) {
            Remove-TenantAllowBlockListItems -ListType "Url" -ListSubType "AdvancedDelivery" -Ids $entry.Identity | Out-Null
        }
        New-TenantAllowBlockListItems -Allow -ListType "Url" -ListSubType "AdvancedDelivery" -Entries $AllURLEntries -NoExpiration | Out-Null
    } 
    else {
        write-log "No changes needed for Tenant Allow/Block list for URLs"
    }
}

#######################################################################################################################

. $IncFile_StdLogEndBlock
