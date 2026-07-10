#######################################################################################################################
# Fix-Holds.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile

)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1

$LogFolder			= "t2t"
$LogFilePrefix		= "t2t-mailbox-hold-remove"
$LogFileFreq		= "Y"

. $ScriptPath\include-Script-Start-Include.ps1

$LogFile 	= New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$t2tscopegroup = "4f23ce7b-6ca1-4fa9-9f3b-549c62313e04"

#######################################################################################################################
. $IncFile_StdLogStartBlock


Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "groups/$t2tscopegroup/members"
$UriSelect = "id,displayName,userPrincipalName,mail"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect -Top 999
[array]$T2TGroupMembers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
write-interactive "T2T group members total: $($T2TGroupMembers.Count)"

Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 120

<#
write-Log "SRC mailboxes total: " -NoNewline
[array]$EXOMailboxes = Get-Mailbox -ResultSize Unlimited
write-Log $EXOMailboxes.count
#>


foreach ($User in $T2TGroupMembers) {

    Try{
    $mailbox = Get-Mailbox -Identity $User.userPrincipalName -ErrorAction Stop
    }
    Catch {
        write-Log "Mailbox not found for user: $($User.userPrincipalName)"
        continue
    }
    #write-host "Processing mailbox: $($Mailbox.primarySMTPAddress)" -ForegroundColor Green
    if ($Mailbox.LitigationHoldEnabled) {
        write-Log " Removing Litigation Hold for mailbox: $($Mailbox.primarySMTPAddress)"
        Set-Mailbox -Identity $Mailbox.primarySMTPAddress -LitigationHoldEnabled $false -Confirm:$false
    }
    if ($Mailbox.DelayHoldApplied) {
        write-Log " Removing Delay Hold for mailbox: $($Mailbox.primarySMTPAddress)"
        Set-Mailbox -Identity $Mailbox.primarySMTPAddress -RemoveDelayHoldApplied
    }
    if ($Mailbox.DelayReleaseHoldApplied) {
        write-Log " Removing Delay Release Hold for mailbox: $($Mailbox.primarySMTPAddress)"
        Set-Mailbox -Identity $Mailbox.primarySMTPAddress -RemoveDelayReleaseHoldApplied
    }
    if ($Mailbox.RetentionHoldEnabled) {
        write-Log " Removing Retention Hold for mailbox: $($Mailbox.primarySMTPAddress)"
        Set-Mailbox -Identity $Mailbox.primarySMTPAddress -RetentionHoldEnabled $false
    }
    if ($Mailbox.ComplianceTagHoldApplied) {
        write-Log " Removing Compliance Tag Hold for mailbox: $($Mailbox.primarySMTPAddress)"
        Set-Mailbox -Identity $Mailbox.primarySMTPAddress -RemoveComplianceTagHoldApplied -ProvideConsent
    }
    if ($Mailbox.InPlaceHolds) {
        write-Log " In-Place Holds exist for mailbox: $($Mailbox.primarySMTPAddress)" -foregroundcolor Red
        write-Log $Mailbox.InPlaceHolds -join ","
    }
}

#######################################################################################################################

. $IncFile_StdLogEndBlock