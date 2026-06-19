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

#######################################################################################################################
. $IncFile_StdLogStartBlock

$Src_AppReg_EXO_MGMT = $AppReg_UJVREZ_EXO_MGMT   
Connect-EXOService -AppRegName $Src_AppReg_EXO_MGMT  -TTL 120

write-Log "SRC mailboxes total: " -NoNewline
[array]$EXOMailboxes = Get-Mailbox -ResultSize Unlimited
write-Log $EXOMailboxes.count

foreach ($Mailbox in $EXOMailboxes) {
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
}

#######################################################################################################################

. $IncFile_StdLogEndBlock