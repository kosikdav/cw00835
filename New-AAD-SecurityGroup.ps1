#######################################################################################################################
# New-AAD-Guest
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    [string]$SourceFile,
    [string]$DisplayName,
    [string]$Description,
    [bool]$MailEnabled = $false,
    [string]$MailNickname
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

. $ScriptPath\include-Script-StdStartBlock.ps1

if (-not $SourceFile) {
    if (-not $DisplayName) {
        Write-Host "ERROR: Missing parameter -DisplayName" -ForegroundColor Red
        exit
    }
}

#######################################################################################################################

$LogFolder			= "new-aad-group"
$LogFilePrefix		= "new-aad-group"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

#######################################################################################################################

. $IncFile_StdLogStartBlock

##############################################################################################

Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30

if ($SourceFile) {
    #process source file
    $SourceList = Import-CSVtoArray -Path $SourceFile
    if ($SourceList.Count -gt 0) {
        write-host $SourceList[0]
        if ($SourceList[0].displayName) {
            foreach ($ListRecord in $SourceList) {
                Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
                $Description = "Datova platforma ESCO - role " + $ListRecord.displayName
                New-GraphSecurityGroup -DisplayName $ListRecord.displayName -Description $Description -MailEnabled $MailEnabled -MailNickname $MailNickname -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken
            }
        }
        else {
            Write-Log "ERROR: source file does not contain required columns" -ForegroundColor "Red"
            Exit
        }
    }
    else {
        Write-Log "ERROR: source file empty" -ForegroundColor "Red"
        Exit
    }
}
else {
    #process single user
    New-GraphSecurityGroup -DisplayName $DisplayName -Description $Description -MailEnabled $MailEnabled -MailNickname $MailNickname -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken
}

#######################################################################################################################

. $IncFile_StdLogEndBlock
