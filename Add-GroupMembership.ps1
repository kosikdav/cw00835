#######################################################################################################################
# New-AAD-Guest
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    [string]$SourceFile,
    [string]$GroupId,
    [string]$Members
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

. $ScriptPath\include-Script-StdStartBlock.ps1

if (-not $SourceFile) {
    if (-not ($GroupId -and $Members)){
        Write-Host "ERROR: Missing parameters" -ForegroundColor Red
        exit
    }
}


#######################################################################################################################

$LogFolder			= "aad-group-mgmt"
$LogFilePrefix		= "add-group-membership"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

#######################################################################################################################

. $IncFile_StdLogStartBlock

##############################################################################################

if ($SourceFile) {
    #process source file
    $SourceList = [array](Import-CSVtoArray -Path $SourceFile)
    if ($SourceList.Count -gt 0) {
        write-host $SourceList[0]
        if (($SourceList[0].GroupId) -and ($SourceList[0].Members)) {
            foreach ($ListRecord in $SourceList) {
                Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
                $userId = $ListRecord.members -split ";"
                Add-GraphGroupMemberById -groupId $ListRecord.groupId -userId $userId -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken
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
    Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
    $userId = $Members -split ";"
    if ($userId.Count -eq 1) {
        $userId = [string]$userId[0]
    }
    Add-GraphGroupMemberById -groupId $GroupId -userId $userId -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken
}

#######################################################################################################################

. $IncFile_StdLogEndBlock
