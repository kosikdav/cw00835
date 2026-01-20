#######################################################################################################################
# Remove-AAD-Guest
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
    [string]$SourceFile,
    [string]$Mail,
    [string]$Identity
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

. $ScriptPath\include-Script-StdStartBlock.ps1

if (-not $SourceFile) {
    if (-not ($Mail -or $Identity)) {
        Write-Host "ERROR: Missing parameter -Mail or -Identity" -ForegroundColor Red
        exit
    }
}

#######################################################################################################################

$LogFolder			= "remove-aad-guest"
$LogFilePrefix		= "remove-aad-guest"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

#######################################################################################################################

. $IncFile_StdLogStartBlock

##############################################################################################
# read Guests from Graph 

if ($SourceFile) {
    #processing source file
    $SourceList = Import-CSVtoArray -Path $SourceFile
    if ($SourceList.Count -gt 0) {
        foreach ($Guest in $SourceList) {
            Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
            Remove-B2BUser -Identity $Guest.Identity -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken
        }
    }
    else {
        Write-Log "ERROR: source file empty" -ForegroundColor "Red"
        Exit
    }
}
else {
    #single user removal
    if ($Identity) {
        Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
        Remove-B2BUser -Identity $Identity -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken      
    }
}

#######################################################################################################################

. $IncFile_StdLogEndBlock
