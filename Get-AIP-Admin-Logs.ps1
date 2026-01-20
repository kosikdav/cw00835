#######################################################################################################################
# Get-AIP-Admin-Logs
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "aip-service-admin-logs"
$LogFilePrefix		= "aip-service-admin-logs"

$OutputFolder		= "aip-service-admin-logs"
$OutputFilePrefix	= "AzRMS-Admin-Log"

. $ScriptPath\include-Script-StdIncBlock.ps1

$TempFolder         = "d:\temp\aip-service-admin-logs-" + $TenantShortName
$DaysBackOffset     = 45
$IgnoredKeyWords    = @(
    "GetConnectorAuthorizations",
    "GetAuditLog",
    "GetTenantFunctionalState",
    "GetAllTemplates"
)

#######################################################################################################################

. $ScriptPath\include-Script-StdBeginBlock.ps1

$AIPFileTimeStamp  = (Get-Date).ToString("yyyyMMddHHmmss")

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$OutputFile = New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Suffix $AIPFileTimeStamp -Ext "csv" -NoDate

$CSVFolder      = [System.IO.Path]::Combine($ROF,$OutputFolder)

New-Folder -Path $TempFolder

$TmpNewRawLog   = "$($TempFolder)\TmpNewRawLog-$($AIPFileTimeStamp).log"
$TmpNewFltLog   = "$($TempFolder)\TmpNewFltLog-$($AIPFileTimeStamp).log"
$TmpCmbOldLog   = "$($TempFolder)\TmpCmbOldLog-$($AIPFileTimeStamp).log"

$FromTime   = (Get-Date).AddDays(-$DaysBackOffset)
$ToTime     = (Get-Date).AddDays(0)

function Test-AIPLogIgnoredRecord {
    param (
		[parameter(Mandatory = $true)][string]$String,
        [parameter(Mandatory = $true)][array]$IgnoredKeyWords
	)
	# main function body ##################################
    if ((-not (Test-StringContainsAnyArrayMember -String $String -Array $IgnoredKeyWords)) -and (Test-StringStartsWithUTC -String $String )) {
        Return $False
    }
    Else {
        Return $True
    }
}

#######################################################################################################################

. $IncFile_StdLogStartBlock

Write-Log "IgnoredKeyWords: $($IgnoredKeyWords)" -ForegroundColor Cyan
Write-Log "TmpNewRawLog: $($TmpNewRawLog)"
write-Log "TmpNewFltLog: $($TmpNewFltLog)"
write-Log "TmpCmbOldLog: $($TmpCmbOldLog)"
write-Log "CSVFolder: $($CSVFolder)"
. $IncFile_AppReg_LOG_READER

# connect AIPService using AAD App
Try {
    Connect-AIPService -CertificateThumbprint $Thumbprint -ApplicationId $ApplicationId -TenantId $TenantId -ServicePrincipal
    Write-Log "AIPService - connection to service was established using $($AppName)"
}
Catch {
    Write-Log "AIPService - error connecting to service using $($AppName)" -MessageType "ERROR"
}

# download AIP admin logs to $TmpNewRawLog
Try {
    Get-AIPServiceAdminLog -Path $TmpNewRawLog -FromTime $FromTime -ToTime $ToTime -Force
}
Catch {
    Write-Log "AIPService - error downloading AIP admin log using $($AppName)" -MessageType "ERROR"
}

#filter ignored records from $TmpNewRawLog and save as $TmpNewFltLog file
$NewRawContent = Get-Content -Path $TmpNewRawLog
foreach ($Line in $NewRawContent) {
    if ($Line) {
        if (Test-AIPLogIgnoredRecord -String $line -IgnoredKeyWords $IgnoredKeyWords) {
            Continue
        }
        Else {
            Add-Content -Path $TmpNewFltLog -Value $line -Encoding "UTF8"
        }
    }
}

#get all existing CSV files from folder $CSVFolder newer than x days ($DaysBackOffset)
$Path = $CSVFolder + "\*"
$Include = "*.csv"
$ExistingCSVFiles = Get-ChildItem -Path $Path -Include $Include | Where-Object {$_.LastWriteTime -gt (Get-Date).AddDays(-$DaysBackOffset)}

Write-Log "Existing CSV files newer than $($DaysBackOffset) days: $($ExistingCSVFiles.Count)"

#filter ignored records from all $ExistingCSVFiles and save as $TmpCmbOldLog file
foreach ($file in $ExistingCSVFiles) {
    #write-host "$($File.FullName) $($File.LastWriteTime)"
    $Content = Get-Content -Path $file.FullName
    foreach ($Line in $Content) {
        if ($Line) {
            if (Test-AIPLogIgnoredRecord -String $line -IgnoredKeyWords $IgnoredKeyWords) {
                Continue
            }
            Else {
                Add-Content -Path $TmpCmbOldLog -Value $line -Encoding "UTF8"
            }
        }
    }
}

#compare existing saved logs ($TmpCmbOldLog) with newly downloaded log ($TmpNewFltLog) and only keep records newly added records in $NewEntries array
$NewEntries = Compare-Object (Get-Content $TmpCmbOldLog) (Get-Content $TmpNewFltLog) | Where-Object -FilterScript { $_.SideIndicator -eq "=>" } | Select-Object -ExpandProperty InputObject

#clean up temp files
Remove-File -Path $TmpNewRawLog
Remove-File -Path $TmpNewFltLog
Remove-File -Path $TmpCmbOldLog

#export $NewEntries array as new CSV output file
Write-Log "Exporting $($NewEntries.Count) new entries to $($OutputFile)"
Add-Content -Path $OutputFile -Value $NewEntries

#######################################################################################################################

. $IncFile_StdLogEndBlock
