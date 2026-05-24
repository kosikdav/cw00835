#######################################################################################################################
# Get-MGMT-API subs.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1

#######################################################################################################################

$LogFolder			    = "exports"
$LogFilePrefix		    = "spo-file-audit"

. $ScriptPath\include-Script-StdIncBlock.ps1
. $IncFile_AIP_labels
. $IncFile_Functions_Audit

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$now = Get-Date

##############################################################################
# get audit log search results
$Resource = "https://manage.office.com"
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30 -Resource $Resource -Authority "login.windows.net" -Force
$PageCount = 0
$Headers = $AuthDB[$AppReg_LOG_READER].AuthHeaders
$Uri = "https://manage.office.com/api/v1.0/$($tenantId)/activity/feed/subscriptions/list"

Do {
    #write-host $Uri -ForegroundColor Cyan
    $Response = Invoke-WebRequest -Headers $Headers -Uri $Uri -Method "GET" -UseBasicParsing
    write-host $response 
    Try{
        $QueryRecords = $Response | ConvertFrom-Json
            $PageCount++
    }
    Catch {
        Write-Host "Error converting JSON" -ForegroundColor Red
    } 
    $AvailableContentBlobs += $QueryRecords
    $Uri = $Response.Headers.NextPageUri
} Until (-not $Uri)

