#######################################################################################################################
# XT Moves Report
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Init.ps1

#######################################################################################################################

$LogFolder					= "t2t"
$LogFilePrefix				= "xt-moves"

$OutputFolder				= "t2t\odfb-xt-moves"
$OutputFilePrefix			= "xt-moves"


$PartnerCrossTenantHostUrl = "https://cezdata-my.sharepoint.com"
$skuId_E5 = "06ebc4ee-1bb5-47dd-8120-11324bc54e06"

$XTMoves = @()

#######################################################################################################################
. $ScriptPath\include-Script-Start-Include.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$OutputFile 	= New-OutputFile -RootFolder $ROF -Folder $OutputFolder -Prefix $OutputFilePrefix -Ext "csv"

#######################################################################################################################
. $IncFile_StdLogStartBlock

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

$UriResource = "users"
$UriFilter = "userType eq 'Member'"
$UriSelect = "id,UserPrincipalName,mail,DisplayName,assignedLicenses,accountEnabled,extension_93b54ce056df45bd8f5f398753fa17c0_employeeNumber"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Filter $UriFilter -Top 999 -Select $UriSelect
[array]$Users = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "users" -ProgressDots
$counter = 0
foreach ($User in $Users) {
    write-host $User.UserPrincipalName -NoNewline
    $XTMove = $null
    $E5licensed = $false
    if ($User.assignedLicenses) {
        foreach ($License in $User.assignedLicenses) {
            if ($License.SkuId -eq $skuId_E5) {
                $E5licensed = $true
            }
        }
    }
    $ReportObject = [PSCustomObject]@{
        id = $User.id
        enabled = $User.accountEnabled
        E5lic = $E5licensed
        UserPrincipalName = $User.UserPrincipalName
        mail = $User.mail
        DisplayName = $User.DisplayName
        employeeNumber = $User.extension_93b54ce056df45bd8f5f398753fa17c0_employeeNumber
        SourceSiteUrl = $null    
        TargetSiteUrl = $null
        MoveJobId = $null
        SourceDataLocation = $null
        DestinationDataLocation = $null
        TimeStamp  = $null
        MoveState = $null
        RequestResult = $null
    }
    Try {
        $XTMove = Get-SPOCrossTenantUserContentMoveState -PartnerCrossTenantHostUrl $PartnerCrossTenantHostUrl -SourceUserPrincipalName $User.UserPrincipalName -errorAction Stop
        if ($XTMove) {
            $ReportObject.SourceSiteUrl = $XTMove.SourceSiteUrl
            $ReportObject.TargetSiteUrl = $XTMove.TargetSiteUrl
            $ReportObject.MoveJobId = $XTMove.MoveJobId
            $ReportObject.SourceDataLocation = $XTMove.SourceDataLocation
            $ReportObject.DestinationDataLocation = $XTMove.DestinationDataLocation
            $ReportObject.TimeStamp = $XTMove.TimeStamp
            $ReportObject.MoveState = $XTMove.MoveState
            $ReportObject.RequestResult = "ok"
            write-host " ok" -ForegroundColor Green
        } 
        else {
            write-host " none" -ForegroundColor Cyan
            $ReportObject.RequestResult = "None"
        }
    }
    Catch {
        write-host " error" -ForegroundColor Red
        $ReportObject.RequestResult = "Error"
    }
    $XTMoves += $ReportObject
}

Export-Report "XT Moves" -Report $XTMoves -Path $OutputFile

#######################################################################################################################

. $IncFile_StdLogEndBlock
