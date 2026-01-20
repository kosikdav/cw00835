#######################################################################################################################
# include-var-init
#######################################################################################################################

[int]$ProgressCountMain = 0
[int]$ProgressTotalMain = 1
[string]$ProgressActivityMain = [string]::Empty

[hashtable]$AuthDB = @{}
[hashtable]$AuthHeaders = @{}

[bool]$GraphError = $false
[int]$GraphErrorCode = 0
$CredFileLoadErr = $false
$interactiveRun = [Environment]::UserInteractive
$PSVersion = (Get-Host).Version.ToString()

$Today                  = $(Get-Date).AddDays(0-$daysBackOffset)
$Yesterday 				= $(Get-Date).AddDays(-1-$daysBackOffset)
$TodayBOD               = Get-Date -Hour 0 -Minute 0 -Second 0
$TodayEOD               = Get-Date -Hour 23 -Minute 59 -Second 59
$YesterdayBOD           = Get-Date ($TodayBOD.AddDays(-1))
$YesterdayEOD           = Get-Date ($TodayEOD.AddDays(-1))

$strYesterday 			= $(Get-Date).AddDays(-1-$daysBackOffset).ToString("yyyy-MM-dd")
$strToday 				= $(Get-Date).AddDays(0-$daysBackOffset).ToString("yyyy-MM-dd")
$currentDate            = $strToday


$strYesterdayUTCStart 	= $strYesterday + "T00:00:00Z"
$strYesterdayUTCEnd 	= $strToday + "T00:00:00Z"
$strNow               = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")

$OutputFileDate         = Get-Date -Format "yyyy-MM-dd"