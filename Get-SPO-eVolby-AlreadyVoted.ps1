$EnableOnScreenLogging = $true
$ManualAuth = $false
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
$Stopwatch =  [system.diagnostics.stopwatch]::StartNew()

. $ScriptPath\include-Functions-Init.ps1
. $ScriptPath\include-Root-Vars.ps1
. $ScriptPath\include-Functions-Common.ps1

##################################################################################################
$LogFolder			= "exports"
$LogFilePrefix		= "eVolby-voted"
$LogFileFreq		= "Y"

$OutputFolder		= "eVolby-voted"
$OutputFilePrefix	= "eVolby"
$OutputFileSuffix	= "voted"
$OutputFileFreq   	= "YMDHM"

##################################################################################################

$LogPath            = Join-Path $root_log_folder $LogFolder
$LogFileName        = $LogFilePrefix + (GetTimestamp($LogFileFreq)) + ".log"

$OutputPath         = Join-Path $root_export_folder $OutputFolder
$OutputFileName     = $OutputFilePrefix + "-" + (GetTimestamp($OutputFileFreq)) + "-" + $OutputFileSuffix + ".csv"
$OutputFile         = Join-Path $OutputPath $OutputFileName

##################################################################################################
. $ScriptPath\include-Script-StartLog-Generic.ps1
. $ScriptPath\include-appreg-CEZ_SPO_PnP_REPORTS.ps1

Connect-PnPOnline -Url $PnPURL -ClientId $ClientId -Tenant $PNPTenant -Thumbprint $Thumbprint
[array]$Result = @()
$SiteURL = "https://cezdata.sharepoint.com/sites/e-volby"
$ListName = "AlreadyVoted"  
# Connect to PnP Online
Connect-PnPOnline -Url $SiteURL -ClientId $ClientId -Tenant $PNPTenant -Thumbprint $Thumbprint       
# Get List Items
$ListItems = (Get-PnPListItem -List $ListName -Fields "ID","ID_kampane","ID_obvodu","eVolby_UPN_volic","Title","Autor")  
ForEach($ListItem in $ListItems) {
    $Result += [pscustomobject]@{
        ID                  = $ListItem["ID"];
        ID_kampane          = $ListItem["ID_kampane"];
        #ID_obvodu           = $ListItem["ID_obvodu"];
        #evolby_UPN_volic    = $ListItem["eVolby_UPN_volic"];
        Title               = $ListItem["Title"]
        #Autor               = $ListItem["Autor"]
    }
} 
# Export the result to CSV
$Result | Export-CSV -path $OutputFile -NoTypeInformation -Encoding UTF8  
