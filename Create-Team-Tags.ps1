cls
$EnableOnScreenLogging = $true
$ScriptName = $MyInvocation.MyCommand.Name
$Stopwatch =  [system.diagnostics.stopwatch]::StartNew()

# Import code for function "GetLogFileName"
. d:\scripts\include-function-GetLogFileName.ps1
# Import code for root vars
. d:\scripts\include-root-vars.ps1

$LogPath 		= $root_log_folder + "team-tags-poradna\"
$LogFilePrefix	= "team-tags-poradna-"
$LogFileName    = GetLogFileName("Y")

$OutputPath 		= $root_export_folder + "\"
$OutputFilePrefix	= "poradna-members-"
$OutputFileNameTms 	= $OutputPath + $OutputFilePrefix + $strToday + ".csv"

# Import code for AAD app reg CEZ_EXO_MBX_MGMT
. d:\scripts\include-appreg-CEZ_TEAMS_MGMT.ps1
# Import code for function "LogWrite"
. d:\scripts\include-function-LogWrite.ps1
# Import code for Get-MSALToken -> $accessToken
. d:\scripts\include-GetMSALToken.ps1
# Connection to Exchange Online

Import-Module Microsoft.Graph.Teams
Select-MgProfile -Name "beta"

$Certificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"
Connect-MgGraph -ClientID $ClientId -TenantId $TenantId -Certificate $Certificate

# Poradna ICT Services
$TeamId = "d9e1b8c5-3dab-416e-9e3d-c332ce177c7b"


#$UserId = "c5c434b5-98cf-4fcb-8e52-7435c2196be7"
<#
New-MgTeamTag -TeamId $TeamId -Description "CEP01" -DisplayName "CEP01" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEP02" -DisplayName "CEP02" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ01" -DisplayName "CEZ01" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ02" -DisplayName "CEZ02" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ03" -DisplayName "CEZ03" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ04" -DisplayName "CEZ04" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ05" -DisplayName "CEZ05" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ06" -DisplayName "CEZ06" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ07" -DisplayName "CEZ07" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ08" -DisplayName "CEZ08" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ09" -DisplayName "CEZ09" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ10" -DisplayName "CEZ10" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ11" -DisplayName "CEZ11" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ12" -DisplayName "CEZ12" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ13" -DisplayName "CEZ13" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ14" -DisplayName "CEZ14" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ15" -DisplayName "CEZ15" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ16" -DisplayName "CEZ16" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ17" -DisplayName "CEZ17" -Members @{UserId =  $UserId}
#>
<#
New-MgTeamTag -TeamId $TeamId -Description "CEZd01" -DisplayName "CEZd01" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd02" -DisplayName "CEZd02" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd03" -DisplayName "CEZd03" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd04" -DisplayName "CEZd04" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd05" -DisplayName "CEZd05" -Members @{UserId =  $UserId}


$userid = "0316c3f1-ca31-4ad8-9c00-bb7ab0ada036"
New-MgTeamTag -TeamId $TeamId -Description "CEZd06" -DisplayName "CEZd06" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd07" -DisplayName "CEZd07" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd08" -DisplayName "CEZd08" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd09" -DisplayName "CEZd09" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd10" -DisplayName "CEZd10" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd11" -DisplayName "CEZd11" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd12" -DisplayName "CEZd12" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd13" -DisplayName "CEZd13" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd14" -DisplayName "CEZd14" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd15" -DisplayName "CEZd15" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd16" -DisplayName "CEZd16" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd17" -DisplayName "CEZd17" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd18" -DisplayName "CEZd18" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd19" -DisplayName "CEZd19" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd20" -DisplayName "CEZd20" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd21" -DisplayName "CEZd21" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd22" -DisplayName "CEZd22" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd23" -DisplayName "CEZd23" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZd24" -DisplayName "CEZd24" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ-DJE01" -DisplayName "CEZ-DJE01" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ-DJE02" -DisplayName "CEZ-DJE02" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ-DJE03" -DisplayName "CEZ-DJE03" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ-DJE04" -DisplayName "CEZ-DJE04" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ-DJE05" -DisplayName "CEZ-DJE05" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ-DJE06" -DisplayName "CEZ-DJE06" -Members @{UserId =  $UserId}


$userid = "d4e1136f-5718-40fd-b646-59f16e5f2c56"
New-MgTeamTag -TeamId $TeamId -Description "CEZ-DJE07" -DisplayName "CEZ-DJE07" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ-DJE08" -DisplayName "CEZ-DJE08" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ-DJE09" -DisplayName "CEZ-DJE09" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ-DJE10" -DisplayName "CEZ-DJE10" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ-DJE11" -DisplayName "CEZ-DJE11" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ-DJE12" -DisplayName "CEZ-DJE12" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ-DJE13" -DisplayName "CEZ-DJE13" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CEZ-DJE14" -DisplayName "CEZ-DJE14" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "COZ01" -DisplayName "COZ01" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CPR01" -DisplayName "CPR01" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CPR02" -DisplayName "CPR02" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CPR03" -DisplayName "CPR03" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CPR04" -DisplayName "CPR04" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CPR05" -DisplayName "CPR05" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CPR06" -DisplayName "CPR06" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CPR07" -DisplayName "CPR07" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "CT01" -DisplayName "CT01" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "Elevion01" -DisplayName "Elevion01" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "ENERGOTRANS01" -DisplayName "ENERGOTRANS01" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "ENERGOTRANS02" -DisplayName "ENERGOTRANS02" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "ESCO01" -DisplayName "ESCO01" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "ESCO02" -DisplayName "ESCO02" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "Guests01" -DisplayName "Guests01" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "ICTS01" -DisplayName "ICTS01" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "ICTS02" -DisplayName "ICTS02" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "ICTS03" -DisplayName "ICTS03" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "SKC01" -DisplayName "SKC01" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "SKC02" -DisplayName "SKC02" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "SKC03" -DisplayName "SKC03" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "SKC04" -DisplayName "SKC04" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "SKC05" -DisplayName "SKC05" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "SKC06" -DisplayName "SKC06" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "SKC07" -DisplayName "SKC07" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "SKC08" -DisplayName "SKC08" -Members @{UserId =  $UserId}
New-MgTeamTag -TeamId $TeamId -Description "TPS01" -DisplayName "TPS01" -Members @{UserId =  $UserId}
#>

$userid = "b260c285-6d94-40a3-9842-4c8cffe3c4d9"
New-MgTeamTag -TeamId $TeamId -Description "ICTS01" -DisplayName "ICTS01" -Members @{UserId =  $UserId}


