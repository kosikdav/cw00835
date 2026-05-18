# Script Parameters and Variables
[CmdletBinding()]
param (
    [Parameter(Mandatory=$false, Position=0)] [string]$MyUPN = "qrhorakmic@cez.cloud",
    [Parameter(Mandatory=$false, Position=1)] [string]$ExpCSVpath = "C:\Scripts\CEZ-Labels"
)

Import-Module ExchangeOnlineManagement -ErrorAction Continue
Connect-IPPSSession -UserPrincipalName $MyUPN

[string]$ExpCSVfile = Join-Path $ExpCSVpath $("LabelList-" + (Get-Date -Format "yyyyMMddHHmmss").ToString() + ".csv")


# Script Execution Code
[array]$LL = Get-Label | where {$_.Name -like 'L*'} | sort Priority
Write-Host ("Label count = " + $LL.Count.ToString()) -ForegroundColor Magenta

Remove-Variable allLabels -ErrorAction SilentlyContinue[array]$allLabelsforeach ($L in $LL) {    $L.Name    Remove-Variable myTable -ErrorAction SilentlyContinue    Remove-Variable myDisplayName -ErrorAction SilentlyContinue    Remove-Variable myTooltip -ErrorAction SilentlyContinue    $myHT = [ordered]@{}    $myDisplayName = ConvertFrom-Json $L.LocaleSettings[0]    $myTooltip = ConvertFrom-Json $L.LocaleSettings[1]    $myHT.Add("Priority",$L.Priority)    $myHT.Add("Name",$L.Name)    $myHT.Add("DisplayName",$L.DisplayName)    $myHT.Add("Tooltip",$L.Tooltip)    foreach ($DN in $myDisplayName.Settings) {$myHT.Add(($myDisplayName.LocaleKey + ".." + $DN.Key),$DN.Value)}    foreach ($DN in $myTooltip.Settings) {$myHT.Add(($myTooltip.LocaleKey + ".." + $DN.Key),$DN.Value)}    $myTable = [pscustomobject]$myHT    [array]$allLabels += $myTable}$allLabels | Export-Csv -Path $ExpCSVfile -NoTypeInformation -Delimiter ";" -Encoding UTF8# Displaying Script Results$allLabels | ft -AutoSize

Write-Host ("Script is done. CSV output file is stored here = " + $ExpCSVfile) -ForegroundColor Green
Get-Item $ExpCSVfile

# notepad $ExpCSVfile
