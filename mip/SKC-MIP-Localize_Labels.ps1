# Script Parameters and Variables
[CmdletBinding()]
param (
    [Parameter(Mandatory=$false, Position=0)] [string]$AdminUPN = "name@domain.onmicrosoft.com",
    [Parameter(Mandatory=$false, Position=1)] [string]$InputFile = "filepath\filename.csv"
)

# Connect MIP Service
Import-Module ExchangeOnlineManagement -Verbose -ErrorAction Continue
Connect-IPPSSession -UserPrincipalName $AdminUPN

# Get Labels from MIP Service
[array]$LL = Get-Label
Write-Host ("MIP Label count = " + $LL.Count.ToString()) -ForegroundColor Magenta -ErrorAction Continue

# Load CSV file with label language definition 
[array]$LangTable = Import-Csv -Path $InputFile -Delimiter ";" -ErrorAction Continue
Write-Host ("CSV Row count = " + $LangTable.Count.ToString()) -ForegroundColor Magenta -ErrorAction Continue

# Aktualizace MIP labelù podle vstupù z CSV
If (($LangTable.Count -gt 0) -and ($LL.Count -gt 0)) {

    $htLL = @{}
    $LL | %{$htLL.Add($_.Priority,$_.Name)}

    [int]$c = 0
    foreach ($T in $LangTable) {

        $c++
        Write-Host
        Write-Host ($c.ToString() + "/" + $LangTable.Count) -ForegroundColor Cyan
        Remove-Variable MyErr -ErrorAction SilentlyContinue
        
        If ($htLL.ContainsValue($T.Name)) {

            $dnLang =
            @([pscustomobject]@{Key = "en-us"; Value = $T.dnEN.ToString()},
            [pscustomobject]@{Key = "cs-cz"; Value = $T.dnCZ.ToString()})
            $dnSettings = 
            @([pscustomobject]@{LocaleKey = "displayName"; Settings = $dnLang})
            $dnJSON = ConvertTo-Json $dnSettings -Depth 99

            $ttLang =
            @([pscustomobject]@{Key = "en-us"; Value = $T.ttEN.ToString()},
            [pscustomobject]@{Key = "cs-cz"; Value = $T.ttCZ.ToString()})
            $ttSettings = 
            @([pscustomobject]@{LocaleKey = "tooltip"; Settings = $ttLang})
            $ttJSON = ConvertTo-Json $ttSettings -Depth 99

            # Fixing MS-screwed-up PS vs API JSON format
            $LangSettingsARRAY = @($dnJSON.Substring(1,$dnJSON.Length-2) , $ttJSON.Substring(1,$ttJSON.Length-2))

            Set-Label -Identity $T.Name -DisplayName $T.dnEN.ToString() -Tooltip $T.ttEN.ToString() -LocaleSettings $LangSettingsARRAY -ErrorVariable MyErr

            If ($MyErr) {Write-Host ("Failure = " + $T.Name + " update failed.") -ForegroundColor Red} Else {Write-Host ("Success = " + $T.Name + " was updated.") -ForegroundColor Green}

        }
        Else {Write-Host ("CSV Label Name ... " + $T.Name + " ... has NO match in the list of current online MIP labels!  Skipping this item.") -ForegroundColor Yellow}

    }

} 
Else {Write-Host "Reading of source data failed. Fix the error cause and try again." -ForegroundColor Red}

# Script Results
Write-Host 
Write-Host ("Script is done.") -ForegroundColor Green
