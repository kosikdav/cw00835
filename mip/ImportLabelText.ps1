# Script Parameters and Variables
[CmdletBinding()]
param (
    [Parameter(Mandatory=$false, Position=0)] [string]$MyUPN = "qrhorakmic@cez.cloud",
    [Parameter(Mandatory=$false, Position=1)] [string]$ImpCSVfile = "C:\Scripts\CEZ-Labels\LabelList.csv"
)

Import-Module ExchangeOnlineManagement -ErrorAction Continue
Connect-IPPSSession -UserPrincipalName $MyUPN


# Script Execution Code

### Validation of sources
[array]$LangTable = Import-Csv -Path $ImpCSVfile -Delimiter ";" -ErrorAction Continue
[array]$LL = Get-Label | where {$_.Name -like 'L*'} | sort Priority
Write-Host ("CSV Row count = " + $LangTable.Count.ToString()) -ForegroundColor Magenta -ErrorAction Continue
Write-Host ("MIP Label count = " + $LL.Count.ToString()) -ForegroundColor Magenta -ErrorAction Continue

### Update Label Descriptions in All Defined Languages
If (($LangTable.Count -gt 0) -and ($LL.Count -gt 0)) {

    $htLL = @{}    $LL | %{$htLL.Add($_.Priority,$_.Name)}    [int]$c = 0    foreach ($T in $LangTable) {        $c++        Write-Host        Write-Host ($c.ToString() + "/" + $LangTable.Count) -ForegroundColor Cyan        Remove-Variable MyErr -ErrorAction SilentlyContinue                If ($htLL.ContainsValue($T.Name)) {            $dnLang =
            @([pscustomobject]@{Key = "en-us"; Value = $T.'displayName..en-us'},
            [pscustomobject]@{Key = "cs-cz"; Value = $T.'displayName..cs-cz'})
            $dnSettings = 
            @([pscustomobject]@{LocaleKey = "displayName"; Settings = $dnLang})
            $dnJSON = ConvertTo-Json $dnSettings -Depth 99

            $ttLang =
            @([pscustomobject]@{Key = "en-us"; Value = $T.'tooltip..en-us'},
            [pscustomobject]@{Key = "cs-cz"; Value = $T.'tooltip..cs-cz'})
            $ttSettings = 
            @([pscustomobject]@{LocaleKey = "tooltip"; Settings = $ttLang})
            $ttJSON = ConvertTo-Json $ttSettings -Depth 99

            # Fixing MS-screwed-up API
            $LangSettingsARRAY = @($dnJSON.Substring(1,$dnJSON.Length-2) , $ttJSON.Substring(1,$ttJSON.Length-2))

            Set-Label -Identity $T.Name -DisplayName $T.DisplayName -Tooltip $T.Tooltip -LocaleSettings $LangSettingsARRAY -ErrorVariable MyErr

            If ($MyErr) {Write-Host ("Failure = " + $T.Name + " update failed.") -ForegroundColor Red} Else {Write-Host ("Success = " + $T.Name + " was updated.") -ForegroundColor Green}

        }
        Else {Write-Host ("CSV Label Name ... " + $T.Name + " ... has NO match in the list of current online MIP labels!  Skipping this item.") -ForegroundColor Yellow}
    }

} 
Else {Write-Host "Reading of source data failed. Fix the error cause and try again." -ForegroundColor Red}
# Script ResultsWrite-Host 
Write-Host ("Script is done.") -ForegroundColor Green
