#######################################################################################################################
# Get-UnifiedGroupGuestMode
#######################################################################################################################
 param (
    [Parameter(Mandatory=$true)][string]$Group,
    [switch]$Silent
)

$GroupId = $Group

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

$TemplateId = $null
$AllowToAddGuestsASIS = $null
$AllowToAddGuestsSettingsId = $null

. $ScriptPath\include-Var-Define.ps1
. $ScriptPath\include-Var-Init.ps1
. $ScriptPath\include-Functions-Common.ps1

#######################################################################################################################

Request-MSALToken -AppRegName "CEZ_AAD_USR_REPORT" -TTL 30 -Silent

$UriResource = "groups/$($groupId)"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
$AADGroup = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB["CEZ_AAD_USR_REPORT"].AccessToken -ContentType $ContentTypeJSON

if ($AADGroup) {
    write-host "$($AADGroup.displayName) ($($AADGroup.Id)): " -ForegroundColor Yellow -NoNewline
    if ($AADGroup.groupTypes.Contains("Unified")) {
        $UriResource = "groupSettingTemplates"
        $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
        $groupSettingTemplates = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB["CEZ_AAD_USR_REPORT"].AccessToken -ContentType $ContentTypeJSON
        ForEach ($Template in $groupSettingTemplates) {
            if ($Template.displayName -eq "Group.Unified.Guest") {
                $TemplateId = $Template.Id
            }
        }
        if ($TemplateId) {
            $UriResource = "groups/$($GroupId)/settings"
            $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
            $GroupSettings = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB["CEZ_AAD_USR_REPORT"].AccessToken -ContentType $ContentTypeJSON
            foreach ($Setting in $GroupSettings) {
                if ($Setting.values.name -eq "AllowToAddGuests") {
                    $AllowToAddGuestsSettingsId = $Setting.id
                    if ($Setting.values.value -eq $true) {
                        $AllowToAddGuestsASIS = $true
                    }
                    else {
                        $AllowToAddGuestsASIS = $false
                    }
                }
            }
            If ($null -ne $AllowToAddGuestsASIS) {
                write-host "AllowToAddGuests value: " -NoNewline
                write-host "$($AllowToAddGuestsASIS)" -ForegroundColor Cyan
            }
            Else {
                write-host "AllowToAddGuests value not set" -ForegroundColor DarkGray
            }
        }
        Else {
            Write-Host "template Group.Unified.Guest not found" -ForegroundColor Red
        }
    }
    Else {
        Write-Host "group type not unified" -ForegroundColor Red
    }
}
Else {
    Write-Host "group not found in AAD"
}
