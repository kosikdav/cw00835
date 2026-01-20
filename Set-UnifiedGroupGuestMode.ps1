#######################################################################################################################
# Set-UnifiedGroupGuestMode
#######################################################################################################################
 param (
    [Parameter(Mandatory=$true)][string]$Group,
    [Parameter(Mandatory=$true)][bool]$AllowToAddGuests,    
    [switch]$Silent
)

$GroupId = $Group
$AllowToAddGuestsTOBE = $AllowToAddGuests

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
$AADGroup = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB["CEZ_AAD_USR_REPORT"].AccessToken -ContentType $ContentTypeJSON -Silent

if ($AADGroup) {
    write-host "$($AADGroup.displayName) ($($AADGroup.Id)) " -NoNewline
    if ($AADGroup.groupTypes.Contains("Unified")) {
        $UriResource = "groupSettingTemplates"
        $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
        $groupSettingTemplates = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB["CEZ_AAD_USR_REPORT"].AccessToken -ContentType $ContentTypeJSON -Silent
        ForEach ($Template in $groupSettingTemplates) {
            if ($Template.displayName -eq "Group.Unified.Guest") {
                $TemplateId = $Template.Id
            }
        }
        if ($TemplateId) {
            $UriResource = "groups/$($GroupId)/settings"
            $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
            $GroupSettings = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB["CEZ_AAD_USR_REPORT"].AccessToken -ContentType $ContentTypeJSON -Silent
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
            write-host "- setting " -NoNewline
            write-host "AllowToAddGuests " -NoNewline -ForegroundColor Cyan
            write-host "to " -NoNewline
            write-host "$($AllowToAddGuestsTOBE)" -ForegroundColor Cyan -NoNewline
            write-host ": " -NoNewline
            if ($AllowToAddGuestsTOBE -ne $AllowToAddGuestsASIS) {
                Request-MSALToken -AppRegName "CEZ_AAD_USR_MGMT" -TTL 30 -Silent
                
                if ($null -ne $AllowToAddGuestsASIS) {
                    # value exists - updating (PATCH)
                    $UriResource = "groups/$($GroupId)/settings/$($AllowToAddGuestsSettingsId)"
                    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
                    $params = @{
                        values = @(
                            @{
                                name = "AllowToAddGuests"
                                value = "$($AllowToAddGuestsTOBE)"
                            }
                        )
                    }
                    $Body = $params | ConvertTo-Json -Depth 3
                    Try {
                        $ResponsePATCH = Invoke-WebRequest -Headers $AuthDB["CEZ_AAD_USR_MGMT"].AuthHeaders -Uri $Uri -Body $Body -Method "PATCH" -ContentType $ContentTypeJSON
                        write-host "success" -ForegroundColor Green
                    }
                    Catch {
                        $errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                        Write-Host "error" -ForegroundColor Red
                        Write-Host "$($errObj.error.code)" -ForegroundColor Red
                    }  
                }
                else {
                    # value not found - creating (POST)
                    $UriResource = "groups/$($GroupId)/settings"
                    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
                    $params = @{
                        templateId = $TemplateId
                        values = @(
                            @{
                                name = "AllowToAddGuests"
                                value = "$($AllowToAddGuestsTOBE)"
                            }
                        )
                    }
                    $Body = $params | ConvertTo-Json -Depth 3
                    Try {
                        $ResponsePOST = Invoke-WebRequest -Headers $AuthDB["CEZ_AAD_USR_MGMT"].AuthHeaders -Uri $Uri -Body $Body -Method "POST" -ContentType $ContentTypeJSON
                        write-host "success" -ForegroundColor Green
                    }
                    Catch {
                        $errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                        Write-Host "error" -ForegroundColor Red
                        Write-Host "$($errObj.error.code)" -ForegroundColor Red
                    }  
                }
            }
            Else {
                write-host "success" -ForegroundColor Green
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
    Write-Host "group $($groupId) not found in AAD" -ForegroundColor Red
}
