#######################################################################################################################
# New-EXOAppRoleAssignment
#######################################################################################################################
param (
    [string]$AppName,
    [string]$AppId,
    [string]$DGName,
    [bool]$Scoped = $true,
    [switch]$Silent
)
if (-not ($AppName -or $AppId)) {
    Write-host "Either AppName or AppId parameter must be used" -ForegroundColor DarkYellow
    Exit
}

$EXOConnected = $false
$AADConnected = $false

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

. $ScriptPath\include-Var-Define.ps1
. $ScriptPath\include-Var-Init.ps1
. $ScriptPath\include-Functions-Common.ps1

[hashtable]$AADSPDB_ByAppId = @{}
[hashtable]$AADSPDB_ByObjectId = @{}
[hashtable]$AADAppDB_ByAppId = @{}
[System.Collections.ArrayList]$ExistingRAs = @()
$AADSPDisplayName = [string]::Empty
$AADSPAppId = [string]::Empty
$AADSPObjectId = [string]::Empty
$EXOSP = $null
$EXOMgmtScope = $null
$DGDN = $null
$DGGuid = $null
$EXOSPDisplayName = $null
$EXOMgmtScopeName = $null
$EXORole = $null

$TableWidthValues = 118
$KeyWidthValues = 28
$KeyWidthLongValues = 35

$TableWidthRoles = 60
$KeyWidthRoles = 3

#############################################################################################################################
function Get-YesNoKeyboardInput {
    param (
        [Parameter(Mandatory=$true)][string]$Prompt
    )
    Write-Host "$($Prompt) [Y/N]" -ForegroundColor Yellow
    :prompt 
    while ($true) {
        switch ([console]::ReadKey($true).Key) {        
            { $_ -eq [System.ConsoleKey]::Y } { Return $true }        
            { $_ -eq [System.ConsoleKey]::N } { Return $false }        
            default { Write-Host "Only 'Y' or 'N' allowed!" }    
        }
    }
}
function Get-PaddedString {
    param (
        [string]$String,
        [Parameter(Mandatory=$true)][int]$Length
    )
    if ($String) {
        return $String.PadRight($Length)
    }
    else {
        return " "*$Length
    }
}

function Write-TableHeader {
    param (
        [Parameter(Mandatory=$true)][string]$String,
        [string]$TextColor = "Cyan",
        [string]$LineColor = "White",
        [Parameter(Mandatory=$true)][int]$Width
    )
    Write-Host $("="*$Width) -ForegroundColor $LineColor
    Write-Host "| " -NoNewline -ForegroundColor $LineColor
    Write-Host "$($String.PadRight($Width-3))" -NoNewline -ForegroundColor $TextColor
    Write-Host "|" -ForegroundColor $LineColor
    Write-Host $("="*$Width) -ForegroundColor $LineColor
}
function Write-TableLine {
    param (
        [Parameter(Mandatory=$true)][string]$Key,
        [string]$Value,
        [Parameter(Mandatory=$true)][int]$Width,
        [int]$KeyWidth = 25,
        [string]$KeyColor = "White",
        [string]$ValueColor = "Green"

    )
    if ($Key.Length -gt ($KeyWidth-1)) {
        $Key = $Key.Substring(0, $KeyWidth-4) + "..."
    }
    if ($Value.Length -gt ($Width - $KeyWidth - 3)) {
        $Value = $Value.Substring(0, $Width - $KeyWidth - 7) + "..."
    }
    Write-Host "| " -NoNewline
    Write-Host "$(Get-PaddedString -String $Key -Length ($KeyWidth-1))" -NoNewline -ForegroundColor $KeyColor
    Write-Host "| " -NoNewline
    write-host "$(Get-PaddedString -String $Value -Length ($Width-$KeyWidth-4))" -NoNewline -ForegroundColor $ValueColor
    Write-Host "|"
}
function Write-TableFooter {
    param (
        [Parameter(Mandatory=$true)][int]$Width = 100,
        [string]$LineColor = "White"
    )
    Write-Host $("="*$Width) -ForegroundColor $LineColor
}
function Write-CurrentValues {
    Clear-Host
    Write-TableHeader -String "Current values" -TextColor "Cyan" -LineColor "White" -Width $TableWidthValues
    Write-TableLine -Key "AAD SP display name" -Value $script:AADSPDisplayName -Width $TableWidthValues -KeyWidth $KeyWidthValues
    Write-TableLine -Key "AAD SP app id" -Value $script:AADSPAppId -Width $TableWidthValues -KeyWidth $KeyWidthValues
    Write-TableLine -Key "AAD SP object id" -Value $script:AADSPObjectId -Width $TableWidthValues -KeyWidth $KeyWidthValues
    Write-TableLine -Key "EXO group name" -Value $script:DGName -Width $TableWidthValues -KeyWidth $KeyWidthValues
    Write-TableLine -Key "EXO group DN" -Value $script:DGDN -Width $TableWidthValues -KeyWidth $KeyWidthValues
    Write-TableLine -Key "EXO group AAD Id" -Value $script:DGGuid -Width $TableWidthValues -KeyWidth $KeyWidthValues
    Write-TableLine -Key "EXO service principal name" -Value $script:EXOSPDisplayName -Width $TableWidthValues -KeyWidth $KeyWidthValues
    Write-TableLine -Key "EXO management scope name" -Value $script:EXOMgmtScopeName -Width $TableWidthValues -KeyWidth $KeyWidthValues
    Write-TableLine -Key "EXO role" -Value $script:EXORole -Width $TableWidthValues -KeyWidth $KeyWidthValues -ValueColor "Cyan"
    Write-TableFooter -Width $TableWidthValues
    
}

function Write-CurrentRoleAssignments {
    param (
        [Parameter(Mandatory=$true)][string]$ObjectId
    )
    [Array]$ExistingRAs = @()
    $EXOMgmtRoleAssignments = Get-ManagementRoleAssignment -RoleAssigneeType "ServicePrincipal"
    foreach ($EXOMgmtRoleAssignment in $EXOMgmtRoleAssignments) {
        if ($EXOMgmtRoleAssignment.RoleAssignee -eq $ObjectId) {
            $RAObject = [PSCustomObject]@{
                Id = $EXOMgmtRoleAssignment.Id
                Role = $EXOMgmtRoleAssignment.Role                
                CustomResourceScope = $EXOMgmtRoleAssignment.CustomResourceScope
            }
            $ExistingRAs += $RAObject
        }
    }
    
    Clear-Host
    
    Write-TableHeader -String "Existing role assignments for $($AADSPDisplayName):" -TextColor "Cyan" -LineColor "White" -Width $TableWidthValues
    foreach ($RA in $ExistingRAs) {
        Write-TableLine -Key $RA.Role -Value $RA.CustomResourceScope -Width $TableWidthValues -KeyWidth $KeyWidthLongValues
    }
    Write-TableFooter -Width $TableWidthValues
}

function Get-AnyKeyboardInput {
    param (
        [Parameter(Mandatory=$true)][string]$Prompt
    )
    Write-Host "$($Prompt)" -ForegroundColor Yellow
    Return [console]::ReadKey($true).Key
}
#############################################################################################################################
#script start

Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 120
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

$UriResource = "servicePrincipals"
$UriSelect = "id,displayName,appId"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
$AADServicePrincipals = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
foreach ($ServicePrincipal in $AADServicePrincipals) {
    $Record = [PSCustomObject]@{
        id = $ServicePrincipal.id
        DisplayName = $ServicePrincipal.DisplayName
        AppId = $ServicePrincipal.AppId
    }
    $AADSPDB_ByAppId.Add($ServicePrincipal.AppId, $Record)
    $AADSPDB_ByObjectId.Add($ServicePrincipal.Id, $Record)
}

$UriResource = "applications"
$UriSelect = "id,displayName,appId"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
$AADApplications = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON
foreach ($Application in $AADApplications) {
    $Record = [PSCustomObject]@{
        id = $Application.id
        DisplayName = $Application.DisplayName
        AppId = $Application.AppId
    }
    $AADAppDB_ByAppId.Add($Application.AppId, $Record)
}

#############################################################################################################################
# AAD SP

Write-CurrentValues

write-host "AAD: " -ForegroundColor cyan -NoNewline
#appId entered
if ($appId) {
    $appId = $appId.Trim()
    if ($AADSPDB_ByAppId.ContainsKey($appId)) {
        $AADSP_byAppId = $AADSPDB_ByAppId[$appId]
        write-host "found existing service principal " -NoNewline
        write-host $AADSP_byAppId.DisplayName -NoNewline -ForegroundColor Green
        write-host " with appId " -NoNewline
        write-host $AADSP_byAppId.appId -ForegroundColor Green
        if ((Get-YesNoKeyboardInput "Continue and use this AAD service principal?")) {
            $AADSPAppId = $AADSP_byAppId.appId
            $AADSPObjectId = $AADSP_byAppId.id
            $AADSPDisplayName = $AADSP_byAppId.DisplayName
            $Name = ($AADSPDisplayName.ToUpper()).Trim()
            if ($Name.StartsWith("CEZ_")) {
                $Name = $Name.Substring(4)
            }
        }
        Else {
            exit
        }
    }
    else {
        write-host "service principal with appId $($appId) not found in AAD" -ForegroundColor DarkYellow
        exit
    }
}
#appId not entered
else {
    $Name = ($AppName.ToUpper()).Trim()
    if ($Name.StartsWith("CEZ_")) {
        $Name = $Name.Substring(4)
    }
    $AADSPDisplayName = "CEZ_$($Name)"

    foreach ($SP in $AADSPDB_ByAppId.Values) {
        if ($SP.DisplayName.ToUpper() -contains $Name) {
            write-host "found existing service principal " -NoNewline
            write-host $SP.DisplayName -NoNewline -ForegroundColor Green
            write-host " with appId " -NoNewline
            write-host $SP.appId -ForegroundColor Green
            if ((Get-YesNoKeyboardInput "Continue and use this AAD service principal?")) {
                $AADSPAppId = $SP.appId
                $AADSPObjectId = $SP.id
                write-host $AADSPAppId -ForegroundColor Green
                write-host $AADSPObjectId -ForegroundColor Green
                break
            }
            Else {
                exit
            }
        }
    }
    if (-not($AADSPAppId -and $AADSPObjectId)) {
        write-host "service principal with display name $($AADSPDisplayName) not found in AAD" -ForegroundColor DarkYellow
        if ((Get-YesNoKeyboardInput "Create it?")) {
            Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
            $UriResource = "applications"
            $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
            $Body = @{
                displayName = $AADSPDisplayName
            } | ConvertTo-Json
            Try {
                write-host "Creating new app $($AADSPDisplayName) in AAD with display name $($AADSPDisplayName)..." -NoNewline
                $AADApp = Invoke-WebRequest -Uri $Uri -Method "POST" -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Body $Body -ContentType $ContentTypeJSON -ErrorAction Stop | ConvertFrom-Json
                write-host "OK" -ForegroundColor Green
                write-host "Checking app availablity" -NoNewline
                for ($i=0; $i -lt 5; $i++) {
                    write-host "." -nonewline
                    Start-Sleep 1
                }
                $AADSPAppId = $AADApp.appId
                write-host "OK" -ForegroundColor Green
                
                $UriResource = "servicePrincipals"
                $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
                $Body = @{
                    appId = $AADSPAppId
                } | ConvertTo-Json
                Try {
                    write-host "Creating new service principal $($AADSPDisplayName) in AAD..." -NoNewline
                    $AADSP = Invoke-WebRequest -Uri $Uri -Method "POST" -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Body $Body -ContentType $ContentTypeJSON -ErrorAction Stop | ConvertFrom-Json
                    $AADSPDisplayName = $AADSP.DisplayName
                    $AADSPObjectId = $AADSP.id
                    write-host "OK" -ForegroundColor Green
                }
                Catch {
                    write-host "Failed to create new service principal in AAD" -ForegroundColor Red
                    exit
                }
            }
            Catch {
                write-host "Failed to create new app in AAD" -ForegroundColor Red
                $ErrorMessage = $_.Exception.Message
                $errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                Write-Host "$($errObj.error.code) - $ErrorMessage" -ForegroundColor Red
                exit
            }  
        }
        Else {
            exit
        }
    }
}

#############################################################################################################################
#Scoped vs. not scoped assigment
Write-CurrentValues
if ((Get-YesNoKeyboardInput "Will this be a scoped role assignment?")) {
    $Scoped = $true
}
Else {
    $Scoped = $false
}

#############################################################################################################################
#EXO group

if ($Scoped) {

    Write-CurrentValues
    if (-not $DGName) {
        $DGName = "CEZ_AAD_Scope_MBX_$($Name)"
    }
    write-host "Exchange Online: " -ForegroundColor Cyan -NoNewline
    Try {
        $ExistingDG = Get-DistributionGroup -Identity $DGName -ErrorAction SilentlyContinue
    }
    Catch {
        $ExistingDG = $null
    }
    if ($ExistingDG) {
        write-host "found existing group " -NoNewline
        write-host "$($ExistingDG.DisplayName) " -NoNewline -ForegroundColor Green
        write-host "($($ExistingDG.ExternalDirectoryObjectId))"
        if ((Get-YesNoKeyboardInput "Continue and use this group?")) {
            $DG = $ExistingDG
        }
        Else {
            exit
        }
    }
    else {
        write-host "will create new mail-enabled security group " -NoNewline
        write-host $DGName -ForegroundColor Green
        if ((Get-YesNoKeyboardInput "Proceed?")) {
            write-host "Creating group..." -NoNewline
            Try {
                New-DistributionGroup -Name $DGName -Type "Security" | out-null
                write-host "OK" -foregroundColor Green
            }
            Catch {
                write-host "Failed to create group $DGName" -ForegroundColor Red
                exit
            }
            write-host "Checking group availablity" -NoNewline
            for ($i=0; $i -lt 10; $i++) {
                write-host "." -nonewline
                Start-Sleep 1
            }
            for ($i=0; $i -lt 50; $i++) {
                Try {
                    $DG = Get-DistributionGroup -Identity $DGName
                    write-host "OK" -NoNewline -ForegroundColor Green
                    break
                }
                Catch {
                    write-host "." -nonewline
                    Start-Sleep 2
                }
            }
        }
        Else {
            exit
        }
    }
    if ($DG) {
        $DGName = $DG.Name
        $DGDN = $DG.DistinguishedName
        $DGGuid = $DG.ExternalDirectoryObjectId
    }
    else {
        write-host "Failed to create group $DGName" -ForegroundColor Red
        exit
    }
}
else {
    $DG = $null
    $DGName = $DGDN = $DGGuid = "n/a"
}

#############################################################################################################################
#EXO service principal

Write-CurrentValues

$EXOSPDisplayName = "EXOSP_$($AADSPDisplayName)"

write-host "EXO SP: " -ForegroundColor Cyan -NoNewline
Try {
    $ExistingEXOSPbyId = Get-ServicePrincipal -Identity $AADSPAppId -ErrorAction Stop
}
Catch {
    $ExistingEXOSPbyId = $null
}

Try {
    $ExistingEXOSPbyName = Get-ServicePrincipal -Identity $EXOSPDisplayName -ErrorAction Stop
}
Catch {
    $ExistingEXOSPbyName = $null
}

if ($ExistingEXOSPbyId) {
    write-host "service principal " -NoNewline
    write-host $ExistingEXOSPbyId.displayName -NoNewline -ForegroundColor Green    
    write-host " already exists ($($ExistingEXOSPbyId.ObjectId))"
    write-host "AAD SP appId    : " -NoNewline
    write-host $ExistingEXOSPbyId.appId -ForegroundColor Green
    write-host "AAD SP object Id: " -NoNewline
    write-host $ExistingEXOSPbyId.objectId -ForegroundColor Green
    

    if ($ExistingEXOSPbyId.DisplayName -eq $EXOSPDisplayName) {
        #all good, existing SP matches appId
        if ((Get-YesNoKeyboardInput "Continue and use this EXO service principal?")) {
            $EXOSP = $ExistingEXOSPbyId
        }
        Else {
            Exit
        }
    }
    else {
        write-host "Does not match expected name " -NoNewline -ForegroundColor Red
        write-host $EXOSPDisplayName
        exit
    }
}

if ((-not($EXOSP)) -and $ExistingEXOSPbyName) {
    write-host "service principal " -NoNewline
    write-host $ExistingEXOSPbyId.displayName -NoNewline -ForegroundColor Green    
    write-host " already exists ($($ExistingEXOSPbyId.ObjectId))"
    write-host "AAD SP appId    : " -NoNewline
    write-host $ExistingEXOSPbyId.appId -ForegroundColor Green
    write-host "AAD SP object Id: " -NoNewline
    write-host $ExistingEXOSPbyId.objectId -ForegroundColor Green
    if ($ExistingEXOSPbyName.appId -eq $AADSPAppId) {
        if ((Get-YesNoKeyboardInput "Continue and use this EXO service principal?")) {
            $EXOSP = $ExistingEXOSPbyName
        }
        Else {
            write-host "Does not match expected AAD appId " -NoNewline -ForegroundColor Red
            write-host $AADSPAppId
            Exit
        }
    }
}

if (-not $EXOSP) {
    # create new EXO Service Principal
    write-host "will create EXO service principal " -NoNewline
    write-host $EXOSPDisplayName -ForegroundColor Green
    if ((Get-YesNoKeyboardInput "Proceed?")) {
        write-host "Creating EXO service principal..." -NoNewline
        Try {
            New-ServicePrincipal -AppId $AADSPAppId -ServiceId $AADSPObjectId -DisplayName $EXOSPDisplayName -ErrorAction Stop | out-null
            write-host "OK" -ForegroundColor Green
        }
        Catch {
            write-host "Failed to create EXO service principal" -ForegroundColor Red
            exit
        }
        write-host "Checking EXO service principal availablity" -NoNewline
        for ($i=0; $i -lt 5; $i++) {
            write-host "." -nonewline
            Start-Sleep 1
        }
        for ($i=0; $i -lt 50; $i++) {
            Try {
                $EXOSP = Get-ServicePrincipal -Identity $EXOSPDisplayName
                if ($EXOSP) {
                    write-host "OK" -ForegroundColor Green
                    break
                }
            }
            Catch {
                write-host "." -nonewline
                Start-Sleep 2
            }
        }
    }
    Else {
        Exit
    }
}
write-host
if ($EXOSP) {
    $EXOSPDisplayName = $EXOSP.DisplayName
}
else {
    write-host "Failed to create EXO SP" -ForegroundColor Red
    exit
}

#############################################################################################################################
# EXO management scope
if ($Scoped) {
    Write-CurrentValues

    $EXOMgmtScopeName = "Scope_MemberOf_$($DGGuid)"
    
    write-host "EXO mgmt scope: " -ForegroundColor Cyan -NoNewline
    Try {
        $ExistingEXOMgmtScope = Get-ManagementScope -Identity $EXOMgmtScopeName -ErrorAction Stop
    }
    Catch {
        $ExistingEXOMgmtScope = $null
    }
    if ($ExistingEXOMgmtScope) {
        write-host "management scope " -NoNewline
        write-host $ExistingEXOMgmtScope.name -NoNewline -ForegroundColor Green    
        write-host " already exists"
        if ((Get-YesNoKeyboardInput "Continue and use this EXO management scope?")) {
            $EXOMgmtScope = $ExistingEXOMgmtScope
        }
        Else {
            Exit
        }
    }
    else {
        # create new EXO management scope 
        write-host "will create EXO management scope " -NoNewline
        write-host $EXOMgmtScopeName -ForegroundColor Green
        if ((Get-YesNoKeyboardInput "Proceed?")) {
            write-host "Creating EXO management scope..." -NoNewline
            Try {
                New-ManagementScope -Name $EXOMgmtScopeName -RecipientRestrictionFilter "MemberOfGroup -eq '$($DG.DistinguishedName)'" | out-null
                write-host "OK" -ForegroundColor Green
            }
            Catch {
                write-host "Failed to create EXO management scope" -ForegroundColor Red
                exit
            }
            write-host "Checking EXO management scope availablity" -NoNewline
            for ($i=0; $i -lt 5; $i++) {
                write-host "." -nonewline
                Start-Sleep 1
            }
            for ($i=0; $i -lt 50; $i++) {
                Try {
                    $EXOMgmtScope = Get-ManagementScope -Identity $EXOMgmtScopeName
                    if ($EXOMgmtScope) {
                        write-host "OK" -ForegroundColor Green
                        break
                    }
                }
                Catch {
                    write-host "." -nonewline
                    Start-Sleep 2
                }
            }
        }
        Else {
            Exit
        }
    }
    write-host
    if (-not($EXOMgmtScope)) {
        write-host "Failed to create EXO management scope" -ForegroundColor Red
        exit
    }
}
else {
    $EXOMgmtScope = $null
    $EXOMgmtScopeName = "n/a"
}



#############################################################################################################################
Write-CurrentValues

$EXOMgmtRoleAssignments = Get-ManagementRoleAssignment -RoleAssigneeType "ServicePrincipal"
foreach ($EXOMgmtRoleAssignment in $EXOMgmtRoleAssignments) {
    if ($EXOMgmtRoleAssignment.RoleAssignee -eq $AADSPObjectId) {
        $RAObject = [PSCustomObject]@{
            Id = $EXOMgmtRoleAssignment.Id
            Role = $EXOMgmtRoleAssignment.Role                
            CustomResourceScope = $EXOMgmtRoleAssignment.CustomResourceScope 
        }
        $ExistingRAs += $RAObject
    }
}
if ($ExistingRAs.Count -gt 0) {
    Write-TableHeader -String "Existing role assignments:" -TextColor "Cyan" -LineColor "White" -Width $TableWidthValues
    foreach ($RA in $ExistingRAs) {
        Write-TableLine -Key $RA.Role -Value $RA.CustomResourceScope -Width $TableWidthValues -KeyWidth $KeyWidthLongValues
    }
    Write-TableFooter -Width $TableWidthValues

    Write-TableHeader -String "Select action:" -TextColor "Yellow" -LineColor "White" -Width $TableWidthRoles
    Write-TableLine -Key "A" -Value "Add role assignment" -Width $TableWidthRoles -KeyWidth $KeyWidthRoles -KeyColor "Yellow" -ValueColor "White"
    Write-TableLine -Key "R" -Value "Replace existing role assignments" -Width $TableWidthRoles -KeyWidth $KeyWidthRoles -KeyColor "Yellow" -ValueColor "White"
    Write-TableLine -Key "X" -Value "Exit" -Width $TableWidthRoles -KeyWidth $KeyWidthRoles -KeyColor "Yellow" -ValueColor "White"
    Write-TableFooter -Width $TableWidthRoles

    $RoleAction = "assign"
    while ($true) {
        switch (Get-AnyKeyboardInput -Prompt "Select action:") {
            'A' { $RoleAction = "add" }
            'R' { $RoleAction = "replace" }
            'X' { Exit }
            default { write-host "Invalid selection" -ForegroundColor Red }
        }
        if ($RoleAction) {
            break
        }
    }
    if ($RoleAction -eq "replace") {
        write-host "Deleting existing role assignments..." -NoNewline
        foreach ($RA in $ExistingRAs) {
            Remove-ManagementRoleAssignment -Identity $RA.Id
            write-host "." -NoNewline
            Start-Sleep 3
        }
        Write-Host "done"
    }
}

Clear-Host

Write-TableHeader -String "Select role to $($RoleAction) to the service principal:" -TextColor "Yellow" -LineColor "White" -Width $TableWidthRoles
Write-TableLine -Key "A" -Value "Mail.Read" -Width $TableWidthRoles -KeyWidth $KeyWidthRoles -KeyColor "Yellow" -ValueColor "White"
Write-TableLine -Key "B" -Value "Mail.ReadBasic" -Width $TableWidthRoles -KeyWidth $KeyWidthRoles -KeyColor "Yellow" -ValueColor "White"
Write-TableLine -Key "C" -Value "Mail.ReadWrite" -Width $TableWidthRoles -KeyWidth $KeyWidthRoles -KeyColor "Yellow" -ValueColor "White"
Write-TableLine -Key "D" -Value "Mail.Send" -Width $TableWidthRoles -KeyWidth $KeyWidthRoles -KeyColor "Yellow" -ValueColor "White"
Write-TableLine -Key "E" -Value "MailboxSettings.Read" -Width $TableWidthRoles -KeyWidth $KeyWidthRoles -KeyColor "Yellow" -ValueColor "White"
Write-TableLine -Key "F" -Value "MailboxSettings.ReadWrite" -Width $TableWidthRoles -KeyWidth $KeyWidthRoles -KeyColor "Yellow" -ValueColor "White"
Write-TableLine -Key "G" -Value "Calendars.Read" -Width $TableWidthRoles -KeyWidth $KeyWidthRoles -KeyColor "Yellow" -ValueColor "White"
Write-TableLine -Key "H" -Value "Calendars.ReadWrite" -Width $TableWidthRoles -KeyWidth $KeyWidthRoles -KeyColor "Yellow" -ValueColor "White"
Write-TableLine -Key "I" -Value "Contacts.Read" -Width $TableWidthRoles -KeyWidth $KeyWidthRoles -KeyColor "Yellow" -ValueColor "White"
Write-TableLine -Key "J" -Value "Contacts.ReadWrite" -Width $TableWidthRoles -KeyWidth $KeyWidthRoles -KeyColor "Yellow" -ValueColor "White"
Write-TableLine -Key "K" -Value "Mail Full Access" -Width $TableWidthRoles -KeyWidth $KeyWidthRoles -KeyColor "Yellow" -ValueColor "White"
Write-TableLine -Key "L" -Value "Exchange Full Access" -Width $TableWidthRoles -KeyWidth $KeyWidthRoles -KeyColor "Yellow" -ValueColor "White"
write-TableLine -Key "M" -Value "EWS.AccessAsApp" -Width $TableWidthRoles -KeyWidth $KeyWidthRoles -KeyColor "Yellow" -ValueColor "White"
write-TableLine -Key "N" -Value "ApplicationImpersonation" -Width $TableWidthRoles -KeyWidth $KeyWidthRoles -KeyColor "Yellow" -ValueColor "White"
write-TableLine -Key "X" -Value "Exit" -Width $TableWidthRoles -KeyWidth $KeyWidthRoles -KeyColor "Magenta" -ValueColor "Magenta"
write-TableFooter -Width $TableWidthRoles

$ExoRole = $null
while ($true) {
    switch (Get-AnyKeyboardInput -Prompt "Select role:") {
        'A' { $EXORole = "Application Mail.Read" }
        'B' { $EXORole = "Application Mail.ReadBasic" }
        'C' { $EXORole = "Application Mail.ReadWrite" }
        'D' { $EXORole = "Application Mail.Send" }
        'E' { $EXORole = "Application MailboxSettings.Read" }
        'F' { $EXORole = "Application MailboxSettings.ReadWrite" }
        'G' { $EXORole = "Application Calendars.Read" }
        'H' { $EXORole = "Application Calendars.ReadWrite" }
        'I' { $EXORole = "Application Contacts.Read" }
        'J' { $EXORole = "Application Contacts.ReadWrite" }
        'K' { $EXORole = "Application Mail Full Access" }
        'L' { $EXORole = "Application Exchange Full Access" }
        'M' { $EXORole = "Application EWS.AccessAsApp" }
        'N' { $EXORole = "ApplicationImpersonation" }
        'X' { Exit }
        default { write-host "Invalid selection" -ForegroundColor Red }
    }
    if ($EXORole) {
        break
    }
}

Write-CurrentValues

if ((Get-YesNoKeyboardInput "Proceed with role assignment?")) {
    if ($Scoped) {
        New-ManagementRoleAssignment -App $AADSPAppId -Role $ExoRole -CustomResourceScope $EXOMgmtScopeName
    }
    else {
        New-ManagementRoleAssignment -App $AADSPAppId -Role $ExoRole
    }   
}
else {
    Exit
}

write-host "Applying role assignment" -NoNewline
for ($i=0; $i -lt 10; $i++) {
    write-host "." -nonewline
    Start-Sleep 1
}

Write-CurrentRoleAssignments -ObjectId $AADSPObjectId

