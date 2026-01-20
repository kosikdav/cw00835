#######################################################################################################################
# Set-SPO-App-Permissions
#######################################################################################################################
param(
    [parameter(Mandatory = $true)][string]$Url,
	[ValidateSet("read","write","fullcontrol","manage")]$Role="read",
	[parameter(Mandatory = $true)][string]$ApplicationId,
    [string]$DisplayName,
    [switch]$Delete,
    [switch]$Force
)

$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-Start-Generic.ps1
. $ScriptPath\include-Script-StdIncBlock.ps1
. $ScriptPath\include-Functions-Common.ps1

[hashtable]$Site_DB = @{}

#######################################################################################################################

$SiteId = [string]::Empty

Request-MSALToken -AppRegName $AppReg_SPO_MGMT -TTL 30
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

#read all sites from SPO and check if site exists
$UriResource = "sites"
$UriSelect = "id,webUrl,name"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect -Top 999
[array]$Sites = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_SPO_MGMT].AccessToken -ContentType $ContentTypeJSON -Text "Getting SPO sites" -ProgressDots
foreach ($Site in $Sites) {
	if ($Site.webUrl -eq $Url) {
        $SiteId = $Site.id
        write-host "Site:        $Url ($($Site.name))" -ForegroundColor Yellow
        write-host "Site id:     $SiteId" -ForegroundColor Gray
    }
}
if ($SiteId -eq [string]::Empty) {
	write-host "Site $Url not found" -ForegroundColor Red
	Exit
}

#look up application registered in AAD
$UriResource = "applications(appId='$ApplicationId')"
$UriSelect = "id,appId,displayName"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
$Application = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Silent
if ($null -eq $Application) {
	$UriResource = "applications/$ApplicationId"
    $UriSelect = "id,appId,displayName"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
	$Application = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Silent
	if ($null -eq $Application) {
        $UriResource = "servicePrincipals/(appId='$ApplicationId')"
        $UriSelect = "id,appId,displayName"
	    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
	    $Application = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Silent
        if ($null -eq $Application) {
            $UriResource = "servicePrincipals/$ApplicationId"
            $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
            $Application = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Silent
        }
        if ($null -eq $Application) {
            $UriResource = "servicePrincipals"
            $UriSelect = "id,appId,displayName"
            $UriFilter = "servicePrincipalType eq 'ManagedIdentity' and appId eq '$ApplicationId'"
            $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect -Filter $UriFilter
            $Application = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Silent
            if ($null -eq $Application) {
                write-host "No application with id=$($ApplicationId) (applicationId or objectId) found" -ForegroundColor Red
                Exit
            }
        }
	}
}

Write-Host "Application: $($Application.displayName) $($Application.appId) (objId: $($Application.id))" -ForegroundColor Yellow
Write-Host "Role:        $($Role)" -ForegroundColor Yellow

#set URI for permissions
$UriResource = "sites/$SiteId/permissions"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource

# check if permission already exists
$ExistingPermissions = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_SPO_MGMT].AccessToken -ContentType $ContentTypeJSON
if ($ExistingPermissions.Count -gt 0) {
    foreach ($Permission in $ExistingPermissions) {
        if ($Permission.grantedToIdentitiesV2.application.id -eq $Application.appId) {
            write-host "Permission already exists - $($Permission.grantedToIdentitiesV2.application.id) $($Permission.grantedToIdentitiesV2.application.displayName)" -ForegroundColor DarkYellow
            if ($Delete) {
                #ask for confirmation
                if (Get-YesNoKeyboardInput -Prompt "Really delete permission?" -ForegroundColor White) {
                    $UriResource = "sites/$SiteId/permissions/$($Permission.id)"
                    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
                    #delete permission
                    Try {
                        $ResponseDELETE = Invoke-WebRequest -Headers $AuthDB[$AppReg_SPO_MGMT].AuthHeaders -Uri $Uri -Body $GraphBody -Method "DELETE" -ContentType $ContentTypeJSON
                        write-host "-------------------------------------------------------------" -ForegroundColor Yellow
                        write-host "Permission deleted: $($Permission.grantedToIdentitiesV2.application.id) $($Permission.grantedToIdentitiesV2.application.displayName)" -ForegroundColor Green
                    }
                    Catch {
                        $ErrorMessageDELETE = $_.Exception.Message
                        $errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                        Write-Host "$($errObj.error.code) - $ErrorMessageDELETE" -ForegroundColor Red
                    }
                }
                Exit
            }
            else {
                if (-not $Force) {
                    write-host "Use -Force to overwrite" -ForegroundColor Red
                    Exit
                }
            }
        }
    }
}

if (Get-YesNoKeyboardInput -Prompt "Really set permission?" -ForegroundColor White) {
    #set timestamp and permission display name
    $TimeStamp = Get-Date -Format "yyyyMMdd-HHmmss"
    Write-Host "Timestamp:   $($TimeStamp)" -ForegroundColor Yellow
    $DisplayName = $Application.displayName + "_" + $Application.appId + "_" + $Role + "_" + $TimeStamp

    #build Grpah body JSON
    $GraphBody = @{
    "roles" = @(
        $Role
    )
    "grantedToIdentities" = @(
        @{
            "application" = @{
                "id" = $Application.appId
                "displayName" = $DisplayName
            }
        }
    )
    } | ConvertTo-Json -Depth 3

    #set permissions
    Try {
    $ResponsePOST = Invoke-WebRequest -Headers $AuthDB[$AppReg_SPO_MGMT].AuthHeaders -Uri $Uri -Body $GraphBody -Method "POST" -ContentType $ContentTypeJSON
    write-host "-------------------------------------------------------------" -ForegroundColor Yellow
    write-host "Permission set: $($displayName)" -ForegroundColor Green
    }
    Catch {
    $ErrorMessagePOST = $_.Exception.Message
    $errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
    Write-Host "$($errObj.error.code) - $ErrorMessagePOST" -ForegroundColor Red
    }
}


