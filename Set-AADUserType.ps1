param(
    [Parameter(Mandatory = $true)][string]$Identity,
    [Parameter(Mandatory = $true)][string][ValidateSet("Guest","Member")]$UserType
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1
. $ScriptPath\include-Script-StdIncBlock.ps1

Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
$UriResource = "users/$($Identity)"
$UriSelect = "mail,userPrincipalName,displayName,userType"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
$User = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -ContentType $ContentTypeJSON
#write-host $User
if ($User) {
    if  ($User.userType -ne $UserType) {
        write-host "Found $($User.userPrincipalName) - $($User.userType), changing to $($UserType)" -ForegroundColor "Yellow"
        $Body = @{
            userType = $UserType
        } | ConvertTo-Json
        $Body = [System.Text.Encoding]::UTF8.GetBytes($Body)
        Try {
            $ResultPATCH = Invoke-RestMethod -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $Body -Method "PATCH" -ContentType $ContentTypeJSON
            Write-Host "$($User.userPrincipalName) converted to userType $($UserType)" -ForegroundColor "Green"
        }
        Catch {
            $ErrorMessagePATCH = $_.ErrorDetails.Message | Out-String
            Write-Host "ERR PATCH userType to $($UserType)" -MessageType Error
            Write-Host $($ErrorMessagePATCH) -MessageType Error
        }
    }
    else {
        Write-Host "$($User.userPrincipalName) already userType $($UserType)" -ForegroundColor "Green"
    }
}
else {
    Write-Host "User with identity $($Identity) not found" -ForegroundColor "Red"
}
