$tenantId     = "b233f9e1-5599-4693-9cef-38858fe25406"
$clientId     = "eae1a90a-5a4f-42c1-a780-fa8512c58e77"
$clientSecret = "uT08Q~PwWPkaoy-NZZf8gzaIpxflqIBFbQaqkdjq"
[array]$AADUsers = @()
[hashtable]$AADUsers_DB = @{}

$scope = "https://graph.microsoft.com/.default"
$tokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $clientId    
    client_secret = $clientSecret    
    scope         = $scope    
    grant_type    = "client_credentials"
}

try {
    $response = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ContentType "application/x-www-form-urlencoded"    
    $accessToken = $response.access_token    
    Write-Host "Access Token acquired successfully!"
}
catch {
    Write-Host "Failed to obtain access token:" -ForegroundColor Red    
    Write-Host $_
}

$Uri = "https://graph.microsoft.com/v1.0/users?`$select=id,userPrincipalName,samaccountname,mail,department,companyName&`$filter=userType+eq+'Member'&`$top=999"
$Headers = @{
    Authorization = "Bearer $accessToken"
}

Write-Host "Retrieving users from Microsoft Entra ID." -NoNewline
do {
    Write-Host "." -NoNewline
    $response = Invoke-RestMethod -Method Get -Uri $Uri -Headers $Headers
    $AADUsers += $response.value
    $Uri = $response.'@odata.nextLink'
} while ($Uri)
Write-Host "done ($($AADUsers.count))"

Write-Host "Creating AAD user database....." -NoNewline
foreach ($user in $AADUsers) {
    $userObject = [PSCustomObject]@{
        Id                  = $user.id
        UserPrincipalName   = $user.userPrincipalName
        Mail                = $user.mail
        SamAccountName      = $user.samaccountname
        Department          = $user.department
        CompanyName         = $user.companyName
    }
    $AADUsers_DB.Add($user.userPrincipalName, $userObject)
}
Write-Host "done ($($AADUsers_DB.Count))"


