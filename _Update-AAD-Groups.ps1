#######################################################################################################################
# Update-AAD-Groups
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder          = "aad-group-mgmt"
$LogFilePrefix      = "aad-group-mgmt"
$LogFileFreq        = "YMD"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Suffix $LogFileSuffix -Ext "log"

$ADCredentialPath = $aad_grp_mgmt_cred

[array]$ASISAzureSubOwnersList = @()
[array]$ASISAzureSubOwnersStdList = @()
[array]$TOBEAzureSubOwnersList = @()

#######################################################################################################################

. $IncFile_StdLogStartBlock

#TAGS##########################################

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30

$UriResource = "teams/$($GroupId_TMS_CloudPoradna)/tags"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
$Tags = Get-GraphOutputREST -Uri $Uri -ContentType $ContentTypeJSON -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
$TeamsTagId = [string]::Empty
foreach ($tag in $Tags) {
    if ($Tag.displayName -eq "Microsoft Azure") {
        $TeamsTagId = $Tag.id
        Write-Log "Teams tag `"Microsoft Azure`" found"
    }
}

$UriResource = "teams/$($GroupId_TMS_CloudPoradna)/tags/$($TeamsTagId)/members"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
$userid = "2d847bc0-ce2f-45c3-a6fd-86a98c5ea8dd"
$GraphBody = [PSCustomObject]@{
    UserId = $userId;
} | ConvertTo-Json
write-host $uri
write-host $GraphBody
Try {
    $resultAddTag = Invoke-WebRequest -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $GraphBody -Method "POST" -ContentType $ContentTypeJSON
}
Catch {
    write-host $_.Exception.Message
    $errObj = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
    Write-Host "$($errObj.error.code)"

}


exit

$ASISAzureSubOwners = Get-GroupMembersFromGraphById -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -id $GroupId_AzureSubOwnersGroup
foreach ($user in $ASISAzureSubOwners) {
    if ($AADUserUPN_DB.ContainsKey($user.userPrincipalName)) {
        $userRecord = $AADUserUPN_DB[$user.userPrincipalName]
        if ($userRecord.msExchExtensionAttribute40) {
            if ($AADUserMail_DB.ContainsKey($userRecord.msExchExtensionAttribute40)) {
                $standardUser = $AADUserMail_DB[$userRecord.msExchExtensionAttribute40]
                $ASISAzureSubOwnersStdList += $standardUser.Id
            }
        }
    }
}


$UriResource = "teams/$($GroupId_TMS_CloudPoradna)/tags"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
$Tags = Get-GraphOutputREST -Uri $Uri -ContentType $ContentTypeJSON -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
$TeamsTagId = [string]::Empty
foreach ($tag in $Tags) {
    if ($Tag.displayName -eq "Microsoft Azure") {
        $TeamsTagId = $Tag.id
        Write-Log "Teams tag `"Microsoft Azure`" found"
    }
}

$ASISTMS_CloudPoradna = (Get-GroupMembersFromGraphById -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -id $GroupId_TMS_CloudPoradna).id
Write-Log "ASISTMS_CloudPoradna: $($ASISTMS_CloudPoradna.Count)"
$UriResource = "teams/$($GroupId_TMS_CloudPoradna)/tags/$($TeamsTagId)/members"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
$ASISTMS_CloudPoradnaTagMembers = (Get-GraphOutputREST -Uri $Uri -ContentType $ContentTypeJSON -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken).userId
Write-Log "ASISTMS_CloudPoradnaTagMembers: $($ASISTMS_CloudPoradnaTagMembers.Count)"

foreach ($userId in $ASISAzureSubOwnersStdList) {
    if ($ASISTMS_CloudPoradna.Contains($userId)) {
        Write-Log "User $userId is in Team"
    }
    if ($ASISTMS_CloudPoradnaTagMembers.Contains($userId)) {
        Write-Log "User $userId has tag"
    }
    
    if (($ASISTMS_CloudPoradna.Contains($userId)) -and (-Not($ASISTMS_CloudPoradnaTagMembers.Contains($userId)))) {
        $missingUser = Get-GraphUserById -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -id $userId
        $GraphBody = [PSCustomObject]@{
            UserId = $userId;
        } | ConvertTo-Json
        Try {
            Write-Log "Adding tag to $($userId) $($missingUser.userPrincipalName)"
            $resultAddTag = Invoke-WebRequest -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $GraphBody -Method "POST" -ContentType $ContentTypeJSON
        }
        Catch {
            Write-Log "Error adding $($userId) $($missingUser.userPrincipalName)" -MessageType "Error"
        }
    }
}
#######################################################################################################################

. $IncFile_StdLogEndBlock
