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

$ASISExtSharingAllowed = @()
$TOBEExtSharingAllowed = @()

[array]$ASISAzureSubOwnersList = @()
[array]$ASISAzureSubOwnersStdList = @()
[array]$TOBEAzureSubOwnersList = @()

$TMS_CloudPoradna_initialRun = $false

$DB_changed = $false
$missingTagFound = $false

#######################################################################################################################

. $IncFile_StdLogStartBlock

if (-not $interactiveRun) {
    Write-Log "AD credential file: $($ADCredentialPath)"
    $ADCredential = Import-Clixml -Path $ADCredentialPath
}

Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30

$AzureSubOwnersGroup_Name = Get-GroupNameFromGraphById -id $GroupId_AzureSubOwnersGroup -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
$LicMgmtQSGroup_Name = Get-GroupNameFromGraphById -id $CEZ_QS_LicMgmt_Team -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
$OCPDevOpsTeam_Name = Get-GroupNameFromGraphById -id $GroupId_OCP_DevOpsTeam -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken


##########################################################################################
#CEZ_AZURE_SUBSCRIPTION_OWNERS ###########################################################
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30

Write-Log $string_Divider
write-log "Processing group $($AzureSubOwnersGroup_Name) membership"

$UriResource = "groups/$($GroupId_AzureSubOwnersGroup)/members"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
[array]$ASISAzureSubOwners = (Get-GraphOutputREST -Uri $Uri -ContentType $ContentTypeJSON -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken).id

$UriResource = "groups/$($GroupId_AzureSubOwnersQS)/members"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
[array]$TOBEAzureSubOwners = (Get-GraphOutputREST -Uri $Uri -ContentType $ContentTypeJSON -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken).id

$UriResource = "groups/$($GroupId_AzureBSIAdmins)/members"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
[array]$TOBEAzureSubOwnersBSI = (Get-GraphOutputREST -Uri $Uri -ContentType $ContentTypeJSON -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken).id

$TOBEAzureSubOwners = $TOBEAzureSubOwners + $TOBEAzureSubOwnersBSI

$ASISAzureSubOwners = $ASISAzureSubOwners | Sort-Object -Unique
$TOBEAzureSubOwners = $TOBEAzureSubOwners | Sort-Object -Unique

write-host "Current $($AzureSubOwnersGroup_Name) members: $($ASISAzureSubOwners.count)"
write-host "Target  $($AzureSubOwnersGroup_Name) members: $($TOBEAzureSubOwners.count)"


Try {
	$Difference = Compare-Object -ReferenceObject $ASISAzureSubOwners -DifferenceObject $TOBEAzureSubOwners
}
Catch {
	$Difference = $true
}

if ($Difference) {
    $missingMembers = $TOBEAzureSubOwners | Where-Object { $ASISAzureSubOwners -notcontains $_ }
	write-log "Missing members: $($missingMembers.count)"
    $extraMembers = $ASISAzureSubOwners | Where-Object { $TOBEAzureSubOwners -notcontains $_ }
	write-log "Extra members: $($extraMembers.count)"
	if ($missingMembers) {
        foreach ($id in $missingMembers) {
            Write-Log "Adding $($id) to $($AzureSubOwnersGroup_Name)"
            Add-GraphGroupMemberById -GroupId $GroupId_AzureSubOwnersGroup -userId $id -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -SkipCurrentMembers:$true
        }
    }
    if ($extraMembers) {
        foreach ($id in $extraMembers) {
            Write-Log "Removing $($id) from $($AzureSubOwnersGroup_Name)"
            Remove-GraphGroupMemberById -GroupId $GroupId_AzureSubOwnersGroup -userId $id -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken
        }
    }
}

Remove-Variable -Name ASISAzureSubOwners
Remove-Variable -Name TOBEAzureSubOwners
Remove-Variable -Name TOBEAzureSubOwnersBSI

##########################################################################################
#OCP DevOps Team  ########################################################################
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
[array]$TOBEOCPMembers = @()
[array]$AllOCPMembers = @()
Write-Log $string_Divider
Write-Log "Processing $($OCPDevOpsTeam_Name)"
$NamePattern = "OCP_*_ADMIN"
$GroupFilter = "Name -like '$($NamePattern)'"

if ($ADCredential) {
    $ADGroups = Get-ADGroup -Credential $ADCredential -Filter $GroupFilter -SearchBase $OU_OCP
}
else {
    $ADGroups = Get-ADGroup -Filter $GroupFilter -SearchBase $OU_OCP
}

$UriResource = "groups/$($GroupId_OCP_DevOpsTeam)/members"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
[array]$ASISOCPMembers = (Get-GraphOutputREST -Uri $Uri -ContentType $ContentTypeJSON -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken).id
write-host "Current $($OCPDevOpsTeam_Name) members: $($ASISOCPMembers.count)"

foreach ($group in $ADGroups) {
    write-host "Processing group $($group.samaccountname) " -NoNewline
    if ($ADCredential) {
        $members  = Get-ADGroupMember -Identity $group.samaccountname -Credential $ADCredential | ForEach-Object { Get-ADUser $_.samaccountname -Credential $ADCredential -Properties "msDS-ExternalDirectoryObjectId" | Select-Object objectGUID,userPrincipalName,samAccountName,msDS-ExternalDirectoryObjectId }
    }
    else {
        $members  = Get-ADGroupMember -Identity $group.samaccountname | ForEach-Object { Get-ADUser $_.samaccountname -Properties "msDS-ExternalDirectoryObjectId" | Select-Object objectGUID,userPrincipalName,samAccountName,msDS-ExternalDirectoryObjectId }
    }
    $AllOCPMembers += $members
    write-host $members.count -ForegroundColor Green
}
$AllOCPMembers = $AllOCPMembers | Sort-Object -Property "objectGUID" -Unique

foreach ($member in $AllOCPMembers) {
    if ($NoExtSharingAccountPrefixes -Contains $member.samAccountName.Substring(0,2)) {
        continue
    }
    $id = ($member.'msDS-ExternalDirectoryObjectId').SubString(5)
    $TOBEOCPMembers += $id
}
$ASISOCPMembers = $ASISOCPMembers | Sort-Object -Unique
$TOBEOCPMembers = $TOBEOCPMembers | Sort-Object -Unique
write-host "Target  $($OCPDevOpsTeam_Name) members: $($TOBEOCPMembers.count)"

Try {
    $Difference = Compare-Object -ReferenceObject $ASISOCPMembers -DifferenceObject $TOBEOCPMembers
}
Catch {
    $Difference = $true
}

If ($Difference) {
    $missingMembers = $TOBEOCPMembers | Where-Object { $ASISOCPMembers -notcontains $_ }
    write-log "Missing members: $($missingMembers.count)"
    $extraMembers = $ASISOCPMembers | Where-Object { $TOBEOCPMembers -notcontains $_ }
    write-log "Extra members: $($extraMembers.count)"
    if ($missingMembers) {
        foreach ($id in $missingMembers) {
            Write-Log "Adding $($id) to $($OCPDevOpsTeam_Name)"
            Add-GraphGroupMemberById -GroupId $GroupId_OCP_DevOpsTeam -userId $id -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -SkipCurrentMembers:$true
        }
    }
}   

Remove-Variable -Name AllOCPMembers
Remove-Variable -Name TOBEOCPMembers
Remove-Variable -Name ASISOCPMembers

##########################################################################################
#LicCez_ owners set ######################################################################
Write-Log $String_Divider
Write-Log "Processing LicCEZ_ groups"

[array]$MembersLicMgmtQS = (Get-GroupMembersFromGraphById -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -id $CEZ_QS_LicMgmt_Team -Properties "id,userPrincipalName").id | Sort-Object
write-host "$($LicMgmtQSGroup_Name): $($MembersLicMgmtQS.count)"

$UriResource = "groups"
$UriSelect = "id,displayName,mail,onPremisesSyncEnabled,securityEnabled,groupTypes"
$UriFilter = "startswith(displayName,'LicCEZ_')+and+securityEnabled+eq+true+and+mailEnabled+eq+false"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect -Filter $UriFilter
[array]$LicGroups = Get-GraphOutputREST -Uri $Uri -ContentType $ContentTypeJSON -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
write-host "LicCEZ_ groups: $($LicGroups.count)"
foreach ($group in $LicGroups) {
    if (($Group.onPremisesSyncEnabled) -or ($Group.groupTypes -contains "DynamicMembership") -or ($Group.groupTypes -contains "Unified")) {
        continue        
    }
    Write-Host "Processing group $($group.displayName) $($group.id)"
    [array]$Owners = (Get-GroupOwnersFromGraphById -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -id $Group.id -Properties "id,userPrincipalName").id | Sort-Object
    Try {
        $Difference = Compare-Object -ReferenceObject $Owners -DifferenceObject $MembersLicMgmtQS
    }
    Catch {
        $Difference = $true
    }
    if ($Difference) {
        Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
        Update-GraphGroupOwnersById -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -TargetGroupId $Group.id -SourceGroupArray $MembersLicMgmtQS
    }
}

Remove-Variable -Name MembersLicMgmtQS
Remove-Variable -Name LicGroups


##########################################################################################
#CEZ_AAD_Usr_No_External_Sharing##########################################################
Write-Log $String_Divider
Write-Log "Processing CEZ_AAD_Usr_No_External_Sharing"

#source group - all SPO licensed users
[array]$MembersSPOLicensedUsers = Get-GroupMembersFromGraphById -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -id $GroupId_SPOLicensedUsers -Properties "id,userPrincipalName"
Write-Log "MembersSPOLicensedUsers: $($MembersSPOLicensedUsers.Count)"
#filter out users with specific prefixes
$MembersSPOLicensedUsers = ($MembersSPOLicensedUsers | Where-Object {$NoExtSharingAccountPrefixes -notcontains $_.userPrincipalName.Substring(0,2)}).id | Sort-Object -Unique
Write-Log "MembersSPOLicensedUsers (prefix filtered): $($MembersSPOLicensedUsers.Count)"

#block group - CEZ_AADSync_Ext_noMail
[array]$MembersExtNoMail = (Get-GroupMembersFromGraphById -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -id $GroupId_ExtNoMail -Properties "id,userPrincipalName").id | Sort-Object
Write-Log "MembersExtNoMail: $($MembersExtNoMail.Count)"

#block group - CEZ_AADSync_OU_Resources
[array]$MembersOUResources = (Get-GroupMembersFromGraphById -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -id $GroupId_OUResources -Properties "id,userPrincipalName").id | Sort-Object
Write-Log "MembersOUResources: $($MembersOUResources.Count)"

#block group - CEZ_AADSync_OU_SyncToAzure
[array]$MembersOUSyncToAzure = (Get-GroupMembersFromGraphById -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -id $GroupId_OUSyncToAzure -Properties "id,userPrincipalName").id | Sort-Object
Write-Log "MembersOUSyncToAzure: $($MembersOUSyncToAzure.Count)"

#block group - CEZ_AAD_Usr_No_External_Sharing
[array]$MembersNoExtSharing = Get-GroupMembersFromGraphById -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -id $GroupId_NoExtSharing -Properties "id,userPrincipalName"
Write-Log "No external sharing users: $($MembersNoExtSharing.Count)"
if ($MembersNoExtSharing.count -gt 0) {
    foreach ($user in $MembersNoExtSharing) {
        Write-Log "  $($user.userPrincipalName)" -foregroundcolor "Yellow"
    }
}
$MembersNoExtSharing = ($MembersNoExtSharing).id | Sort-Object

#current external sharing allowed users
$ASISExtSharingAllowed = (Get-GroupMembersFromGraphById -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -id $GroupId_ExtSharingAllowed).id | Sort-Object -Unique
Write-Log "ASIS ExtSharingAllowed users: $(Get-Count $ASISExtSharingAllowed)"

$TOBEExtSharingAllowed = $MembersSPOLicensedUsers
$TOBEExtSharingAllowed = $TOBEExtSharingAllowed | Where-Object { $MembersExtNoMail -notcontains $_ }
$TOBEExtSharingAllowed = $TOBEExtSharingAllowed | Where-Object { $MembersOUResources -notcontains $_ }
$TOBEExtSharingAllowed = $TOBEExtSharingAllowed | Where-Object { $MembersOUSyncToAzure -notcontains $_ }
$TOBEExtSharingAllowed = $TOBEExtSharingAllowed | Where-Object { $MembersNoExtSharing -notcontains $_ } 

# add whitelisted users
Write-Log "ExtSharingWhitelist users: $(Get-Count $ExtSharingWhitelist)"
$TOBEExtSharingAllowed = $TOBEExtSharingAllowed + $ExtSharingWhitelist

$TOBEExtSharingAllowed = $TOBEExtSharingAllowed | Sort-Object -Unique

Write-Log "TOBE ExtSharingAllowed users: $(Get-Count $TOBEExtSharingAllowed)"
Try {
    $Difference = Compare-Object -ReferenceObject $ASISExtSharingAllowed -DifferenceObject $TOBEExtSharingAllowed
}
Catch {
    $Difference = $true
}

if ($Difference) {
    Write-Log "Syncing..."
    Update-GraphGroupMembersById -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -TargetGroupId $GroupId_ExtSharingAllowed -SourceGroupArray $TOBEExtSharingAllowed -GroupType "User"
}
else {
    Write-Log "No changes detected"
}

Remove-Variable -Name MembersSPOLicensedUsers
Remove-Variable -Name MembersNoExtSharing
Remove-Variable -Name MembersExtNoMail
Remove-Variable -Name ASISExtSharingAllowed
Remove-Variable -Name TOBEExtSharingAllowed
#>

##########################################################################################
#TMS_CloudPoradna MEMBERS#################################################################
Write-Log $string_Divider
$TMS_CloudPoradna_Name = Get-GroupNameFromGraphById -id $GroupId_TMS_CloudPoradna -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
Write-Log "Processing $($TMS_CloudPoradna_Name) membership"

if (test-path $DBFileTMS_CloudPoradna) {
    Try {
        $TMS_CloudPoradna_DB = Import-Clixml -Path $DBFileTMS_CloudPoradna
        Write-Log "DB file $($DBFileTMS_CloudPoradna) imported successfully, $($TMS_CloudPoradna_DB.count) records found"
    } 
    Catch {
        Write-Log "Error importing $($DBFileTMS_CloudPoradna), creating empty DB" -MessageType "Error"
        [hashtable]$TMS_CloudPoradna_DB = @{}
    }
}
else {
    Write-Log "DB file $($DBFileTMS_CloudPoradna) not found, creating empty DB" -MessageType "Error"
    [hashtable]$TMS_CloudPoradna_DB = @{}
}


$ASISAzureSubOwners = (Get-GroupMembersFromGraphById -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -id $GroupId_AzureSubOwnersStd).id | Sort-Object -Unique
write-host "Current $($AzureSubOwnersGroup_Name) members: $($ASISAzureSubOwners.count)"

#get members of TMS_CloudPoradna group
$ASISTMS_CloudPoradna = (Get-GroupMembersFromGraphById -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -id $GroupId_TMS_CloudPoradna).id | Sort-Object -Unique
Write-Log "$($TMS_CloudPoradna_Name) current users: $($ASISTMS_CloudPoradna.count)"

#Initial run - fill database with existing users so they appeared as added/processed
if ($TMS_CloudPoradna_initialRun) {
    Write-Log "Initial run - fill database with existing users so they appeared as added/processed"
    foreach ($userId in $ASISTMS_CloudPoradna) {
        $missingUser = Get-GraphUserById -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -id $userId
        $newUserRecord = [PSCustomObject]@{
            UserId                  = $userId;
            UserPrincipalName       = $missingUser.userPrincipalName;
            DateTimeAdded           = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss");
            DateTimeFoundMissing    = [string]::Empty
        }
        $TMS_CloudPoradna_DB.Add($userId, $newUserRecord)
    }
    $DB_changed = $true
    Write-Log "TMS_CloudPoradna_DB initialized with $($TMS_CloudPoradna_DB.count) records"
    Start-Sleep -Seconds 60
}

$missingUsers = $ASISAzureSubOwners | Where-Object { $ASISTMS_CloudPoradna -notcontains $_ }

Write-Log "$($TMS_CloudPoradna_Name) missing users: $($missingUsers.count)"
foreach ($userId in $missingUsers) {
	$missingUser = Get-GraphUserById -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -id $userId
    if ($TMS_CloudPoradna_DB.ContainsKey($userId)) {
        $existingUserRecord = $TMS_CloudPoradna_DB[$userId]
        Write-Log "Skipping adding $($userId) $($missingUser.userPrincipalName) - user added previously $($existingUserRecord.DateTimeAdded)"
        if (-not($existingUserRecord.DateTimeFoundMissing)) {
            $existingUserRecord.DateTimeFoundMissing = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            $TMS_CloudPoradna_DB[$userId] = $existingUserRecord
            $DB_changed = $true
        }
    }
    else {
        $newUserRecord = [PSCustomObject]@{
            UserId                  = $userId;
            UserPrincipalName       = $missingUser.userPrincipalName;
            DateTimeAdded           = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss");
            DateTimeFoundMissing    = [string]::Empty
        }
        #$UriResource = "teams/$($GroupId_TMS_CloudPoradna)/tags/$($TeamsTagId)/members"
        #$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
        Try {
            Write-Log "Adding $($userId) $($missingUser.userPrincipalName)"
            Add-GraphGroupMemberById -AccessToken $AuthDB[$AppReg_USR_MGMT].AccessToken -GroupId $GroupId_TMS_CloudPoradna -userId $userId
            $TMS_CloudPoradna_DB.Add($userId, $newUserRecord)
            $DB_changed = $true
            $ASISAzureSubOwnersStdList+= $userId
        }
        Catch {
            Write-Log $_.Exception.Message -MessageType "Error"
            Write-Log "Error adding $($userId) $($missingUser.userPrincipalName)" -MessageType "Error"
        }
    }
}

#saving DB XML if needed
if (($TMS_CloudPoradna_DB.count -gt 0) -and ($DB_changed)){
    Try {
        $TMS_CloudPoradna_DB | Export-Clixml -Path $DBFileTMS_CloudPoradna
        Write-Log "DB file $($DBFileTMS_CloudPoradna) exported successfully, $($TMS_CloudPoradna_DB.count) records saved"
    }
    Catch {
        Write-Log "Error exporting $($DBFileTMS_CloudPoradna)" -MessageType "Error"
    }
}

##########################################################################################
#TMS_CloudPoradna TAGS####################################################################
Write-Log $string_Divider
Write-Log "Processing $($TMS_CloudPoradna_Name) tags"

#reading tag from team
$UriResource = "teams/$($GroupId_TMS_CloudPoradna)/tags"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
$Tags = Get-GraphOutputREST -Uri $Uri -ContentType $ContentTypeJSON -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken
Write-Log "Tags in $($TMS_CloudPoradna_Name): $($Tags.count)"
foreach ($tag in $Tags) {
    if ($Tag.displayName -eq $TMS_CloudPoradna_TagName_Azure) {
        #get tag members of TMS_CloudPoradna group
        Write-Log "Teams tag `"$($TMS_CloudPoradna_TagName_Azure)`" found: $($Tag.id.Substring(0, 8)).....$($Tag.id.Substring($Tag.id.Length - 8, 8))"
        $UriResource = "teams/$($GroupId_TMS_CloudPoradna)/tags/$($Tag.id)/members"
        $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
        [array]$ASISTMS_CloudPoradnaTagMembers = (Get-GraphOutputREST -Uri $Uri -ContentType $ContentTypeJSON -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken).userId
        Write-Log "Teams tag `"$($TMS_CloudPoradna_TagName_Azure)`" members: $($ASISTMS_CloudPoradnaTagMembers.count)"
        #assign tag to all users in ASISAzureSubOwnersStdList who are in TMS_CloudPoradna but have not tag
        if ($ASISTMS_CloudPoradnaTagMembers) {
            #get members of TMS_CloudPoradna group
            $ASISTMS_CloudPoradna = (Get-GroupMembersFromGraphById -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -id $GroupId_TMS_CloudPoradna).id
            foreach ($userId in $ASISAzureSubOwnersStdList) {
                if (($ASISTMS_CloudPoradna.Contains($userId)) -and (-Not($ASISTMS_CloudPoradnaTagMembers.Contains($userId)))) {
                    $missingTagFound = $true
                    $missingUser = Get-GraphUserById -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -id $userId
                    $GraphBody = [PSCustomObject]@{
                        UserId = $userId;
                    } | ConvertTo-Json
                    Try {
                        Write-Log "Adding tag to $($userId) $($missingUser.userPrincipalName)"
                        Invoke-WebRequest -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Body $GraphBody -Method "POST" -ContentType $ContentTypeJSON
                    }
                    Catch {
                        Write-Log $_.Exception.Message -MessageType "Error"
                        Write-Log "Error adding $($userId) $($missingUser.userPrincipalName)" -MessageType "Error"
                    }
                }
            }
            if (-not($missingTagFound)) {
                Write-Log "No missing tag members found"
            }
        }
    }
}

#######################################################################################################################

. $IncFile_StdLogEndBlock
