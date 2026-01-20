#######################################################################################################################
# Update-AD-Groups.ps1
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "ad-group-mgmt"
$LogFilePrefix		= "ad-group-mgmt"

#######################################################################################################################

$QLPrefixes = @("QLZR","QLMT","QLEX","QLTB","QLZP")

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

if ($InteractiveRun) {
	$ADCredentialPath = "c:\cred\qp_aad_grp_mgmt\qp_aad_grp_mgmt_qskosikdav.cred"
}
else {
	$ADCredentialPath = $aad_grp_mgmt_cred
}

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

# get array of AD group members array, contains only UPNs
function Get-ADGroupMembersUPN {
	param(
		[string]$GroupName,
		[pscredential]$Credential
	)
	[array]$MembersUPN = @()
	Try {
		$ADGroupMembers = Get-ADGroupMember -Credential $Credential -Identity $GroupName -ErrorAction Stop
		foreach ($Member in $ADGroupMembers) {
			$User = Get-ADUser -Credential $Credential $Member.samAccountName -ErrorAction Stop
			$MembersUPN += $User.userPrincipalName
		}
	}
	Catch {
		Write-Log "Critical AD error: $($_.Exception.Message), exiting" -MessageType Error
		Exit
	}
	return $MembersUPN
}

# fill target group with user objects from specific OU
function Update-ADGroupMembersByOU {
	param(
		[string]$TargetGroupName,
		[array]$ASISUsers,
		[array]$TOBEUsers,
		[pscredential]$Credential
	)
	
	foreach ($user in $TOBEUsers) {
		if (-not($ASISUsers.samAccountName -contains $user.samAccountName)) {
			Write-Log "$($user.samAccountName) - adding to $($TargetGroupName)"
			Try {
				Add-ADGroupMember -Identity $TargetGroupName -Members $user.samAccountName -Credential $Credential
			}
			Catch {
				Write-Log "Error adding $($user.samAccountName) to $($TargetGroupName): $($_.Exception.Message)" -MessageType "ERROR"
			}
		}
	}
	
	foreach ($user in $ASISUsers) {
		if (-not($TOBEUsers.samAccountName -contains $user.samAccountName)) {
			Write-Log "$($user.samAccountName) - removing from $($TargetGroupName)"
			Try {
				Remove-ADGroupMember -Identity $TargetGroupName -Members $user.samAccountName -Confirm:$false -Credential $Credential
			}
			Catch {
				Write-Log "Error removing $($user.samAccountName) from $($TargetGroupName): $($_.Exception.Message)" -MessageType "ERROR"
			}
		}
	}
}

# get array of users assigned to specific type of Entra PIM role schedule, contains only UPNs
function Get-RoleScheduleMembersUPN {
	param(
		[string]$Schedule,
		[string]$AccessToken,
		[array]$IgnoredRoles
	)
	[array]$ScheduleMembers = @()
	$Headers = @{Authorization = "Bearer $accessToken"}
	$UriResource = "roleManagement/directory/$($Schedule)"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
	[array]$RoleSchedules = Get-GraphOutputREST -Uri $Uri -AccessToken $AccessToken -ContentType $ContentTypeJSON
	if ($GraphError) {
		Write-Log "Graph API error, exiting" -MessageType "ERROR"
		Exit
	}
	foreach ($RoleSchedule in $RoleSchedules) {
		if ($IgnoredRoles -Contains ($RoleSchedule.RoleDefinitionId)) {
			continue
		}
		#if ($RoleSchedule.PrincipalId -eq "4d207f54-afa3-4162-a5da-9adbb90b2fa1") {
		#		write-host $RoleSchedule -foregroundcolor Red
		#}
		$Principal = $User = $null
		$UriResource = "directoryObjects/$($RoleSchedule.PrincipalId)"
		$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
		try {
			$Principal = Invoke-RestMethod -Headers $Headers -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON
		}
		catch {
			$Principal = $null
		}
		if ($Principal."@odata.type" -eq $TypeUser) {
			$UriResource = "users/$($Principal.Id)"
			$UriSelect = "id,userPrincipalName,accountEnabled,companyName,department,jobTitle,userType"
			$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
			try {
				$User = Invoke-RestMethod -Headers $Header -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON
				if ((-not $User) -or (-not $user.accountEnabled) -or ($User.userType -ne "Member")) {
					continue
				}
				else {
					$ScheduleMembers += $User.userPrincipalName
				}
			}
			catch {
				$User = $null
			}
		} 
		else {
			if ($Principal."@odata.type" -eq $TypeGroup) {
				$membersUPN = (Get-GroupMembersFromGraphById -id $Principal.Id -AccessToken $AccessToken -Transitive:$true -ExcludeGuests).userPrincipalName 

				$ScheduleMembers += $membersUPN
			}
		}
	}
	$ScheduleMembers = $ScheduleMembers | Sort-Object -unique
	return $ScheduleMembers
}

#######################################################################################################################
#######################################################################################################################
#######################################################################################################################

. $IncFile_StdLogStartBlock
Write-Log "AD credential file: $($ADCredentialPath)"
Write-Log $String_Divider

$ADCredential = Import-Clixml -Path $ADCredentialPath

##########################################################################################
#CEZ_AZURE_SUBSCRIPTION_OWNERS ###########################################################
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30

Write-Log $string_Divider
Write-Log "Processing CEZ_Azure_Subscription_Owners_Std group"
$NamePattern = "CEZ_Azure_*_SA"
$GroupFilter = "Name -like '$($NamePattern)'"

write-log "Processing group $($GroupName_AzureSubOwnersQS) membership"
write-log "Processing group $($GroupName_AzureSubOwnersStd) membership"

$TOBEAzureSubOwnersQS = @()
$TOBEAzureSubOwnersStd = @()

if ($ADCredential) {
    $ADGroups = Get-ADGroup -Credential $ADCredential -Filter $GroupFilter -SearchBase $OU_AzSub
	$ASISAzureSubOwnersQS = (Get-ADGroupMember -Identity $GroupName_AzureSubOwnersQS -Credential $ADCredential).samaccountname
    $ASISAzureSubOwnersStd = (Get-ADGroupMember -Identity $GroupName_AzureSubOwnersStd -Credential $ADCredential).samaccountname
}
else {
    $ADGroups = Get-ADGroup -Filter $GroupFilter -SearchBase $OU_AzSub
	$ASISAzureSubOwnersQS = (Get-ADGroupMember -Identity $GroupName_AzureSubOwnersQS).samaccountname
    $ASISAzureSubOwnersStd = (Get-ADGroupMember -Identity $GroupName_AzureSubOwnersStd).samaccountname
}
write-log "$($ADGroups.count) AD groups found in $($OU_AzSub) matching pattern $($NamePattern)"
write-log "$($GroupName_AzureSubOwnersQS) current members: $($ASISAzureSubOwnersQS.count)"
write-log "$($GroupName_AzureSubOwnersStd) current members: $($ASISAzureSubOwnersStd.count)"

foreach ($group in $ADGroups) {
	write-host "Processing group $($group.name) " -ForegroundColor Yellow -NoNewline
    if ($ADCredential) {
        $members  = Get-ADGroupMember -Identity $group.samaccountname -Credential $ADCredential | ForEach-Object { Get-ADUser $_.samaccountname -Credential $ADCredential -Properties "extensionAttribute3" | Select-Object userPrincipalName,samAccountName,extensionAttribute3 }
    }
    else {
        $members  = Get-ADGroupMember -Identity $group.samaccountname | ForEach-Object { Get-ADUser $_.samaccountname -Properties "extensionAttribute3" | Select-Object userPrincipalName,samAccountName,extensionAttribute3 }
    }
	write-host $members.count -ForegroundColor Green
	$TOBEAzureSubOwnersQS += $members.samAccountName
    $TOBEAzureSubOwnersStd += $members.extensionAttribute3
}

$TOBEAzureSubOwnersQS = $TOBEAzureSubOwnersQS | Sort-Object -Unique
$TOBEAzureSubOwnersStd = $TOBEAzureSubOwnersStd | Sort-Object -Unique

$ASISAzureSubOwnersQS = $ASISAzureSubOwnersQS | Sort-Object -Unique
$ASISAzureSubOwnersStd = $ASISAzureSubOwnersStd | Sort-Object -Unique

write-host "TOBE members QS: $($TOBEAzureSubOwnersQS.count)" -ForegroundColor Cyan
write-host "ASIS members QS: $($ASISAzureSubOwnersQS.count)" -ForegroundColor Cyan

write-host "TOBE members Std: $($TOBEAzureSubOwnersStd.count)" -ForegroundColor Yellow
write-host "ASIS members Std: $($ASISAzureSubOwnersStd.count)" -ForegroundColor Yellow

Try {
	$DifferenceQS = Compare-Object -ReferenceObject $ASISAzureSubOwnersQS -DifferenceObject $TOBEAzureSubOwnersQS
}
Catch {
	$DifferenceQS = $true
}

Try {
	$DifferenceStd = Compare-Object -ReferenceObject $ASISAzureSubOwnersStd -DifferenceObject $TOBEAzureSubOwnersStd
}
Catch {
	$DifferenceStd = $true
}

if ($DifferenceQS) {
    $missingMembersQS = $TOBEAzureSubOwnersQS | Where-Object { $ASISAzureSubOwnersQS -notcontains $_ }
	write-log "Missing members QS: $($missingMembersQS.count)"
    $extraMembersQS = $ASISAzureSubOwnersQS | Where-Object { $TOBEAzureSubOwnersQS -notcontains $_ }
	write-log "Extra members QS: $($extraMembersQS.count)"
	
	if ($missingMembersQS) {
        foreach ($samAccountName in $missingMembersQS) {
            Write-Log "Adding $($samAccountName) to $($GroupName_AzureSubOwnersQS)"
            if ($ADCredential) {
                Add-ADGroupMember -Identity $GroupName_AzureSubOwnersQS -Members $samAccountName -Credential $ADCredential
            }
            else {
                Add-ADGroupMember -Identity $GroupName_AzureSubOwnersQS -Members $samAccountName
            }
        }
    }
    if ($extraMembersQS) {
        foreach ($samAccountName in $extraMembersQS) {
            Write-Log "Removing $($samAccountName) from $($GroupName_AzureSubOwnersQS)"
            if ($ADCredential) {
                Remove-ADGroupMember -Identity $GroupName_AzureSubOwnersQS -Members $samAccountName -Credential $ADCredential -Confirm:$false
            }
            else {
                Remove-ADGroupMember -Identity $GroupName_AzureSubOwnersQS -Members $samAccountName -Confirm:$false
            }
        }
    }
}

if ($DifferenceStd) {
	$missingMembersStd = $TOBEAzureSubOwnersStd | Where-Object { $ASISAzureSubOwnersStd -notcontains $_ }
	write-log "Missing members Std: $($missingMembersStd.count)"
	$extraMembersStd = $ASISAzureSubOwnersStd | Where-Object { $TOBEAzureSubOwnersStd -notcontains $_ }
	write-log "Extra members Std: $($extraMembersStd.count)"
	
	if ($missingMembersStd) {
		foreach ($samAccountName in $missingMembersStd) {
			Write-Log "Adding $($samAccountName) to $($GroupName_AzureSubOwnersStd)"
			if ($ADCredential) {
				Add-ADGroupMember -Identity $GroupName_AzureSubOwnersStd -Members $samAccountName -Credential $ADCredential
			}
			else {
				Add-ADGroupMember -Identity $GroupName_AzureSubOwnersStd -Members $samAccountName
			}
		}
	}
	if ($extraMembersStd) {
		foreach ($samAccountName in $extraMembersStd) {
			Write-Log "Removing $($samAccountName) from $($GroupName_AzureSubOwnersStd)"
			if ($ADCredential) {
				Remove-ADGroupMember -Identity $GroupName_AzureSubOwnersStd -Members $samAccountName -Credential $ADCredential -Confirm:$false
			}
			else {
				Remove-ADGroupMember -Identity $GroupName_AzureSubOwnersStd -Members $samAccountName -Confirm:$false
			}
		}
	}
}	

#######################################################################################################################
# CEZ_Lic_AAD_Prem_P2 #################################################################################################
#######################################################################################################################

[array]$ADGroupSAM = @()

[array]$ASISAADP2Users = @()
[array]$ASISE5SecUsers = @()
[array]$RASUsers = @()
[array]$RESUsers = @()

Write-Log $String_Divider
Write-Log "Processing AD license group: $($AADP2LicenseGroup)"

# ASIS AADP2 onprem users ########################################################################
$ASISAADP2Users = Get-ADGroupMembersUPN -GroupName $AADP2LicenseGroup -Credential $ADCredential
Write-Log "Current $($AADP2LicenseGroup) members: $($ASISAADP2Users.count)"

# ASIS E5Sec onprem users ########################################################################
$ASISE5SecUsers = Get-ADGroupMembersUPN -GroupName $E5SecLicenseGroup -Credential $ADCredential
Write-Log "Current $($E5SecLicenseGroup) members: $($ASISAADP2Users.count)"

# RAS ############################################################################################
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$RASUsers = Get-RoleScheduleMembersUPN -Schedule "roleAssignmentSchedules" -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -IgnoredRoles $NoPIMAdminRoles
Write-Log "RASUsers: $($RASUsers.count)"

# RES ############################################################################################
Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$RESUsers = Get-RoleScheduleMembersUPN -Schedule "roleEligibilitySchedules" -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -IgnoredRoles $NoPIMAdminRoles
Write-Log "RESUsers: $($RESUsers.count)"

# TOBE ###########################################################################################
$TOBEAADP2Users = $RASUsers + $RESUsers | Sort-Object -Unique
$TOBEAADP2Users = $TOBEAADP2Users | Where-Object {$ASISE5SecUsers -notcontains $_}
$TOBEAADP2Users = $TOBEAADP2Users | Where-Object {$NoPIMAccountPrefixes -notcontains $_.Substring(0,2)}
$TOBEAADP2Users = $TOBEAADP2Users | Where-Object {$NoPIMUsers -notcontains $_}

Write-Log "TOBE admin group members: $($TOBEAADP2Users.count)"

$missingUsersOnprem = $TOBEAADP2Users | Where-Object -FilterScript { $_ -notin $ASISAADP2Users }
$extraUsersOnprem = $ASISAADP2Users | Where-Object -FilterScript { $_ -notin $TOBEAADP2Users }

Write-Log "Missing users: $($missingUsersOnprem.count)"
foreach ($missingUpn in $missingUsersOnprem) {
	Write-Log "Adding $($missingUpn)"
	$ADUser = Get-ADUser -Credential $ADCredential -Filter "UserPrincipalName -eq '$missingUpn'"
	Add-ADGroupMember -Credential $ADCredential -Identity $AADP2LicenseGroup -Members $ADUser -Confirm:$false
}

Write-Log "Extra users: $($extraUsersOnprem.count)"
foreach ($extraUpn in $extraUsersOnprem) {
	Write-Log "Removing $($extraUpn)"
	$ADUser = Get-ADUser -Credential $ADCredential -Filter "UserPrincipalName -eq '$extraUpn'"
	Remove-ADGroupMember -Credential $ADCredential -Identity $AADP2LicenseGroup -Members $ADUser -Confirm:$false
}

Write-Log $String_Divider
Write-Log "Processing OU based groups"

#######################################################################################################################
# OU Resources ########################################################################################################
#######################################################################################################################
Write-Log "OU Resources"
[array]$TOBEOUResourcesUsers = Get-ADUser -Credential $ADCredential -SearchBase $OUResourcesDN -SearchScope Subtree -Filter * -Properties *
[array]$ASISOUResourcesUsers = Get-ADGroupMember -Credential $ADCredential -Identity $OUResourcesADGroup -ErrorAction Stop

Update-ADGroupMembersByOU -TargetGroupName $OUResourcesADGroup -ASISUsers $ASISOUResourcesUsers -TOBEUsers $TOBEOUResourcesUsers -Credential $ADCredential

#######################################################################################################################
# OU SyncToAzure ######################################################################################################
#######################################################################################################################
Write-Log "OU SyncToAzure"
[array]$TOBEOUSyncToAzureUsers = Get-ADUser -Credential $ADCredential -SearchBase $OUSyncToAzureDN -SearchScope Subtree -Filter * -Properties *
[array]$ASISOUSyncToAzureUsers = Get-ADGroupMember -Credential $ADCredential -Identity $OUSyncToAzureADGroup -ErrorAction Stop

Update-ADGroupMembersByOU -TargetGroupName $OUSyncToAzureADGroup -ASISUsers $ASISOUSyncToAzureUsers -TOBEUsers $TOBEOUSyncToAzureUsers -Credential $ADCredential

#######################################################################################################################
# CEZ_Lic_M365_BasSrd_F3 ##############################################################################################
#######################################################################################################################

[array]$ASISF3SharedMembers = @()
$counter = 0
Write-Log $String_Divider
Write-Log "Processing AD license group: $($M365F3SharedLicenseGroup)"

Try {
	$ADGroupSAM = Get-ADGroupMember -Credential $ADCredential -Identity $M365F3SharedLicenseGroup -ErrorAction Stop
	foreach ($user in $ADGroupSAM) {
		$AsisUser = Get-ADUser -Credential $ADCredential $user.samAccountName -Properties "userAccountControl" -ErrorAction Stop
		$ASISF3SharedMembers += $AsisUser
	}
}
Catch {
	Write-Log "Critical AD error: $($_.Exception.Message), exiting" -MessageType Error
	Exit
}

Write-Log "Current $($M365F3SharedLicenseGroup) members: $($ADGroupSAM.count)"

foreach ($User in $ASISF3SharedMembers) {
	if ($user.samAccountName.startswith("QL")) {
		if (($User.userAccountControl -bAnd 2) -and ($QLPrefixes -contains ($user.SamAccountName.Substring(0,4)))) {
			Write-Log "$($user.userPrincipalName) is disabled, removing from group $($M365F3SharedLicenseGroup)"
			Remove-ADGroupMember -Credential $ADCredential -Identity $M365F3SharedLicenseGroup -Members $User -Confirm:$false
			$counter++
		}
	}
}

Write-Log "Removed $($counter) disabled QL user accounts from $($M365F3SharedLicenseGroup)"


#######################################################################################################################

. $IncFile_StdLogEndBlock
