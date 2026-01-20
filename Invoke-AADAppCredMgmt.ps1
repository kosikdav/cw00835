#######################################################################################################################
# Invoke-AADAppCredMgmt
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################


$LogFolder          = "aad-apps-mgmt"
$LogFilePrefix      = "aad-apps-mgmt"
$LogFileFreq        = "MD"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$AADAppCredMgmt_initialRun = $true

$regexMail = '[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
$firstNotfication 	= 30
$secondNotfication 	= 3
[array]$NotificationThresholds = (30,3)
$credentialDeleteThreshold = 31
$DB_changed = $false
function Get-CredentialValidDaysLeft {
	[CmdletBinding()]
    param (
		[Parameter(Mandatory)]$Credential
    )
	# main function body ##################################
	$now = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
	Return (New-TimeSpan -Start $now -End $Credential.endDateTime).Days
}

function Get-CredentialIsValid {
	[CmdletBinding()]
    param (
		[Parameter(Mandatory)]$Credential
    )
	# main function body ##################################
	$now = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
	if ($Credential.endDateTime -gt $now) {
		Return $true
	} else {
		Return $false
	}
}

function Get-NewCredDBObject {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)]$Credential,
		[string]$ApplicationId,
		[string]$ApplicationName
	)
	# main function body ##################################
	$dbObject = [PSCustomObject]@{
		keyId 				= $Credential.keyId;
		applicationId		= $Application.id;
		applicationName		= $Application.displayName;
		startDateTime 		= $Credential.startDateTime;
		endDateTime 		= $Credential.endDateTime;
		lastNotificication 	= $null
	}
	Return $dbObject
}
function Initialize-DB {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)]$Applications
	)
	# main function body ##################################
	write-host "Apps: $($Applications.count)" -ForegroundColor Yellow
	[hashTable]$DB = @{}
	foreach ($Application in $Applications) {
		if ($Application.passwordCredentials) {
			foreach ($Credential in $Application.passwordCredentials) {
				$dbObject = Get-NewCredDBObject -Credential $Credential -ApplicationId $Application.id -ApplicationName $Application.displayName
				$DB.Add($Credential.keyId, $dbObject) 
			}
		}
		if ($Application.keyCredentials) {
			foreach ($Credential in $Application.keyCredentials) {
				$dbObject = Get-NewCredDBObject -Credential $Credential -ApplicationId $Application.id -ApplicationName $Application.displayName
				try {
					$DB.Add($Credential.keyId, $dbObject) 
				}
				Catch {
					Write-Log "Error adding $($Credential.keyId) to DB $($Application.DisplayName)" -MessageType "Error"}
				
			}
		}
	}
	Write-Host "DB initialized with $($DB.count) records" -ForegroundColor Yellow
	return $DB
}

##################################################################################################

. $IncFile_StdLogBeginBlock

##########################################################################################
Write-Log $string_Divider
Write-Log "Initial run: $($AADAppCredMgmt_initialRun)"
Write-Log "Credential delete threshold: $($credentialDeleteThreshold)"
Write-Log "First notification: $($firstNotfication)"
Write-Log "Second notification: $($secondNotfication)"
Write-Log "DB file: $($DBFileAADCreds)"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "applications"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
$AADApplications = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -Text "AAD applications" -ProgressDots

#Initial run - fill database with existing users so they appeared as added/processed
if ($AADAppCredMgmt_initialRun) {
    Write-Log "Initial run - fill database with cred data"
	$AADCreds_DB = Initialize-DB -Applications $AADApplications
	Try {
        $AADCreds_DB | Export-Clixml -Path $DBFileAADCreds
        Write-Log "DB file $($DBFileAADCreds) exported successfully, $($AADCreds_DB.count) records saved"
    }
    Catch {
        Write-Log "Error exporting $($DBFileAADCreds)" -MessageType "Error"
    }
}
#load DB from disk
else {
	if (test-path $DBFileAADCreds) {
		Try {
			$AADCreds_DB = Import-Clixml -Path $DBFileAADCreds
			Write-Log "DB file $($DBFileAADCreds) imported successfully, $($AADCreds_DB.count) records found"
		} 
		Catch {
			Write-Log "Error importing $($DBFileAADCreds), creating empty DB" -MessageType "Error"
			[hashtable]$AADCreds_DB = @{}
		}
	}
	else {
		Write-Log "DB file $($DBFileAADCreds) not found, creating empty DB" -MessageType "Error"
		[hashtable]$AADCreds_DB = @{}
	}	
}

foreach ($Application in $AADApplications) {
	write-host $string_Divider -ForegroundColor Yellow
	write-host "$($Application.displayName)" -ForegroundColor Cyan
	$notesExtractedMails = $null
	if ($Application.notes) {
		$notesExtractedMails = [regex]::Matches($Application.notes, $regexMail) -join ";"
	}
	
	if ($Application.passwordCredentials) {
		write-host "passwordCredentials: $($Application.passwordCredentials.count)" -ForegroundColor Cyan
		$NewestValidCred = $NewestValidCredDaysLeft = $null
		$NotificationSent = $false
		foreach ($Credential in $Application.passwordCredentials) {
			$DaysLeft = Get-CredentialValidDaysLeft -Credential $Credential
			if (Get-CredentialIsValid -Credential $Credential) {
				if ($NewestValidCredDaysLeft -lt $DaysLeft) {
					$NewestValidCred = $Credential
					$NewestValidCredDaysLeft = $DaysLeft
				}
			}
			if ($DaysLeft -lt -$CredentialDeleteThreshold) {
				write-host "deleting password credential: $($Credential.keyId) daysLeft: $($daysLeft)" -ForegroundColor Red
			}
		}
		
		if ($NewestValidCred) {
			write-host "newest valid credential: $($NewestValidCred.keyId) daysLeft: $($NewestValidCredDaysLeft)" -ForegroundColor Cyan
			Write-Host $NewestValidCred -ForegroundColor DarkGray
			$CredDBRecord = $null
			$DaysSinceLastNotification = 9999
			if ($AADCreds_DB.Contains($NewestValidCred.keyId)) {
				write-host "found"
				$CredDBRecord = $AADCreds_DB[$NewestValidCred.keyId]
				if ($CredDBRecord.lastNotificication) {
					$DaysSinceLastNotification = (Get-Date).Subtract($CredDBRecord.lastNotificication).Days
				}
				#if AAD app was renamed, update name in DB
				if ($CredDBRecord.ApplicationName -ne $Application.displayName) {
					$CredDBRecord.ApplicationName = $Application.displayName
					$AADCreds_DB[$NewestValidCred.keyId] = $CredDBRecord
					$DB_changed = $true
				}
			}
			else {
				write-host "not found"
				#create new record in DB
				$CredDBRecord = Get-NewCredDBObject -Credential $NewestValidCred -ApplicationId $Application.id -ApplicationName $Application.displayName
				$AADCreds_DB.Add[$NewestCred.keyId, $CredDBRecord]
				$DB_changed = $true
			}
			write-host "NewestValidCredDaysLeft:$NewestValidCredDaysLeft DaysSinceLastNotification:$DaysSinceLastNotification" -ForegroundColor DarkGray
			#write-host $AADCreds_DB[$NewestCred.keyId] -ForegroundColor DarkGray
			for ($i=0; $i -lt $NotificationThresholds.count; $i++) {

				write-host "$i NotificationThreshold: $($NotificationThresholds[$i])" -ForegroundColor Magenta
				if (($NewestValidCredDaysLeft -le $NotificationThresholds[$i]) -and ($DaysSinceLastNotification -gt $NotificationThresholds[$i])) {
					Try {
						write-host "$($Application.displayName) credential: $($NewestValidCred.displayName) daysLeft: $($NewestValidCredDaysLeft) sending notfication" -ForegroundColor Yellow
						$NotificationSent = $true
						$CredDBRecord.lastNotificication = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
						$CredDBRecord.ApplicationName = $Application.displayName
						$AADCreds_DB[$NewestValidCred.keyId] = $CredDBRecord
						$DB_changed = $true
						break
					}
					Catch {
						Write-Host $_.Exception.Message
						Write-Log "Error sending notification email" -MessageType "Error"
					}
				}
			}

		}#foreach passwordCredential
	}#if app has passwordCredentials
}#foreach application

#saving DB XML if needed
if (($AADCreds_DB.count -gt 0) -and ($DB_changed)){
    Try {
        $AADCreds_DB | Export-Clixml -Path $DBFileAADCreds
        Write-Log "DB file $($DBFileAADCreds) exported successfully, $($AADCreds_DB.count) records saved"
    }
    Catch {
        Write-Log "Error exporting $($DBFileAADCreds)" -MessageType "Error"
    }
}

. $IncFile_StdLogEndBlock
