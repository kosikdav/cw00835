#######################################################################################################################
#######################################################################################################################
# INCLUDE-FUNCTIONS-COMMON
#######################################################################################################################
#######################################################################################################################
#
#
#
########################################################################################
# Get-DateTimeStamp
########################################################################################
function Get-DateTimeStamp {
	param (
		[string][ValidateSet("Y","YM","YMD","YMDH","YMDHM","YMDHMS")]$DateFormat = "YMD",
		[int]$DayOffset = 0
	)
	# main function body ##################################
	$Date = (Get-Date).AddDays($DayOffset)
	Switch ($DateFormat) {
		'Y' 		{$Timestamp = $Date.ToString("yyyy")}
		'YM' 		{$Timestamp = $Date.ToString("yyyy-MM")}
		'YMD' 		{$Timestamp = $Date.ToString("yyyy-MM-dd")}
		'YMDH' 		{$Timestamp = $Date.ToString("yyy-MM-dd-HH")}
		'YMDHM' 	{$Timestamp = $Date.ToString("yyyy-MM-dd-HHmm")}
		'YMDHMS' 	{$Timestamp = $Date.ToString("yyyy-MM-dd-HHmmss")}
	}
	return $Timestamp
}
########################################################################################
# Test-IsGuid
########################################################################################
function Test-IsGuid{
	[OutputType([bool])]    
	param (        
		[Parameter(Mandatory = $true)][string]$StringGuid
	)    
	$ObjectGuid = [System.Guid]::empty   
	return [System.Guid]::TryParse($StringGuid,[System.Management.Automation.PSReference]$ObjectGuid)
}

########################################################################################
# Get-Count
########################################################################################
function Get-Count {
	param (
		[Parameter(Position=0)]$Object
	)
	# main function body ##################################
	If ($Object.Count -ge 1) {
		Return $Object.Count
	}
	Else {
		Return 0
	}
}

function Get-YesNoKeyboardInput {
    param (
        [Parameter(Mandatory=$true)][string]$Prompt,
		[string]$ForegroundColor
    )
    if (-not($ForegroundColor)) {
		$ForegroundColor = "Gray"
	}
	Write-Host "$($Prompt) [Y/N]" -ForegroundColor $ForegroundColor
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


########################################################################################
# New-OutputFile
########################################################################################
function New-OutputFile {
	param (
		[string]$RootFolder,
		[string]$Folder,
		[parameter(Mandatory = $true)][string]$Prefix,
		[string]$Suffix,
		[string][ValidateSet("Y","YM","YMD","YMDH","YMDHM")]$Freq,
		[switch]$NoDate,
		[string]$SpecificDate,
		[parameter(Mandatory = $true)][string][ValidateSet("LOG","CSV")]$Ext,
		[switch]$FileDateYesterday,
		[int]$FileDateDayOffset = 0
	)
	# main function body ##################################
	if ($FileDateYesterday) {
		$FileDateDayOffset = -1
	}
	
	if (-not($RootFolder)) {
		switch ($Ext) {
			"LOG" {$RootFolder = $script:root_log_folder}
			"CSV" {$RootFolder = $script:root_output_folder}
		}
	}
	
	if (-not($Freq)) {
		switch ($Ext) {
			"LOG" {$Freq = "Y"}
			"CSV" {$Freq = "YMD"}
		}
	}
	
	If ($NoDate) {
		$FileName = $Prefix.Trim("-")
	}
	Else {
		if ($SpecificDate) {
			$FileName = $Prefix.Trim("-") + "-" + "$($SpecificDate)"
		}
		else {
			$FileName = $Prefix.Trim("-") + "-" + "$(Get-DateTimeStamp -DateFormat $Freq -DayOffset $FileDateDayOffset)"
		}
		
	}
	
	if ($Suffix) {
		$FileName = $FileName + "-" + $Suffix.Trim("-") + "." +$ext.ToLower()
	}
	else {
		$FileName = $FileName + "." + $ext.ToLower()
	}
	$File = [System.IO.Path]::Combine($RootFolder,$Folder,$FileName)
	return $File
}

########################################################################################
# Write-Log
########################################################################################
function Write-Log {
	[CmdletBinding(SupportsShouldProcess = $true)]
	param (
		[Parameter(Position=0)][string]$String,
		[string][ValidateSet("Info","Warning","Warn","Error","Err")]$MessageType = "Info",
		[string]$ForegroundColor,
		[string]$BackgroundColor = "Black",
		[switch]$NoNewLine,
		[switch]$NoLinePrefix,
		[switch]$ForceOnScreen,
		[switch]$ForceOffScreen,
		[string]$LogString = $null,
		[string]$AlternateLogfile = $null
	)
	# main function body ##################################
	if ($MessageType -eq "Warn") {
		$MessageType = "Warning"
	}
	if ($MessageType -eq "Err") {
		$MessageType = "Error"
	}
	if ($AlternateLogfile) {
		$File = $AlternateLogfile
	}
	else {
		$File = $script:LogFile
	}
	If ($LogString) {
		$String = $LogString
	}
	if ($File) {
		$Folder = Split-Path -Parent $File
		If (-not(Test-Path -Path $Folder)) {
			Try {
				New-Item -ItemType "directory" -Path $Folder
			}
			Catch {
				Write-Host "Unable to create log folder $($Folder)"
			}
		}
		$TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
		if ($NoLinePrefix) {
			$LinePrefix = ""
		}
		Else {
			switch ($MessageType) {
				"Info"		{$LineType = "INFO"}
				"Warning" 	{$LineType = "WARN"}
				"Error" 	{$LineType = "ERR"}
				Default 	{$LineType = "INFO"}
			}
			$LinePrefix = $TimeStamp + " [" + ($LineType.PadRight(4," ")).ToUpper() + "] "
		}
		if ($NoNewLine) {
			Add-Content $File -Value ($LinePrefix + $String) -NoNewline
		}
		Else {
			Add-Content $File -Value ($LinePrefix + $String)
		}
	}
	If (-not($ForceOffScreen) -and $interactiveRun) {
		If (($EnableOnScreenLogging -or $ForceOnScreen)) {
			if  (-not($ForegroundColor)) {
				switch ($MessageType) {
					"Info"		{$ForegroundColor = "Gray"}
					"Warning" 	{$ForegroundColor = "DarkYellow"}
					"Error" 	{$ForegroundColor = "Red"}
					Default {$ForegroundColor = "Gray"}
				}
			}
			if ($NoNewLine) {
				Write-Host $String -ForegroundColor $ForegroundColor -BackgroundColor $BackgroundColor -NoNewline
			}
			else {
				Write-Host $String -ForegroundColor $ForegroundColor -BackgroundColor $BackgroundColor
			}
		}
	}
}

########################################################################################
# New-Folder
########################################################################################
function New-Folder {
	param (
		[string]$RootFolder,
		[string]$Folder,
		[string]$Path,
		[switch]$StopOnError
	)
	# main function body ##################################
	if (-not $Path) {
		If ($RootFolder -and $Folder) {
			$Path = [System.IO.Path]::Combine($RootFolder,$Folder)
		}	
	}
	if (-not(Test-Path -Path $Path)) {
		Try {
			New-Item -ItemType "directory" -Path $Path
		}
		Catch {
			Write-Log "Error creating temp folder $($Path)" -MessageType "ERROR"
			If ($StopOnError) {
				Exit
			}
		}
	}
}



########################################################################################
# Export-Report
########################################################################################
function Export-Report {
	[CmdletBinding()]
    param (
		[Parameter(Position=0)][string]$Text,    
		[Parameter(Mandatory)]$Report,
		[Parameter(Mandatory)][string]$Path,
		[string]$SortProperty,
		[string]$Delimiter = ",",
		[string]$Encoding = "UTF8",
		[boolean]$Append = $False
   )
	# main function body ##################################
	if ($Report) {
		$Folder = Split-Path -Parent $Path
		New-Folder -Path $Folder
		Write-Log "Exporting $($Text) ($($Report.Count) records) to: $($Path)"
		Try {
			if ($SortProperty) {
				$Report | Sort-Object -Property $SortProperty | Export-Csv -Path $Path -Delimiter $Delimiter -Encoding $Encoding -NoTypeInformation -Append:$Append
			}
			else {
				$Report | Export-Csv -Path $Path -Delimiter $Delimiter -Encoding $Encoding -NoTypeInformation -Append:$Append
			}
		}
		Catch {
			Write-Log "Export failed: $($_.Exception.Message)" -MessageType Error
		}
	} else {
		Write-Log "Report $($Text) empty, nothing to export"
	}
}

########################################################################################
# Start-SleepDots
########################################################################################
function Start-SleepDots {
	[CmdletBinding()]
    param (
		[Parameter(Position=0)][string]$String,    
		[Parameter(Mandatory)][int]$Seconds,
		[string]$ForegroundColor = $null,
		[switch]$NoNewLine
    )
	# main function body ##################################
	for ($sec = 1; $sec -le $Seconds; $sec++) {
		Start-Sleep -s 1
		if ($ForegroundColor) {
			Write-Host "." -ForegroundColor $ForegroundColor -NoNewline
		}
		else {
			Write-Host "." -NoNewline
		}
	}
	if (-not($NoNewLine)) {
		write-host
	}
}

########################################################################################
# INITIALIZE-PROGRESSBARMAIN
########################################################################################
function Initialize-ProgressBarMain {
	[alias("InitProgressBarMain")]
	param (
		[string]$Activity, 
		[int]$Total
	)
	# main function body ##################################
	if ($script:interactiveRun) {
		$script:ProgressCountMain = 0
		$script:ProgressActivityMain = $Activity
		$script:ProgressTotalMain = $Total
	}
}

########################################################################################
# UPDATE-PROGRESSBARMAIN
########################################################################################
function Update-ProgressBarMain {
	[alias("UpdateProgressBarMain")]
	param ()
	# main function body ##################################
	if ($script:interactiveRun -and ($script:ProgressTotalMain -gt 10)) {
		$script:ProgressCountMain++
		$ProgressPct = [int](($script:ProgressCountMain/$script:ProgressTotalMain)*100)
		#Write-Progress -Activity $script:ProgressActivityMain -Status "$($ProgressPct)% complete" -PercentComplete $ProgressPct
	}
}

########################################################################################
# COMPLETE-PROGRESSBARMAIN
########################################################################################
function Complete-ProgressBarMain {
	[alias("FinishProgressBarMain")]
	param ()
	# main function body ##################################
	if ($script:interactiveRun -and ($script:ProgressTotalMain -gt 10)) {
		Write-Progress -Activity $script:ProgressActivityMain -Completed
	}
}

########################################################################################
# NEW-GRAPHURI
########################################################################################
function New-GraphUri {
	param (
		[parameter(Mandatory = $true)][string][ValidateSet("v1.0","beta")]$Version,
		[parameter(Mandatory = $true)][string]$Resource,
		[int]$Top = $null,
		[string]$Filter = $null,
		[string]$Select = $null,
		[string]$Search = $null,
		[string]$Expand = $null,
		[string]$OrderBy = $null,
		[string][ValidateSet("D7","D30","D90","D180")]$ReportPeriod = $null,
		[string]$EqualsParam = $null,
		[switch]$Count
	)
	# main function body ##################################
	$BaseUri = "https://graph.microsoft.com"
	$Uri = $BaseUri + "/" + $Version + "/" + $Resource

	if ($Top -or $Filter -or $Select -or $Search -or $Expand -or $OrderBy -or $Count) {
		$Uri = $Uri + "?"
	}
	if ($Top) {
		$Uri = $Uri + "`$Top=$($Top)&"
	}
	if ($Filter) {
		if ($Filter -match " ") {
			$Filter = [System.Web.HttpUtility]::UrlEncode($Filter)
		}
		$Uri = $Uri + "`$Filter=$($Filter)&"
	}
	if ($Select) {
		$Select = $Select -replace " ", ""
		$Uri = $Uri + "`$Select=$($Select)&"
	}
	if ($Count) {
		$Uri = $Uri + "`$Count=true&"
	}
	if ($Search) {
		$Uri = $Uri + "`$Search=`"$($Search)`"&"
	}
	if ($Expand) {
		$Uri = $Uri + "`$Expand=$($Expand)&"
	}
	if ($OrderBy) {
		$Uri = $Uri + "`$OrderBy=$($OrderBy)&"
	}
	
	$Uri = $Uri.Trim("&")
	
	if ($ReportPeriod) {
		$Uri = $Uri + "(Period='$($ReportPeriod)')"
	}
	if ($EqualsParam) {
		$Uri = $Uri + "($($EqualsParam))"
	}
	Return $Uri
}

########################################################################################
# NEW-ADOURI
########################################################################################
function New-ADOUri {
	param (
		[parameter(Mandatory = $true)][string][ValidateSet("5.1","6.0","7.0","7.1")]$Version,
		[parameter(Mandatory = $true)][string]$Resource,
		[parameter(Mandatory = $true)][string]$Organization,
		[int]$Top = $null,
		[string]$Filter = $null,
		[string]$Select = $null,
		[string]$Search = $null,
		[string]$Expand = $null,
		[string]$OrderBy = $null,
		[string][ValidateSet("D7","D30","D90","D180")]$ReportPeriod = $null,
		[string]$EqualsParam = $null
	)
	# main function body ##################################
	$BaseUri = "https://dev.azure.com"
	$Uri = $BaseUri + "/" + $Organization + "/_apis/" + $Resource

	if ($Top -or $Filter -or $Select -or $Search -or $Expand -or $OrderBy) {
		$Uri = $Uri + "?"
	}
	if ($Top) {
		$Uri = $Uri + "`$Top=$($Top)&"
	}
	if ($Filter) {
		$Uri = $Uri + "`$Filter=$($Filter)&"
	}
	if ($Select) {
		$Uri = $Uri + "`$Select=$($Select)&"
	}
	if ($Search) {
		$Uri = $Uri + "`$Search=`"$($Search)`"&"
	}
	if ($Expand) {
		$Uri = $Uri + "`$Expand=$($Expand)&"
	}
	if ($OrderBy) {
		$Uri = $Uri + "`$OrderBy=$($OrderBy)&"
	}
	
	$Uri = $Uri.Trim("&")
	
	if ($ReportPeriod) {
		$Uri = $Uri + "(Period='$($ReportPeriod)')"
	}
	if ($EqualsParam) {
		$Uri = $Uri + "($($EqualsParam))"
	}

	$Uri = $Uri + "?api-version=" + $Version

	Return $Uri
}


########################################################################################
# REQUEST-MSALTOKEN
########################################################################################
function Request-MSALToken {
	param (
		[parameter(Mandatory = $true)][string]$AppRegName,
		[int]$TTL = 20,
		[string]$Authority = "login.microsoftonline.com",
		[string]$Scope = "https://graph.microsoft.com/.default",
		[string]$Resource,
		[switch]$Silent,
		[switch]$Force
	)
	# main function body ##################################
	$AuthRecord = $null
	$MinutesElapsed = 0
	$Operation = "requested"
	
	if ($script:AuthDB.ContainsKey($AppRegName)) {
		$AuthRecord = $script:AuthDB[$AppRegName]
		$MinutesElapsed = [math]::abs((New-TimeSpan -End $AuthRecord.CreatedDateTime).Minutes)
		$Operation = "refreshed"
	}
	if ((-Not($AuthRecord)) -or ($MinutesElapsed -ge $TTL) -or ($Force)) {
		[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
		$incFileName = "include-appreg-" + $AppRegName + ".ps1"
		$incFile = [System.IO.Path]::Combine($incFolder,$incFileName)
		. $incFile
		
		$AuthorityURI = "https://$($Authority)/$($tenantId)"
		if ($Authority -eq "login.microsoftonline.com") {
			$tokenEndpoint = "$($AuthorityURI)/oauth2/v2.0/token"
		}
		if ($Authority -eq "login.windows.net") {
			$tokenEndpoint = "$($AuthorityURI)/oauth2/token"
		}
		
		$CertificateBase64Hash = [System.Convert]::ToBase64String($Certificate.GetCertHash())  
		$StartDate = (Get-Date "1970-01-01T00:00:00Z" ).ToUniversalTime()  
		$JWTExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End (Get-Date).ToUniversalTime().AddMinutes(60)).TotalSeconds  
		$JWTExpiration = [math]::Round($JWTExpirationTimeSpan,0)  
		$NotBeforeExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End ((Get-Date).ToUniversalTime())).TotalSeconds  
		$NotBefore = [math]::Round($NotBeforeExpirationTimeSpan,0)  
		$JWTHeader = @{  
			alg = "RS256"  
			typ = "JWT"  
			x5t = $CertificateBase64Hash -replace '\+','-' -replace '/','_' -replace '='  
		}  

		$JWTPayLoad = @{  
			aud = "$($AuthorityURI)/oauth2/token"  
			exp = $JWTExpiration  
			iss = $ClientId  
			jti = [guid]::NewGuid()  
			nbf = $NotBefore  
			sub = $ClientId  
		}  
		
		$JWTHeaderToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTHeader | ConvertTo-Json))  
		$EncodedHeader = [System.Convert]::ToBase64String($JWTHeaderToByte)  
		$JWTPayLoadToByte =  [System.Text.Encoding]::UTF8.GetBytes(($JWTPayload | ConvertTo-Json))  
		$EncodedPayload = [System.Convert]::ToBase64String($JWTPayLoadToByte)  
		$JWT = $EncodedHeader + "." + $EncodedPayload  
		$PrivateKey = ([System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate))  
		$RSAPadding = [Security.Cryptography.RSASignaturePadding]::Pkcs1  
		$HashAlgorithm = [Security.Cryptography.HashAlgorithmName]::SHA256  
		$Signature = [Convert]::ToBase64String($PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($JWT),$HashAlgorithm,$RSAPadding)) -replace '\+','-' -replace '/','_' -replace '='  
		$JWT = $JWT + "." + $Signature
		
		$client_assertion = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($Certificate))

		$body = @{
			client_id = $clientId
			client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
			client_assertion = $JWT
			grant_type = "client_credentials"
		}
		if ($Resource) {
			$body.Add("resource",$Resource)
		} else {
			$body.Add("scope",$scope)
		}
		Try {
			$Token = Invoke-RestMethod -Uri $tokenEndpoint -Method "POST" -Body $body
			$ExpiresOn = (Get-Date).AddSeconds($Token.expires_in)
			if (-not $Silent) {
				Write-Host "MSAL access token for $($TenantShortName) (app $($AppRegName)) $($Operation) - expires: $($ExpiresOn), TTL $($TTL)" -ForegroundColor DarkGray
			}
			
			$AuthRecordNew = [pscustomobject]@{
				CreatedDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss";
				AccessToken = $Token.access_token;
				AuthHeaders = @{Authorization = "Bearer $($Token.access_token)"}
				ExpiresOn = $ExpiresOn
			}
			$script:AuthDB[$AppRegName] = $AuthRecordNew
			$script:AccessToken = $Token.access_token
			$script:AuthHeaders = @{Authorization = "Bearer $($Token.access_token)"}
		}
		Catch {
			Write-Log -String $_.Exception.Message -MessageType Error -ForceOnScreen -ForegroundColor "Red"
		}
	}
}

########################################################################################
# CONNECT-SPOSERVICEPNP
########################################################################################
function Connect-SPOServicePnP {
	param (
		[parameter(Mandatory = $true)]
		[string]$AppRegName,
		[int]$TTL = 60,
		[string]$URL = $script:PnPURL,
		[bool]$Silent = $false
	)
	# main function body ##################################
	$AuthRecord = $null
	$MinutesElapsed = 0
	$Operation = "requested"
	$incFileName = "include-appreg-" + $AppRegName + ".ps1"
	$incFile = [System.IO.Path]::Combine($incFolder,$incFileName)
	if ($script:AuthDB.ContainsKey($AppRegName)) {
		$AuthRecord = $script:AuthDB[$AppRegName]
		$MinutesElapsed = [math]::abs((New-TimeSpan -End $AuthRecord.CreatedDateTime).Minutes)
		$Operation = "refreshed"
	}
	if ((-Not($AuthRecord)) -or ($MinutesElapsed -ge $TTL) -or ($URL -ne $script:PnPURL)) {
		. $incFile
		[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
		Try {
			if (-not $Silent) {
				Write-Host "PnP PowerShell connection to $($script:TenantShortName) ($($URL)) $($Operation)" -ForegroundColor DarkGray
			}
			Connect-PnPOnline -Url $URL -ClientId $ClientId -Tenant $PNPTenant -Thumbprint $Thumbprint
			$AuthRecordNew = [pscustomobject]@{
				CreatedDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
				AccessToken = "n/a"
				AuthHeaders = "n/a"
				ExpiresOn = "n/a"
			}
			$script:AuthDB[$AppRegName] = $AuthRecordNew
		}
		Catch {
			Write-Log $_.Exception.Message -MessageType Error -ForceOnScreen -ForegroundColor "Red"
		}
	}
}

########################################################################################
# CONNECT-EXOSERVICE
########################################################################################
function Connect-EXOService {
	param (
		[string]$AppRegName,
		[int]$TTL,
		[switch]$ForceReconnect,
		[string]$CredentialFile
	)
	# main function body ##################################
	if ($AppRegName -and $CredentialFile) {
		Write-Log "AppRegName and CredentialFile parameters are mutually exclusive for Connect-EXOService function" -MessageType "ERR"
		return
	}
	if (-not($AppRegName) -and -not($CredentialFile)) {
		Write-Log "Either AppRegName or CredentialFile parameter must be provided for Connect-EXOService function" -MessageType "ERR"
		return
	}
	
	$AuthRecord = $null
	$MinutesElapsed = 0
	$Operation = "requested"

	if ($AppRegName) {
		$incFileName = "include-appreg-" + $AppRegName + ".ps1"
		$incFile = [System.IO.Path]::Combine($incFolder,$incFileName)
		if ($script:AuthDB.ContainsKey($AppRegName)) {
			$AuthRecord = $script:AuthDB[$AppRegName]
			$MinutesElapsed = [math]::abs((New-TimeSpan -End $AuthRecord.CreatedDateTime).Minutes)
			$Operation = "refreshed"
		}
		if ((-Not($AuthRecord)) -or ($MinutesElapsed -ge $TTL) -or ($ForceReconnect)) {
			. $incFile
			[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
			Try {
				Write-Host "EXO PowerShell connection to $($TenantShortName) (app): " -NoNewline -ForegroundColor DarkGray
				Connect-ExchangeOnline -CertificateThumbPrint $ThumbPrint -AppID $ClientId -Organization $TenantName -ShowBanner:$false
				Write-Host $Operation -ForegroundColor DarkGray
				$AuthRecordNew = [pscustomobject]@{
					CreatedDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
					AccessToken = "n/a"
					AuthHeaders = "n/a"
					ExpiresOn = "n/a"
				}
				$script:AuthDB[$AppRegName] = $AuthRecordNew
			}
			Catch {
				Write-Log $_.Exception.Message -MessageType "ERR"
			}
		}
	}
if ($CredentialFile) {
	if (-not(Test-Path -Path $CredentialFile)) {
		Write-Log "Credential file $($CredentialFile) not found" -MessageType "ERR"
		return
	}
	$CredFilename = [System.IO.Path]::GetFileNameWithoutExtension($CredentialFile)
	write-host $CredFilename -foregroundcolor cyan
	Try {
		$Credential = Import-Clixml -Path $CredentialFile
	}
	Catch {
		Write-Log $_.Exception.Message -MessageType "ERR"
	}
	if ($script:AuthDB.ContainsKey($CredFilename)) {
		$AuthRecord = $script:AuthDB[$CredFilename]
		$MinutesElapsed = [math]::abs((New-TimeSpan -End $AuthRecord.CreatedDateTime).Minutes)
		$Operation = "refreshed"
	}
	if ((-Not($AuthRecord)) -or ($MinutesElapsed -ge $TTL) -or ($ForceReconnect)) {
		[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
		Try {
			Write-Host "EXO PowerShell connection to $($TenantShortName) (user): " -NoNewline -ForegroundColor DarkGray
			Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
			Write-Host $Operation -ForegroundColor DarkGray
			$AuthRecordNew = [pscustomobject]@{
				CreatedDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
				AccessToken = "n/a"
				AuthHeaders = "n/a"
				ExpiresOn = "n/a"
			}
			$script:AuthDB[$CredFilename] = $AuthRecordNew
		}
		Catch {
			Write-Log $_.Exception.Message -MessageType "ERR"
		}
	}
}

}



########################################################################################
# CONNECT-TEAMS
########################################################################################
function Connect-Teams {
	param (
		[parameter(Mandatory = $true)][string]$AppRegName,
		[int]$TTL,
		[switch]$ForceReconnect
	)
	# main function body ##################################
	$AuthRecord = $null
	$MinutesElapsed = 0
	$Operation = "requested"
	$incFileName = "include-appreg-" + $AppRegName + ".ps1"
	$incFile = [System.IO.Path]::Combine($incFolder,$incFileName)
	if ($script:AuthDB.ContainsKey($AppRegName)) {
		$AuthRecord = $script:AuthDB[$AppRegName]
		$MinutesElapsed = [math]::abs((New-TimeSpan -End $AuthRecord.CreatedDateTime).Minutes)
		$Operation = "refreshed"
	}
	if ((-Not($AuthRecord)) -or ($MinutesElapsed -ge $TTL) -or ($ForceReconnect)) {
		. $incFile
		[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
		Try {
			Write-Host "Teams PowerShell connection to $($TenantShortName) $($Operation)" -ForegroundColor DarkGray
			Connect-MicrosoftTeams -CertificateThumbPrint $ThumbPrint -ApplicationId $ClientId -TenantId $TenantId
			$AuthRecordNew = [pscustomobject]@{
				CreatedDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
				AccessToken = "n/a"
				AuthHeaders = "n/a"
				ExpiresOn = "n/a"
			}
			$script:AuthDB[$AppRegName] = $AuthRecordNew
		}
		Catch {
			Write-Log -String $_.Exception.Message -MessageType Error -ForceOnScreen -ForegroundColor Red
		}
	}
}


########################################################################################
# Connect-MGModule
########################################################################################
function Connect-MGModule {
	param (
		[parameter(Mandatory = $true)][string]$AppRegName,
		[int]$TTL,
		[switch]$ForceReconnect
	)
	# main function body ##################################
	$AuthRecord = $null
	$MinutesElapsed = 0
	$Operation = "requested"
	$incFileName = "include-appreg-" + $AppRegName + ".ps1"
	$incFile = [System.IO.Path]::Combine($incFolder,$incFileName)
	. $incFile
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
	Try {
		Write-Log -String "Mg PowerShell connection to $($TenantShortName) $($Operation)" -ForceOnScreen -ForegroundColor DarkGray
		Connect-MgGraph -ClientID $ClientId -TenantId $TenantId -CertificateThumbprint $Thumbprint
	}
	Catch {
		Write-Log -String $_.Exception.Message -MessageType Error -ForceOnScreen -ForegroundColor "Red"
	}
}

########################################################################################
# Connect-EntraModule
########################################################################################
function Connect-EntraModule {
	param (
		[parameter(Mandatory = $true)][string]$AppRegName,
		[int]$TTL,
		[switch]$ForceReconnect
	)
	# main function body ##################################
	$AuthRecord = $null
	$MinutesElapsed = 0
	$Operation = "requested"
	$incFileName = "include-appreg-" + $AppRegName + ".ps1"
	$incFile = [System.IO.Path]::Combine($incFolder,$incFileName)
	. $incFile
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
	Try {
		Write-Log -String "Entra PowerShell connection to $($TenantShortName) $($Operation)" -ForceOnScreen -ForegroundColor DarkGray
		Connect-Entra -ClientID $ClientId -TenantId $TenantId -CertificateThumbprint $Thumbprint
	}
	Catch {
		Write-Log -String $_.Exception.Message -MessageType Error -ForceOnScreen -ForegroundColor "Red"
	}
	
}

########################################################################################
# Update-GuestAuditRecordDB
########################################################################################
function Update-GuestAuditRecordDB  {
    param(
		[Parameter(Mandatory)][string]$id,
        [Parameter(Mandatory)]$DateTime,
        [Parameter(Mandatory)][hashtable]$hashtableDB
    )
    if ($hashtableDB.ContainsKey($id)) {
        if ([datetime]$DateTime -gt $hashtableDB[$id]) {
            $hashtableDB[$id] = [datetime]$DateTime
        }
    }
    else {
        $hashtableDB.Add($id,[datetime]$DateTime)
    }
}

########################################################################################
# Write-GuestAuditRecordDBToEntra
########################################################################################
function Write-GuestAuditRecordDBToEntra {
    param(
        [string]$AccessToken,
        [hashtable]$hashtableDB,
        [string]$EntraAttribute = "employeeHireDate",
		[string]$LogFile,
		[string]$AuditType
    )
    $Headers = @{Authorization = "Bearer $($AccessToken)"}
    foreach ($Key in $hashtableDB.Keys) {
        if (Test-IsGuid -StringGuid -$Key) {
			$Id = $Key
		}
		else {
			$Id = [System.Web.HttpUtility]::UrlEncode($Key)
		}
        $UTCDateTimeString = $hashtableDB[$Key].ToString("yyyy-MM-ddTHH:mm:ssZ")
        $UriResource = "users/$($Id)"
        $UriSelect = "id,userPrincipalName,userType,$($EntraAttribute)"
        $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
        Try {
            $Guest = Get-GraphOutputREST -Uri $Uri -AccessToken $AccessToken -ContentType $ContentTypeJSON
            $AttributeValue = $Guest.$EntraAttribute
            if ($Guest -and ($Guest.userType -eq "Guest") -and (($null -eq $AttributeValue) -or ([datetime]($AttributeValue) -lt [datetime]$hashtableDB[$Key]))) {
                $GraphBody = @{
                    $EntraAttribute = $UTCDateTimeString
                } | ConvertTo-Json
                #write-host $GraphBody
                Try {
                    $Result = Invoke-RestMethod -Uri $Uri -Headers $Headers -Method "PATCH" -Body $GraphBody -ContentType $ContentTypeJSON
                    write-host "[$($AuditType)] $($Key) - $($EntraAttribute): $($UTCDateTimeString)"
					Write-Log "[$($AuditType)] $($Guest.id) $($Guest.userPrincipalName) - $($EntraAttribute): $($UTCDateTimeString)" -AlternateLogfile $LogFile
                }
                Catch {
                    write-host "[$($AuditType)] $($Key) - $($EntraAttribute): $($UTCDateTimeString)" -ForegroundColor Red
                    write-host $_.Exception.Message -ForegroundColor Red
                }
            }
        }
        Catch {
            write-host $_.Exception.Message -ForegroundColor Red
        }
    }
}


########################################################################################
# GET-BYTESFROMSTRING
########################################################################################
Function GetBytesFromString ($prmByteString) {
    if (($prmByteString) -and ($prmByteString.Length -ge 12)) {  
        $result = $prmByteString.SubString($prmByteString.IndexOf("(")+1)
        $result = $result.Replace(")","")
        $result = $result.Replace(",","")
        $result = $result.Replace("\s","")
        $result = $result.Replace("bytes","")
        return [int64]($result.Trim())
    }
    else {
        return $null
    }
}

########################################################################################
# Import-CSVtoHashDB
########################################################################################
function Import-CSVtoHashDB {
	[alias("ImportCSVtoHashDB")]
	param (
		[string]$Path, 
		[string]$KeyName
	)
	# main function body ##################################
	[hashtable]$result = @{}
	[array]$Headers = @()
	$CSVFilename = (Get-Item $Path).Name
	#$CSVBaseFilename = (Get-Item $Path).Basename
	Write-Host "Importing $($CSVFilename) as hashtable." -NoNewline
	$ImportedCSV = Import-Csv -Path $Path
	$CSVHeaders = ($ImportedCSV[0].psobject.Properties).name
	if ($CSVHeaders -Contains $KeyName) {
		foreach ($Header in $CSVHeaders) {
			$Headers += $Header
		}
		$ImportedCSV = $ImportedCSV | Sort-Object -Property $KeyName -Unique
		$step = [math]::Truncate($ImportedCSV.count / 10)
		$count = 0
		foreach ($row in $ImportedCSV) {
			$count++
			if ($count % $step -eq 0) {
				Write-Host "." -NoNewline
			}
			$rowObject = [PSCustomObject]@{
			}
			foreach ($Header in $Headers) {
				$rowObject | Add-Member -MemberType NoteProperty -Name $Header -Value $row.$Header
			}
			try {
				$result.Add($row.$KeyName,$rowObject)
			}
			catch {
				write-host $row -ForegroundColor Red
				write-host "Critical Error: $($_.Exception.Message)" -ForegroundColor Red
			}
		}
		Write-Host "done ($($result.count))"
	}
	else {
		Write-Host "Key name: $($KeyName) not found, exiting."
		Break
	}

	Remove-Variable ImportedCSV
	Return $result
}

########################################################################################
# Import-CSVtoArray
########################################################################################
function Import-CSVtoArray {
	[alias("ImportCSVtoArray")]
	param (
		[string]$Path
	)
	# main function body ##################################
	[array]$result = @()
	$CSVFilename = (Get-Item $Path).Name
	Write-Host "Importing $($CSVFilename) as array..." -NoNewline
	$result = Import-Csv -Path $Path
	Write-Host "done ($($result.count))"
	Return $result
}

########################################################################################
# Get-GraphOutputREST
########################################################################################
function Get-GraphOutputREST {
	[CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$Uri,
        [string]$AccessToken,
		[string]$AppRegName,
        [string]$ContentType = "application/json",
        [string]$ConsistencyLevel,
		[string]$Text,
		[switch]$ProgressDots,
		[switch]$Silent,
		[switch]$IncludeUnknownEnumMembers,
		[int]$Retries = 10
    )
	# main function body ##################################
	$script:GraphError = $False
	if (-not $AccessToken -and $AppRegName) {
		Request-MSALToken -AppRegName $AppRegName
		$AccessToken = $AuthDB[$AppRegName].AccessToken 
	}
	$DotCounter = 0
	$GraphUri = $Uri
	if ($GraphUri.Contains("`$top=")) {
		$Top = (($GraphUri.Substring($GraphUri.IndexOf("`$top=")+5,3)).Trim('&')).Trim("$")
	}
	$Headers = @{Authorization = "Bearer $AccessToken"}
	if ($ConsistencyLevel -eq "eventual") {
		$Headers["ConsistencyLevel"] = "eventual"
	}
	if ($includeUnknownEnumMembers) {
		$Headers["Prefer"] = "include-unknown-enum-members"
	}
	If ($text -and $interactiveRun -and (-Not $Silent)) {
		if ($Top) {
			Write-Host "Reading $($text) from Graph (Top=$($Top))" -NoNewline
		}
		Else {
			Write-Host "Reading $($text) from Graph" -NoNewline
		}
		If (-not($ProgressDots)) {
			Write-Host "..." -NoNewline
		}
	}
	Try {
		Switch ($ContentType) {
			"application/json" { 
				Do {
					$retryCount = 0
					if ($ProgressDots -and $interactiveRun -and (-Not $Silent)) {
						if ((($DotCounter % 10) -eq 0) -and ($DotCounter -gt 1)){
							Write-Host $DotCounter -NoNewline -ForegroundColor White
						}
						else {
							Write-Host "." -NoNewline -ForegroundColor Gray
						}
						$DotCounter++
					}
					Do {
						$retry = $false
						$ResponseHeaders = $null
						[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
						Try {
							if ((Get-Host).Version.Major -ge 7) {
								$Query = Invoke-RestMethod -Headers $Headers -Uri $GraphUri -Method "GET" -ContentType "application/json" -ResponseHeadersVariable ResponseHeaders -ErrorAction Stop -WarningAction Stop
							} else {
								$Query = Invoke-RestMethod -Headers $Headers -Uri $GraphUri -Method "GET" -ContentType "application/json" -ErrorAction Stop -WarningAction Stop
							}
							if ($Query.Value) {
									$Result += $Query.Value
							}	
							Else {
								$Result += $Query
							}
							$GraphUri = $Query.'@odata.nextlink'
							if ($ResponseHeaders) {
								$ResponseHeaders | fl
							}
						}
						catch {
							if ($ResponseHeaders) {
								$ResponseHeaders | fl
							}
							$msg = $_.Exception.Message
							foreach ($string in $GraphAPIRetryableErrors) {
								if ($msg.contains($string)) {
									$retry = $true
									$retryInterval = 10
									Continue
								}
							}
							foreach ($string in $GraphAPIThrottlingErrors) {
								if ($msg.contains($string)) {
									$retry = $true
									$retryInterval = 120
									Continue
								}
							}
							foreach ($string in $GraphAPIAuthErrors -and $AppRegName) {
								if ($msg.contains($string) -and $AppRegName) {
									$retry = $true
									$retryInterval = 10
									Request-MSALToken -AppRegName $AppRegName -Force
									$AccessToken = $script:AuthDB[$AppRegName].AccessToken
									Continue
								}
							}
							if ($retry) {
								#Write-Host "." -NoNewline -ForegroundColor Red -BackgroundColor Yellow
								Write-Log -String "ERR, retrying: $($retryCount) $($msg) $($GraphUri.Substring(0,80))" -MessageType Error
								Start-Sleep -Seconds $retryInterval
								$retryCount++
							}
							else {
								$script:GraphError = $true
								$script:GraphErrorMsg = $msg
								if ($Silent) {
									Write-Log -String "ERR: $($GraphUri) $($_.Exception.Message) $($retry) $($retryCount)" -MessageType Error -ForceOffScreen	
								}
								Else {
									Write-Log -String "ERR: $($GraphUri) $($_.Exception.Message) $($retry) $($retryCount)" -MessageType Error -ForceOnScreen
								}
							}
						}
					}
					Until (($retry -eq $false) -or ($retryCount -ge $Retries))
					if ($retryCount -gt 0) {
						if ($Silent) {
							Write-Log -String "Retries: $($retryCount)" -ForceOffScreen	
						}
						Else {
							Write-Log -String "Retries: $($retryCount)" -ForceOnScreen
						}
						
					}
				}
				Until (($null -eq $GraphUri) -or $script:GraphError)
			}
			"text/csv" { 
				[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
				$Query = Invoke-RestMethod -Uri $GraphUri -Headers $Headers -Method "GET" -ContentType "text/csv" -ErrorAction Stop -WarningAction Stop
				$Result = ConvertFrom-Csv -InputObject $Query
				if ($ProgressDots -and $interactiveRun -and (-Not $Silent)) {
					Write-Host "..." -NoNewline
				}
			}
		}	
	}
	Catch {
		$script:GraphError = $true
		$script:GraphErrorMsg = $true
		if ($Silent) {
			Write-Log -String "ERR: $($GraphUri) $($_.Exception.Message)" -ForceOffScreen
		}
		Else {
			Write-Log -String "ERR: $($GraphUri) $($_.Exception.Message)" -ForceOnScreen -ForegroundColor Magenta
		}
		
	}
	if ($ProgressDots -and $interactiveRun -and (-Not $Silent)) {
		Write-Host " done ($(Get-Count($Result)))"
	}
	Return $Result
}

########################################################################################
# New-GraphSecurityGroup
########################################################################################
function New-GraphSecurityGroup {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][string]$DisplayName,
		[Parameter(Mandatory)][string]$AccessToken,
		[string]$Description = [string]::Empty,
		[bool]$MailEnabled = $false,
		[string]$MailNickname = [string]::Empty
	)
	# main function body ##################################

	if (-not $MailNickname) {
		$MailNickname = $DisplayName -replace '[^a-zA-Z0-9]','-'
	}
	
	$params = @{
		displayName = $DisplayName
		mailEnabled = $MailEnabled
		mailNickname = $MailNickname
		securityEnabled = $true
	}
	if ($Description) {
		$params.Add("description",$Description)
	}

	$Body = $params | ConvertTo-Json
	write-host $Body
	$UriResource = "groups"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
	Try {
		$Result = Invoke-RestMethod -Headers @{Authorization = "Bearer $AccessToken"} -Uri $Uri -Method "POST" -Body $Body -ContentType $ContentTypeJSON
		Return $Result
	}
	Catch {
		Write-Host $_.Exception.Message -ForegroundColor Red
		Return $null
	}

}

########################################################################################
# Get-GraphUserById
########################################################################################
function Get-GraphUserById {
	[alias("Get-UserFromGraphById")]
	[CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$id,
        [Parameter(Mandatory)][string]$AccessToken,
		[string]$Properties = "id,userPrincipalName,DisplayName",
		[string][ValidateSet("v1.0","beta")]$Version="v1.0",
		[switch]$Dbg
    )
	# main function body ##################################
	$script:GraphError = $False
	$UriResource = "users/$($id)"
	$UriSelect = $Properties
	$Uri = New-GraphUri -Version $Version -Resource $UriResource -Select $UriSelect
	$retryCount = 0
	Do {
		try {
			$retry = $false
			$Result = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -Uri $Uri -Method GET -ContentType $ContentTypeJSON
		}
		catch {
			$script:GraphError = $true
			$result = $null
			$msg = $_.Exception.Message.ToLower()
			if ($msg.Contains("404") -or $msg.Contains("not found") -or $msg.Contains("notfound")) {
				$retry = $false
				$script:GraphErrorCode = 404
				if ($Dbg) {
					Write-Host  "Get-GraphUserById $($id) - $($msg)" -ForegroundColor Red -BackgroundColor Yellow
				}
			}
			Else {
				$retry = $true
				$retryCount++
				if ($Dbg) {
					Write-Host  "Get-GraphUserById $($id) - $($msg)" -ForegroundColor Red -BackgroundColor Yellow
				}
				Else {
					write-host "." -NoNewline -ForegroundColor Red -BackgroundColor Yellow
				}
				Start-Sleep -Seconds 10
			}
		}
	} Until (($retry -eq $false) -or ($retryCount -ge 10))
	Return $Result
}


########################################################################################
# Get-GraphServicePrincipalById
########################################################################################
function Get-GraphServicePrincipalById {
	[CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$id,
        [Parameter(Mandatory)][string]$AccessToken,
		[string]$Properties = "id,appId,appDisplayName,accountEnabled",
		[string][ValidateSet("v1.0","beta")]$Version="v1.0",
		[switch]$Dbg
    )
	# main function body ##################################
	$script:GraphError = $False
	$UriResource = "servicePrincipals/$($id)"
	$UriSelect = $Properties
	$Uri = New-GraphUri -Version $Version -Resource $UriResource -Select $UriSelect
	$retryCount = 0
	Do {
		try {
			$retry = $false
			$Result = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON
		}
		catch {
			$script:GraphError = $true
			$result = $null
			$msg = $_.Exception.Message.ToLower()
			if ($msg.Contains("404") -or $msg.Contains("not found") -or $msg.Contains("notfound")) {
				$retry = $false
				$script:GraphErrorCode = 404
				if ($Dbg) {
					Write-Host  "Get-GraphServicePrincipalById $($id) - $($msg)" -ForegroundColor Red -BackgroundColor Yellow
				}
			}
			Else {
				$retry = $true
				$retryCount++
				if ($Dbg) {
					Write-Host  "Get-GraphServicePrincipalById $($id) - $($msg)" -ForegroundColor Red -BackgroundColor Yellow
				}
				Else {
					write-host "." -NoNewline -ForegroundColor Red -BackgroundColor Yellow
				}
				Start-Sleep -Seconds 10
			}
		}
	} Until (($retry -eq $false) -or ($retryCount -ge 10))
	Return $Result
}

########################################################################################
# Get-GroupNameFromGraphById
########################################################################################
function Get-GroupNameFromGraphById {
	[CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$id,
        [Parameter(Mandatory)][string]$AccessToken,
		[string][ValidateSet("v1.0","beta")]$Version="v1.0"
    )
	# main function body ##################################
	$script:GraphError = $False
	$UriResource = "groups/$($id)"
	$Uri = New-GraphUri -Version $Version -Resource $UriResource
	$Group = Get-GraphOutputREST -Uri $Uri -AccessToken $AccessToken -ContentType $ContentTypeJSON
	If ($Group) {
		Return $Group.DisplayName
	}
	Else {
		Return $null
	}
}

########################################################################################
# Get-GroupMembersFromGraphById
########################################################################################
function Get-GroupMembersFromGraphById {
	[CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$id,
        [Parameter(Mandatory)][string]$AccessToken,
		[string][ValidateSet("v1.0","beta")]$Version="v1.0",
		[string]$Properties,
		[bool]$Transitive = $true,
		[switch]$ExcludeGuests
    )
	# main function body ##################################
	$script:GraphError = $False
	if ($Transitive) {
		$UriResource = "groups/$($id)/transitiveMembers"
	}
	else {
		$UriResource = "groups/$($id)/members"
	}
	if ($Properties) {
		$UriSelect = $Properties
	}
	else {
		$UriSelect = "id,mail,userPrincipalName"
	}
	$Uri = New-GraphUri -Version $Version -Resource $UriResource -Select $UriSelect -Count -Top 999
	#$Uri = New-GraphUri -Version $Version -Resource $UriResource -Top 999
	Try {
		[array]$Members = Get-GraphOutputREST -Uri $Uri -AccessToken $AccessToken -ContentType $ContentTypeJSON -ConsistencyLevel "eventual"
		If ($Members -and $Transitive) {
			$Members = $Members | Where-Object { $_."@odata.type" -eq $TypeUser }
		}
		If ($Members -and $ExcludeGuests) {
			$Members = $Members | Where-Object { $_.userPrincipalName -notlike "*$($GuestUPNSuffix)" }
		}
		Return $Members
	}
	Catch {
		write-host "ERROR: $($id) $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Yellow
		$script:GraphError = $true
		Return $null
	}
}

########################################################################################
# Get-GroupOwnersFromGraphById
########################################################################################
function Get-GroupOwnersFromGraphById {
	[CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$id,
        [Parameter(Mandatory)][string]$AccessToken,
		[string][ValidateSet("v1.0","beta")]$Version="v1.0",
		[string]$Properties
    )
	# main function body ##################################
	$script:GraphError = $False
	$UriResource = "groups/$($id)/owners"
	if ($Properties) {
		$UriSelect = $Properties
	}
	else {
		$UriSelect = "id,mail,userPrincipalName"
	}
	$Uri = New-GraphUri -Version $Version -Resource $UriResource -Select $UriSelect
	Try {
		[array]$Owners = Get-GraphOutputREST -Uri $Uri -AccessToken $AccessToken -ContentType $ContentTypeJSON -ConsistencyLevel "eventual"
		Return $Owners
	}
	Catch {
		write-host "ERROR: $($id) $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Yellow
		$script:GraphError = $true
		Return $null
	}
}

########################################################################################
# Add-GraphGroupMemberById
########################################################################################
function Add-GraphGroupMemberById {
	[CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$groupId,
		[Parameter(Mandatory)]$userId,
        [Parameter(Mandatory)][string]$AccessToken,
		[string][ValidateSet("v1.0","beta")]$Version="v1.0",
		[bool]$SkipCurrentMembers = $false,
		$TestRun = $false
    )
	# main function body ##################################
	$script:GraphError = $False
	$GroupName = Get-GroupNameFromGraphById -Id $GroupId -AccessToken $AccessToken
	if ($SkipCurrentMembers) {
		$UriResource = "groups/$($GroupId)/members"
		$UriSelect = "id,userPrincipalName"
		$Uri = New-GraphUri -Version $Version -Resource $UriResource -Select $UriSelect -Top 999
		[array]$CurrentMembers = Get-GraphOutputREST -Uri $Uri -AccessToken $AccessToken -ContentType $ContentTypeJSON
	}
	$addCounter = 0
	try {
        $Headers = @{Authorization = "Bearer $($AccessToken)"}
        if (($userId.count) -eq 1) {
			$Method = "POST"
			if ($userId -is [array]) {
				$userId = [string]$userId[0]
			}
			if (-not($CurrentMembers) -or ($CurrentMembers -and ($CurrentMembers.id -NotContains $userId))) {
				$UriResource = "groups/$($GroupId)/members/`$ref"
				$Uri = New-GraphUri -Version $Version -Resource $UriResource
				$Body = @{
					"@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($userId)"
				}
				$Body = $Body | ConvertTo-Json
				#write-host $Body
				if (-not $TestRun) {
					$Result = Invoke-RestMethod -Uri $Uri -Headers $Headers -Body $Body -Method $Method -ContentType $ContentTypeJSON
				}
				$addCounter++
			}
		}
		else {
			$Method = "PATCH"
	        $UriResource = "groups/$($GroupId)"
			$PageSize = 10
			$StartIndex = 0
			Do {
				if (($StartIndex + $PageSize) -le $userId.Count) {
					$EndIndex = $StartIndex + $PageSize - 1
				}
				else {
					$EndIndex = $userId.Count - 1
				}
				$ODataArray = @()
				for ($i=$StartIndex; $i -le $EndIndex; $i++) {
					if (-not($CurrentMembers) -or ($CurrentMembers -and ($CurrentMembers.id -NotContains $userId[$i]))) {
						$ODataArray += "https://graph.microsoft.com/v1.0/directoryObjects/$($userId[$i])"
					}
				}
				$Body = @{
					"members@odata.bind" = $ODataArray
				}
				$Uri = New-GraphUri -Version $Version -Resource $UriResource
				$Body = $Body | ConvertTo-Json
				#write-host $Body
				if (-not $TestRun) {
					$Result = Invoke-RestMethod -Uri $Uri -Headers $Headers -Body $Body -Method $Method -ContentType $ContentTypeJSON
				}
				$addCounter += $ODataArray.Count
				$StartIndex = $EndIndex + 1
			} Until ($EndIndex -eq $userId.Count - 1)
		}
		write-host "Added $($addCounter) member(s) to group $($GroupName)" -ForegroundColor Green
		Return $Result
	}
    Catch {
        write-host $Uri -ForegroundColor Red
		write-host $Body -ForegroundColor Red
		$script:GraphError = $true
		Return "ERROR " + $_.Exception.Message
    }
}


########################################################################################
# Add-GraphGroupOwnerById
########################################################################################
function Add-GraphGroupOwnerById {
	[CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$groupId,
		[Parameter(Mandatory)]$userId,
        [Parameter(Mandatory)][string]$AccessToken,
		[string][ValidateSet("v1.0","beta")]$Version="v1.0",
		[bool]$SkipCurrentOwners = $false
    )
	# main function body ##################################
	$script:GraphError = $False
	if ($SkipCurrentOwners) {
		$UriResource = "groups/$($GroupId)/owners"
		$UriSelect = "id,userPrincipalName"
		$Uri = New-GraphUri -Version $Version -Resource $UriResource -Select $UriSelect
		[array]$CurrentOwners = Get-GraphOutputREST -Uri $Uri -AccessToken $AccessToken -ContentType $ContentTypeJSON
	}

	try {
        $Headers = @{Authorization = "Bearer $($AccessToken)"}
			$Method = "POST"
			if (-not($CurrentOwners) -or ($CurrentOwners -and ($CurrentOwners.id -NotContains $userId))) {
				$UriResource = "groups/$($GroupId)/owners/`$ref"
				$Uri = New-GraphUri -Version $Version -Resource $UriResource
				$Body = @{
					"@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($userId)"
				}
				$Body = $Body | ConvertTo-Json
				#write-host $Body
				$Result = Invoke-RestMethod -Uri $Uri -Headers $Headers -Body $Body -Method $Method -ContentType $ContentTypeJSON
			}
		Return $Result
	}
    Catch {
        write-host $Uri -ForegroundColor Red
		write-host $Body -ForegroundColor Red
		$script:GraphError = $true
		Return "ERROR " + $_.Exception.Message
    }
}

########################################################################################
# Remove-GraphGroupMemberById
########################################################################################
function Remove-GraphGroupMemberById {
	[CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$groupId,
		[Parameter(Mandatory)][string]$userId,
        [Parameter(Mandatory)][string]$AccessToken,
		[string][ValidateSet("v1.0","beta")]$Version="v1.0"
    )
	# main function body ##################################
	try {
        $UriResource = "groups/$($GroupId)/members/$userId/`$ref"
        $Uri = New-GraphUri -Version $Version -Resource $UriResource
        $Headers = @{Authorization = "Bearer $($AccessToken)"}
        $Result = Invoke-RestMethod -Uri $Uri -Headers $Headers -Method "DELETE" -ContentType $ContentTypeJSON
		Return $Result
	}
    Catch {
        Return "ERROR " + $_.Exception.Message
    }
}

########################################################################################
# Remove-GraphGroupOwnerById
########################################################################################
function Remove-GraphGroupOwnerById {
	[CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$groupId,
		[Parameter(Mandatory)][string]$userId,
        [Parameter(Mandatory)][string]$AccessToken,
		[string][ValidateSet("v1.0","beta")]$Version="v1.0"
    )
	# main function body ##################################
	try {
        $UriResource = "groups/$($GroupId)/owners/$userId/`$ref"
        $Uri = New-GraphUri -Version $Version -Resource $UriResource
        $Headers = @{Authorization = "Bearer $($AccessToken)"}
        $Result = Invoke-RestMethod -Uri $Uri -Headers $Headers -Method "DELETE" -ContentType $ContentTypeJSON
		Return $Result
	}
    Catch {
        Return "ERROR " + $_.Exception.Message
    }
}

########################################################################################
# Sync-GraphGroups
########################################################################################
function Sync-GraphGroups {
    param(
        [Parameter(Mandatory=$true)][string]$SourceGroup,
        [Parameter(Mandatory=$true)][string]$TargetGroup,
        [Parameter(Mandatory=$false)][string]$AccessToken,
        [bool]$Mirror = $true
    )
	
	if ($SourceGroup -eq $TargetGroup) {
		Write-Log "Source and target groups are the same, nothing to do." -MessageType Warning
		return
	}
	
	$SourceGroupName = Get-GroupNameFromGraphById -id $SourceGroup -AccessToken $AccessToken
	if (-not $SourceGroupName) {
		Write-Log "Source group $($SourceGroup) not found." -MessageType Error
		return
	}
	
	$TargetGroupName = Get-GroupNameFromGraphById -id $TargetGroup -AccessToken $AccessToken
	if (-not $TargetGroupName) {
		Write-Log "Target group $($TargetGroup) not found." -MessageType Error
		return
	}

	[array]$SourceMembers = (Get-GroupMembersFromGraphById -id $SourceGroup -AccessToken $AccessToken).id
    if ($SourceMembers.Count -eq 0) {
		Write-Log "Source group $($SourceGroup) has no members." -MessageType Warning
		return
	}
	
	[array]$TargetMembers = (Get-GroupMembersFromGraphById -id $TargetGroup -AccessToken $AccessToken).id
    
	$MembersToAdd = $SourceMembers | Where-Object { $_ -notin $TargetMembers }
	$MembersToRemove = $TargetMembers | Where-Object { $_ -notin $SourceMembers }

	$counterAdd = $counterRemove = 0
	foreach ($Member in $MembersToAdd) {
        Write-Log "Adding $($Member) to $($TargetGroupName) ($($TargetGroup))"
		$result = Add-GraphGroupMemberById -GroupId $TargetGroup -userId $Member -AccessToken $AccessToken
		$counterAdd++
    }
    if ($Mirror) {
        foreach ($Member in $MembersToRemove) {
			Write-Log "Removing $($Member) from $($TargetGroupName) ($($TargetGroup))"
            $result = Remove-GraphGroupMemberById -GroupId $TargetGroup -userId $Member -AccessToken $AccessToken
			$counterRemove++
        }
    }
		Write-Log "source:$($SourceGroup) target:$($TargetGroup) mirror:$($Mirror) added:$($counterAdd) removed:$($counterRemove)"
	Write-Log "======================================================"
}

########################################################################################
# Update-GraphGroupMembersById
########################################################################################
Function Update-GraphGroupMembersById {    
	param (        
		[parameter(Mandatory = $true)][string]$AccessToken,        
		[parameter(Mandatory = $true)][string]$TargetGroupId,        
		[parameter(Mandatory = $true)][array]$SourceGroupArray,        
		[parameter(Mandatory = $true)][string][ValidateSet("device","user")]$GroupType,
		[switch]$FastMode  
		)    
	$GroupName = Get-GroupNameFromGraphById -id $TargetGroupId -AccessToken $AccessToken
	Write-Log "Updating group $($GroupName) ($($TargetGroupId))"
	$UriResource = "groups/$($TargetGroupId)/members"    
	$UriSelect = "id,displayName"    
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Select $UriSelect 
	[array]$ASIS_Members = (Get-GraphOutputREST -Uri $Uri -AccessToken $AccessToken -ContentType $ContentTypeJSON).id
	$SourceGroupArray = $SourceGroupArray | Sort-Object -Unique    
	$MISSING_Members = $SourceGroupArray | Where-Object { $ASIS_Members -notcontains $_ }    
	$EXTRA_Members = $ASIS_Members | Where-Object { $SourceGroupArray -notcontains $_ }    

	Write-Log "Missing $($GroupType)s: $($MISSING_Members.count)"    
	if ($MISSING_Members.Count -gt 0) {        
		foreach ($Id in $MISSING_Members) {            
			if ($FastMode) {  
				Write-Log "Adding $($GroupType.ToLower()) id $($id)"
			}
			else {				
				if ($GroupType -eq "device") {                
					$missingObject = Get-GraphDeviceById -AccessToken $AccessToken -id $id            
				}            
				else {                
					$missingObject = Get-GraphUserById -AccessToken $AccessToken -id $id           
				}
				Write-Log "Adding $($GroupType.ToLower()) $($missingObject.displayName.PadRight(50)) (AAD id: $($id))" 	          
			}
			Add-GraphGroupMemberById -AccessToken $AccessToken -GroupId $TargetGroupId -userId $Id        
		}        
	}
	
	Write-Log "Extra $($($GroupType))s:   $($EXTRA_Members.count)"   
	if ($EXTRA_Members.Count -gt 0) {
		foreach ($Id in $EXTRA_Members) {            
			if ($FastMode) {                
				Write-Log "Removing $($GroupType.ToLower()) id $($id)"            
			}            
			else {
				if ($GroupType -eq "device") {
					$extraObject = Get-GraphDeviceById -AccessToken $AccessToken -id $id
				}
				else {
					$extraObject = Get-GraphUserById -AccessToken $AccessToken -id $id
				}
				Write-Log "Removing $($GroupType.ToLower()) $($extraObject.displayName.PadRight(50)) (AAD id: $($id))"    
			}       
			Remove-GraphGroupMemberById -AccessToken $AccessToken -GroupId $TargetGroupId -userId $Id
		}        
	}
}

########################################################################################
# Update-GraphGroupOwnersById
########################################################################################
Function Update-GraphGroupOwnersById {    
	param (        
		[parameter(Mandatory = $true)][string]$AccessToken,        
		[parameter(Mandatory = $true)][string]$TargetGroupId,        
		[parameter(Mandatory = $true)][array]$SourceGroupArray        
		)    
	$GroupName = Get-GroupNameFromGraphById -id $TargetGroupId -AccessToken $AccessToken
	Write-Log "Updating group $($GroupName) ($($TargetGroupId))"
	$UriResource = "groups/$($TargetGroupId)/owners"    
	$UriSelect = "id,displayName"    
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect 
	[array]$ASIS_Owners = (Get-GraphOutputREST -Uri $Uri -AccessToken $AccessToken -ContentType $ContentTypeJSON).id
	$SourceGroupArray = $SourceGroupArray | Sort-Object -Unique    
	$MISSING_Owners = $SourceGroupArray | Where-Object { $ASIS_Owners -notcontains $_ }    
	$EXTRA_Owners = $ASIS_Owners | Where-Object { $SourceGroupArray -notcontains $_ }    

	Write-Log "Missing owners: $($MISSING_Owners.count)"    
	if ($MISSING_Owners.Count -gt 0) {        
		foreach ($Id in $MISSING_Owners) {            
			$missingObject = Get-GraphUserById -AccessToken $AccessToken -id $id           
			Write-Log "Adding user $($missingObject.displayName.PadRight(50)) (AAD id: $($id))" 	          
			Add-GraphGroupOwnerById -AccessToken $AccessToken -GroupId $TargetGroupId -userId $Id -SkipCurrentOwners $true
		}        
	}

	Write-Log "Extra owners:   $($EXTRA_Owners.count)"   
	if ($EXTRA_Owners.Count -gt 0) {
		foreach ($Id in $EXTRA_Owners) {            
			$extraObject = Get-GraphUserById -AccessToken $AccessToken -id $id
			Write-Log "Removing user $($extraObject.displayName.PadRight(50)) (AAD id: $($id))"         
			Remove-GraphGroupOwnerById -AccessToken $AccessToken -GroupId $TargetGroupId -userId $Id
		}        
	}
}

function Remove-B2BUser {
    param (
        [Parameter(Mandatory=$true)][string]$Identity,
        [Parameter(Mandatory=$true)][string]$AccessToken,
		$Silent = $false
    )
    $UriResource = "users/$($Identity)"
    $UriSelect = "id,userPrincipalName,displayName,mail,userType,onPremisesExtensionAttributes"
    $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Select $UriSelect
    $AADGuest = Get-GraphOutputREST -Uri $Uri -AccessToken $AccessToken -ContentType $ContentTypeJSON

    if ($AADGuest) {
        if ($AADGuest.userType -ne "Guest") {
            Write-Log "ERROR:  User $($AADGuest.userPrincipalName) is not a B2B Guest user" -ForegroundColor "Red"
            return
        }
        $ext15 = $Guest.onPremisesExtensionAttributes.extensionAttribute15
        if ($ext15 -and ($ext15.StartsWith("XTSync_"))) {
            Write-Log "ERROR:  B2B user $($AADGuest.userPrincipalName) is synced from tenant $($ext15.Substring(7))" -ForegroundColor "Red"
            return
        }
        else {
            Try {
                $UriResource = "users/$($Identity)"
                $Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
                $Uri = [System.Web.HttpUtility]::UrlPathEncode($Uri)
                $Headers = @{Authorization = "Bearer $($AccessToken)"}
                $ResultDELETE = Invoke-RestMethod -Headers $Headers -Uri $Uri -Method "DELETE" -ContentType $ContentTypeJSON
                if (-not $Silent) {
					Write-Log "SUCCESS removed $($AADGuest.userPrincipalName) ($($AADGuest.id))" -ForegroundColor "Green"
                }
            }
            Catch {
                $ErrorMessageDELETE = $($_.Exception.Message)
                Write-Log "ERR DELETE $($Mail) $($Identity) $($ErrorMessageDELETE)" -MessageType Error
            }
        }
    }
    else {
        Write-Log "ERROR: B2B user $($Identity) does not exist" -ForegroundColor "Red"
        return
    }
}


########################################################################################
# Get-AADGroupListByOnpremOU
########################################################################################
function Get-AADGroupListByOnpremOU {
	[CmdletBinding()]
	param (
		[pscredential]$Credential,
		[Parameter(Mandatory)][string]$AccessToken,
		[Parameter(Mandatory)][string]$OU,
		[string]$Properties,
		[string]$Filter
	)
	# main function body ##################################
	[array]$ADGroupSidList = @()
	[array]$GroupReport = @()
	if ($Filter) {
		$GroupFilter = $Filter
	}
	else {
		$GroupFilter = "*"
	}
	#get AAD groups
	$UriResource = "groups"
	if ($Properties) {
		$UriSelect = $Properties
		$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999	
	}
	else {
		$UriSelect = "id,OnPremisesSecurityIdentifier,DisplayName"
		$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Select $UriSelect
	}
	$AADGroups = Get-GraphOutputREST -Uri $Uri -AccessToken $AccessToken -ContentType $ContentTypeJSON

	#read groups from AD OU
	Try {
		if ($Credential) {
			$ADGroups = Get-ADGroup -Credential $Credential -Filter $GroupFilter -SearchBase $OU
		}
		else {
			$ADGroups = Get-ADGroup -Filter $GroupFilter -SearchBase $OU
		}
	}
	Catch {
		return $null
	}
	write-host "$($OU) AAD groups: $($AADGroups.Count) AD groups: $($ADGroups.Count)"
	if ($AADGroups -and $ADGroups){
		foreach ($group in $ADGroups) {
			$ADGroupSidList += $group.sid.ToString().Trim()
		}
		foreach ($group in $AADGroups) {
			if ($group.onPremisesSecurityIdentifier) {
				$sid = $group.onPremisesSecurityIdentifier.ToString().Trim()
				if ($ADGroupSidList.Contains($sid)) {
					$GroupReport += [pscustomobject]@{
						AAD_id = $group.id;
						AAD_displayName = $group.displayName;
						AD_sid = $sid
					}
				}
			}
		}
		write-host "Matching groups: $($GroupReport.Count)"
		return $GroupReport
	}
	else {
		return $null
	}
}

########################################################################################
# Get-AADGroupMemberReportByIdList
########################################################################################
function Get-AADGroupMemberReportByIdList {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][string]$AccessToken,
		[Parameter(Mandatory)][array]$GroupList,
		[string]$Properties
	)
	# main function body ##################################


	foreach ($groupId in $GroupList) {
		$members = Get-GroupMembersFromGraphById -id $group.AAD_id -AccessToken $AccessToken
		write-host "Group: $($group.AAD_DisplayName) $($group.AAD_id) has $($members.Count) members"
		foreach ($member in $members) {
			$memberObject = [pscustomobject]@{
				GroupId = $groupId
				GroupName = $group.AAD_DisplayName
				UserId = $member.id
				DisplayName = $member.displayName
				UserPrincipalName = $member.userPrincipalName
				Mail = $member.mail
				JobTitle = $member.jobTitle
			}
			$GroupReportPwrMem += $memberObject
		}
	}
}

########################################################################################
# Get-ADGroupMembersUpn
########################################################################################
function Get-ADGroupMembersUpn {
	[CmdletBinding()]
	param (
		$CredentialFile,
		$Credential,
		[Parameter(Mandatory)][string]$Identity
	)
	# main function body ##################################
	[array]$ADGroupMembersUpn = @()
	# load credential file
	if (-not $Credential) {
		if ($CredentialFile) {
			Try {
				$Credential = Import-Clixml -Path $ADCredentialPath
			}
			Catch {
				return $null
			}
		}
		else {
			return $null
		}
	}

	#read groups from AD OU
	Try {
		if ($Credential) {
			$ADGroupMembersUpn = Get-ADGroupMember -Credential $Credential -Identity $Identity | ForEach-Object { Get-ADUser $_.samaccountname -Credential $Credential | Select-Object userPrincipalName }
		}
		else {
			$ADGroupMembersUpn = Get-ADGroupMember -Identity $Identity | ForEach-Object { Get-ADUser $_.samaccountname | Select-Object userPrincipalName }
		}
		return $ADGroupMembersUpn
	}
	Catch {
		return $null
	}
}

########################################################################################
# Convert-CollectionToString
########################################################################################
function Convert-CollectionToString {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)]
		$Collection,
		[string]$Delim = ","
    )
	# main function body ##################################
	[string]$result = $null
	if ($Collection.Count -gt 0) {
		foreach ($member in $Collection) {
			$result = $result + $member + $Delim
		}
		$result = $result.Trim($Delim)
	}
	return $result
}

########################################################################################
# Convert-SecToHMS
########################################################################################
function Convert-SecToHMS {
	[CmdletBinding()]
    param (
		[Parameter(Position=0)][int]$Seconds
    )
	# main function body ##################################
	$ts =  [timespan]::fromseconds($Seconds)
	return  $ts.ToString("dd'days:'hh'hr:'mm'min:'ss'sec'")
}

########################################################################################
# Test-IsServiceAccount
########################################################################################
function Test-IsServiceAccount {
	[CmdletBinding()]
    param (
		[Parameter(Position=0)][string]$Upn
    )
	# main function body ##################################
	$first2 = $Upn.SubString(0,2).ToUpper()
	if ($script:knownServiceAccountPrefixes.Contains($first2)) {
		return $true 
	}
	else {
		return $false 
	}
}

########################################################################################
# Test-ContainsOfficeLicense
########################################################################################
function Test-ContainsOfficeLicense {
	[CmdletBinding()]
    param (
		[Parameter(Position=0)][string]$LicenseString
    )
	# main function body ##################################
	$lic = $LicenseString.ToUpper() 
	if ($lic.Contains("MICROSOFT 365 E") -or $lic.Contains("MICROSOFT 365 F") -or $lic.Contains("OFFICE 365 E")) {
		return $true 
	}
	else {
		return $false 
	}
}

########################################################################################
# Test-ContainsOfficeLicense
########################################################################################
function Test-ContainsMTRProLicense {
	[CmdletBinding()]
    param (
		[Parameter(Position=0)][string]$LicenseString
    )
	# main function body ##################################
	$lic = $LicenseString.ToUpper() 
	if ($lic.Contains("MICROSOFT TEAMS ROOMS PRO")) {
		return $true 
	}
	else {
		return $false 
	}
}


########################################################################################
# Get-MailFromGuestUPN
########################################################################################
function Get-MailFromGuestUPN {
	[CmdletBinding()]
	[alias("GetMailFromGuestUPN")]
    param (
		[Parameter(Position=0)][string]$GuestUPN
    )
	# main function body ##################################
	Try {
		$rawMail = ($GuestUPN -split "#ext#@")[0]
		$lastUnderscoreIndex = $rawMail.LastIndexOf("_")
		if ($lastUnderscoreIndex -ge 0) {
			$localPart = $rawMail.Substring(0, $lastUnderscoreIndex)
			$domainPart = $rawMail.Substring($lastUnderscoreIndex + 1)
			return "$($localPart)@$($domainPart)".ToLower()
		} else {
			return $null
		}
	}
	Catch {
		return $null
	}
}

########################################################################################
# Get-DomainFromGuestUPN
########################################################################################
function Get-DomainFromGuestUPN {
	[CmdletBinding()]
	param (
		[Parameter(Position=0)][string]$GuestUPN
    )
	# main function body ##################################
	Try {
		$GuestMail = Get-MailFromGuestUPN -GuestUPN $GuestUPN
		return ($GuestMail.Split("@")[1]).ToLower()
	}
	Catch {
		return $null
	}
}

########################################################################################
# Get-DomainFromAddress
########################################################################################
function Get-DomainFromAddress {
	[CmdletBinding()]
	param (
		[Parameter(Position=0)][string]$Address
    )
	# main function body ##################################
	Try {
		return ($Address.Split("@")[1]).ToLower()
	}
	Catch {
		return $null
	}
}

########################################################################################
# Convert-IdentitiesToString
########################################################################################
function Convert-IdentitiesToString {
	[CmdletBinding()]
	[alias("IdentitiesToString")]
    param (
		[Parameter(Mandatory)][AllowNull()]$Identities
    )
	# main function body ##################################
	if ($Identities) {
		$strIdentity0 = $Identities | ConvertTo-Json -Compress | Out-String
		$strIdentity1 = $stridentity0.Replace(":","=")
		$strIdentity2 = $strIdentity1.Replace('"',"")
		$strIdentity3 = $strIdentity2.Replace(',',"__")
		return ($strIdentity3.Trim("`t`n`r")).ToLower()
	}
	else {
		return $null
	}

}

########################################################################################
# Convert-TextToBool
########################################################################################
Function Convert-ValueToBool {
	[CmdletBinding()]
    Param (
		[Parameter(Mandatory)][AllowNull()]$Value
    )
	# main function body ##################################
	[bool]$Result = $False
	If (($null -eq $Value) -or ($Value -eq "")) {
		$Result = $False
	}
	Else {
		If ($Value -is [int]) {
			If ($Value -eq 0) {
				$Result = $False
			}
			else {
				$Result = $True
			}
		}
		Else {
			If ($Value -eq "False") {
				$Result = $False
			}
			Else {
				If ($Value -eq "True") {
					$Result = $True
				}
			}
		}
	}
	Return $Result
}

########################################################################################
# Get-PrimarySMTPAddress
########################################################################################
function Get-PrimarySMTPAddress {
	[CmdletBinding()]
	[alias("GetPrimarySMTPAddress")]
    param (
		[Parameter(Mandatory)][AllowNull()]$Addresses
    )
	# main function body ##################################
	if ($Addresses -is [string]) {
		$addrArray = $Addresses.Split(",")
	} 
	else {
		$addrArray = $Addresses
	}
	foreach ($addr in $addrArray) {
		if ($addr.StartsWith("SMTP:")) {
			return $addr.SubString(5).ToLower()
		}
		else {
			return $null
		}
	}
}

function Get-IntlFormatPhoneNumber {
	[CmdletBinding()]
    param (
		[Parameter(Mandatory)][AllowNull()]$PhoneNumber,
		[switch]$EntraMFAFormat
    )
	# main function body ##################################
	if ($PhoneNumber) {
		[string]$PhoneString = $PhoneNumber.Trim()
		[int]$prefixLength=0
  
		foreach ($char in $phoneNumberRemoveChars) {
			$PhoneString = $PhoneString.Replace($char,"")
		}
		
		while ($PhoneString.StartsWith("0")) {
			$PhoneString = $PhoneString.Substring(1)
		}
		
		if ($prefixLength -eq 0) {
			foreach ($prefix in $prefixLength1) {
				If ($PhoneString.StartsWith($prefix)) {
					$prefixLength=1
					break
				}
			}
		}
		
		if ($prefixLength -eq 0) {
			foreach ($prefix in $prefixLength2) {
				If ($PhoneString.StartsWith($prefix)) {
					$prefixLength=2
					break
				} 
			}
		}
		
		if ($prefixLength -eq 0) {
			foreach ($prefix in $prefixLength3) {
				If ($PhoneString.StartsWith($prefix)) {
					$prefixLength = 3
					break
				} 
			}
		}
	
		If ($prefixLength -gt 0) {
			$prfx = $PhoneString.Substring(0,$prefixLength)
			$nmbr = $PhoneString.Substring($prefixLength)
			If ($nmbr.StartsWith("0")) {
				$nmbr = $nmbr.Substring(2)
			}
			if ($EntraMFAFormat) {
				$Result = "+" + $prfx + [char]32 + $nmbr
			}
			else {
				$Result = "+" + $prfx + $nmbr
			}
			Return $Result
		}
		Else{
			Return $null		
		}
	}
	Else {
	  Return $null
	}
  }
  
  Function Remove-File {
    param (
		[parameter(Mandatory = $true)][string]$Path
	)
	# main function body ##################################
    If (Test-Path -Path $Path) {
        Try {
            Remove-Item -Path $Path
        }
        Catch {
            Write-Host "Error removing $($Path)"
        }
    }
}

Function Test-StringContainsAnyArrayMember {
    param (
		[parameter(Mandatory = $true)][string]$String,
		[parameter(Mandatory = $true)][array]$Array
	)
	# main function body ##################################
    $Result = $False
    foreach ($Member in $Array) {
        if ($String.ToLower().Contains($Member.ToLower())) {
            $Result = $True
            Break
        }
    }
    Return $Result
}

Function Test-StringStartsWithUTC {
    param (
		[parameter(Mandatory = $true)][string]$String
	)
	# main function body ##################################
    Try {
        $UTC = [datetime]::ParseExact($String.Substring(0,19), "yyyy-MM-ddTHH:mm:ss", $null)
        Return $True
    }
    Catch {
        Return $False
    }
    
}


# SIG # Begin signature block
# MIIpXgYJKoZIhvcNAQcCoIIpTzCCKUsCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC+VEr7NVd0ozaw
# c/FS+godfZeiauVBvaITNKP/WEvVMKCCDhMwggawMIIEmKADAgECAhAIrUCyYNKc
# TJ9ezam9k67ZMA0GCSqGSIb3DQEBDAUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNV
# BAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0yMTA0MjkwMDAwMDBaFw0z
# NjA0MjgyMzU5NTlaMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwg
# SW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcg
# UlNBNDA5NiBTSEEzODQgMjAyMSBDQTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAw
# ggIKAoICAQDVtC9C0CiteLdd1TlZG7GIQvUzjOs9gZdwxbvEhSYwn6SOaNhc9es0
# JAfhS0/TeEP0F9ce2vnS1WcaUk8OoVf8iJnBkcyBAz5NcCRks43iCH00fUyAVxJr
# Q5qZ8sU7H/Lvy0daE6ZMswEgJfMQ04uy+wjwiuCdCcBlp/qYgEk1hz1RGeiQIXhF
# LqGfLOEYwhrMxe6TSXBCMo/7xuoc82VokaJNTIIRSFJo3hC9FFdd6BgTZcV/sk+F
# LEikVoQ11vkunKoAFdE3/hoGlMJ8yOobMubKwvSnowMOdKWvObarYBLj6Na59zHh
# 3K3kGKDYwSNHR7OhD26jq22YBoMbt2pnLdK9RBqSEIGPsDsJ18ebMlrC/2pgVItJ
# wZPt4bRc4G/rJvmM1bL5OBDm6s6R9b7T+2+TYTRcvJNFKIM2KmYoX7BzzosmJQay
# g9Rc9hUZTO1i4F4z8ujo7AqnsAMrkbI2eb73rQgedaZlzLvjSFDzd5Ea/ttQokbI
# YViY9XwCFjyDKK05huzUtw1T0PhH5nUwjewwk3YUpltLXXRhTT8SkXbev1jLchAp
# QfDVxW0mdmgRQRNYmtwmKwH0iU1Z23jPgUo+QEdfyYFQc4UQIyFZYIpkVMHMIRro
# OBl8ZhzNeDhFMJlP/2NPTLuqDQhTQXxYPUez+rbsjDIJAsxsPAxWEQIDAQABo4IB
# WTCCAVUwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUaDfg67Y7+F8Rhvv+
# YXsIiGX0TkIwHwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0P
# AQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGCCsGAQUFBwEBBGswaTAk
# BggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAC
# hjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9v
# dEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNybDAcBgNVHSAEFTATMAcGBWeBDAED
# MAgGBmeBDAEEATANBgkqhkiG9w0BAQwFAAOCAgEAOiNEPY0Idu6PvDqZ01bgAhql
# +Eg08yy25nRm95RysQDKr2wwJxMSnpBEn0v9nqN8JtU3vDpdSG2V1T9J9Ce7FoFF
# UP2cvbaF4HZ+N3HLIvdaqpDP9ZNq4+sg0dVQeYiaiorBtr2hSBh+3NiAGhEZGM1h
# mYFW9snjdufE5BtfQ/g+lP92OT2e1JnPSt0o618moZVYSNUa/tcnP/2Q0XaG3Ryw
# YFzzDaju4ImhvTnhOE7abrs2nfvlIVNaw8rpavGiPttDuDPITzgUkpn13c5Ubdld
# AhQfQDN8A+KVssIhdXNSy0bYxDQcoqVLjc1vdjcshT8azibpGL6QB7BDf5WIIIJw
# 8MzK7/0pNVwfiThV9zeKiwmhywvpMRr/LhlcOXHhvpynCgbWJme3kuZOX956rEnP
# LqR0kq3bPKSchh/jwVYbKyP/j7XqiHtwa+aguv06P0WmxOgWkVKLQcBIhEuWTatE
# QOON8BUozu3xGFYHKi8QxAwIZDwzj64ojDzLj4gLDb879M4ee47vtevLt/B3E+bn
# KD+sEq6lLyJsQfmCXBVmzGwOysWGw/YmMwwHS6DTBwJqakAwSEs0qFEgu60bhQji
# WQ1tygVQK+pKHJ6l/aCnHwZ05/LWUpD9r4VIIflXO7ScA+2GRfS0YW6/aOImYIbq
# yK+p/pQd52MbOoZWeE4wggdbMIIFQ6ADAgECAhACpA3ePY9KcbYaWBDrkJ+xMA0G
# CSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwg
# SW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcg
# UlNBNDA5NiBTSEEzODQgMjAyMSBDQTEwHhcNMjYwMTI3MDAwMDAwWhcNMjcwMTI2
# MjM1OTU5WjBjMQswCQYDVQQGEwJDWjEOMAwGA1UEBxMFUHJhaGExITAfBgNVBAoM
# GMSMRVogSUNUIFNlcnZpY2VzLCBhLiBzLjEhMB8GA1UEAwwYxIxFWiBJQ1QgU2Vy
# dmljZXMsIGEuIHMuMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAr17W
# fJTZXULA+4gGqSAyiB7+Hmyq8Yj2rhwu9kssJM/a4iQYQl+ymhjhMDsHolpoxSIK
# RfddrC9AknLsrzNN3Fs7ayWq+98olBQFEOf13+4XDehMljRWewQ2kLC6RsKJWLGR
# M/gaUa3Z/AsUKDh/fLic4r2eadoaGJxHhZ/MMKbwZ1ZU7ZbIhFQfCNvxLUSOtcq7
# yAw54pZT22ty8v4oSz0y/T3R9wrOoaWWqVqERF3Knc6FMRIXZukxaYacLiJ30DG8
# UTGEvlgZ7xybw4q5b7InQ8VLnMMZE/REtomi1sCSygAvToIkzuopVFZ9d55XUZKU
# sS3F+4CmsVKGTRCcm/ED6LFf3V7YHJInGa8k+Ycu820m99sBv8UtPZbpLZXhqPJU
# Ucrltk5/BePQ9g3bYMa4Tqw6ckxqOoY5iTqBQet3BE1kL5ndQGvMOpr72m4uXugm
# rGghIge+zXHDbBQaiiJcQe9zhl6Vkffs6RpyBWpCTtTSSoV8JI8sqBZXH5eylIeM
# uaJ252znv0+eh7IvYIKrzQd+mHutyv+T478UCl9DtUCTjtp/KyCfwZjTcxYZGrgC
# 0/ehGZLez2ZJ33Wj47myKDb9MkIiQnMkXSTSvg6eWqNJTsnkCDS7oLwVwFNt1MjF
# vMXeRJiBWet+Fpks8LdDZrXQoqjAY/l4LkpytQ8CAwEAAaOCAgMwggH/MB8GA1Ud
# IwQYMBaAFGg34Ou2O/hfEYb7/mF7CIhl9E5CMB0GA1UdDgQWBBRxuve8b/bQKh3K
# 734LdB5uue4QwjA+BgNVHSAENzA1MDMGBmeBDAEEATApMCcGCCsGAQUFBwIBFhto
# dHRwOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMIG1BgNVHR8Ega0wgaowU6BRoE+GTWh0dHA6Ly9jcmwz
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNENvZGVTaWduaW5nUlNBNDA5
# NlNIQTM4NDIwMjFDQTEuY3JsMFOgUaBPhk1odHRwOi8vY3JsNC5kaWdpY2VydC5j
# b20vRGlnaUNlcnRUcnVzdGVkRzRDb2RlU2lnbmluZ1JTQTQwOTZTSEEzODQyMDIx
# Q0ExLmNybDCBlAYIKwYBBQUHAQEEgYcwgYQwJAYIKwYBBQUHMAGGGGh0dHA6Ly9v
# Y3NwLmRpZ2ljZXJ0LmNvbTBcBggrBgEFBQcwAoZQaHR0cDovL2NhY2VydHMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0Q29kZVNpZ25pbmdSU0E0MDk2U0hB
# Mzg0MjAyMUNBMS5jcnQwCQYDVR0TBAIwADANBgkqhkiG9w0BAQsFAAOCAgEAZv1X
# ZTBjkfBEBRtgDDbXFCX542eqsZi4pUPNKO4JoDzszPa8kd+T/fmt1LZuY023yAgb
# vvVzVpznu8r6i9UmRVMUc/sHl2kfA3jS8w4fljy9kBJJltIp5yXzKt5GeJYL73II
# ik3xL+YwGOGcBfE93v4hNIV9YCa+D/WJpExOmfdTivgvvLvikFVRNc5mgqKyaAMa
# 2yl77dS9WxZoy/iY3QpFhruXEih+XU145K4d+Z74hmeyNH2KoAwmaq7IsDUZkKt4
# EPjVPKtjVv99yNwWWuYIwz6QYcV8Y0ZJSF6qgOpACytvwRFOneMS+kgJ2dmo1uZ0
# mHJ4PrrpkfkXPz1GKKlq5GOWGhmGjmGVK2QOdnFPiVBfpx7NRcBTjnSmvdgpruIh
# AOH5UqXjYw+a+lPslAz/kWtIfaJCs75j7DeFoG2123uOzpwEx/iBBkwWtJKnmp0d
# WOTmPvikvASNumewQobIzgKieNV1JSmGtN6DTv6e/pTNYY7VybDnZOOvsWQcyWok
# wcrkJ/rCFiu/rNwqZ5j9gn8rh3HzQfLihC2IjJQCls2QEazXZNKdVKIolO0b6MZZ
# laAsmSUbabyMTYjizxRCm3Gk2ErDJ3ctu4kuM2Oc9GazU9Zg9PNhyH/9lf3qGNBK
# ammsqrKfxZYzyd9ZD2BSpbbKvAeLmyAXU8A0plsxghqhMIIanQIBATB9MGkxCzAJ
# BgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGln
# aUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcgUlNBNDA5NiBTSEEzODQgMjAy
# MSBDQTECEAKkDd49j0pxthpYEOuQn7EwDQYJYIZIAWUDBAIBBQCgfDAQBgorBgEE
# AYI3AgEMMQIwADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3
# AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgyU2+JA/DX2xnk/lb
# 8prN25q1xx+aRkVUYG6kuOCTwZMwDQYJKoZIhvcNAQEBBQAEggIAVn4TdRVazag7
# QYE/yb2GVVFyuYkm6/UdLRZ1JH/IqaCoAZQIh4pz42K3m2yxLbKQR0xPMIZViq6e
# YaaK6SDLsDPYXmkfSkF+KmoQ/bvs/5Q0gep1KR4Lw3lGO7UvRuQI4I6Y50IqK5uf
# 9RKxYHqKcYQman200FqZbwJWxIYabBWhcyDmUf6+Wlmbo+mG3fgSuh+BAUP1WwIg
# B6t/BXwnXHvBj5diLwiJnzs0evpTVVEiimXI9+BT7t5W9JDGMhoFkjYC9ciyoVz1
# WIhfRA/dbN/rmVKXY0xicoSDcAgkcmdDFuyUzZ6ORn05LblzruDR2rqWxxpLnoRs
# YSxZrDL8hhttpY0TF//o5BItADiQMGjIWV/qkN+zgQ0Vn3MxoZd27l05iK7TyTt4
# w01JB0iAzRzmNeA9lm6lK6vK53q5fC0spDRi9O0aXmUW2707Q1iR8HQSTCz5cg7I
# dGGpexJMwSu+hbE2XSpn9CVVha5nIFzZQ9llwdGyWGtnNFXswit680aJopoT8P4v
# HeRrsvCGAiZO5WG5z2cFuwMVW7TJqo+D0r0az4ZfZbLstJT8WQss24Hfn9wnTvhO
# SDzpIIZ3EzGDSmjTpNFGkQ96nsNjXRhfWZoNP9V6mPn2wA3GaOXqSCO+Km3wxa0Y
# +09RVnHL3gxNpUkUS4Kxm4f7sbwfAHehghd3MIIXcwYKKwYBBAGCNwMDATGCF2Mw
# ghdfBgkqhkiG9w0BBwKgghdQMIIXTAIBAzEPMA0GCWCGSAFlAwQCAQUAMHgGCyqG
# SIb3DQEJEAEEoGkEZzBlAgEBBglghkgBhv1sBwEwMTANBglghkgBZQMEAgEFAAQg
# UMmDYl2/4/JOQx+rg9rPjHQrU7WarIhPCsTmORwjE/QCEQDJdvSi3gbcQ4M6LOu8
# Q8EkGA8yMDI2MDMwNTEyNDUwMlqgghM6MIIG7TCCBNWgAwIBAgIQCoDvGEuN8QWC
# 0cR2p5V0aDANBgkqhkiG9w0BAQsFADBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMO
# RGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGlt
# ZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0ExMB4XDTI1MDYwNDAwMDAw
# MFoXDTM2MDkwMzIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBTSEEyNTYgUlNBNDA5NiBUaW1l
# c3RhbXAgUmVzcG9uZGVyIDIwMjUgMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBANBGrC0Sxp7Q6q5gVrMrV7pvUf+GcAoB38o3zBlCMGMyqJnfFNZx+wvA
# 69HFTBdwbHwBSOeLpvPnZ8ZN+vo8dE2/pPvOx/Vj8TchTySA2R4QKpVD7dvNZh6w
# W2R6kSu9RJt/4QhguSssp3qome7MrxVyfQO9sMx6ZAWjFDYOzDi8SOhPUWlLnh00
# Cll8pjrUcCV3K3E0zz09ldQ//nBZZREr4h/GI6Dxb2UoyrN0ijtUDVHRXdmncOOM
# A3CoB/iUSROUINDT98oksouTMYFOnHoRh6+86Ltc5zjPKHW5KqCvpSduSwhwUmot
# uQhcg9tw2YD3w6ySSSu+3qU8DD+nigNJFmt6LAHvH3KSuNLoZLc1Hf2JNMVL4Q1O
# pbybpMe46YceNA0LfNsnqcnpJeItK/DhKbPxTTuGoX7wJNdoRORVbPR1VVnDuSeH
# VZlc4seAO+6d2sC26/PQPdP51ho1zBp+xUIZkpSFA8vWdoUoHLWnqWU3dCCyFG1r
# oSrgHjSHlq8xymLnjCbSLZ49kPmk8iyyizNDIXj//cOgrY7rlRyTlaCCfw7aSURO
# wnu7zER6EaJ+AliL7ojTdS5PWPsWeupWs7NpChUk555K096V1hE0yZIXe+giAwW0
# 0aHzrDchIc2bQhpp0IoKRR7YufAkprxMiXAJQ1XCmnCfgPf8+3mnAgMBAAGjggGV
# MIIBkTAMBgNVHRMBAf8EAjAAMB0GA1UdDgQWBBTkO/zyMe39/dfzkXFjGVBDz2GM
# 6DAfBgNVHSMEGDAWgBTvb1NK6eQGfHrK4pBW9i/USezLTjAOBgNVHQ8BAf8EBAMC
# B4AwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwgZUGCCsGAQUFBwEBBIGIMIGFMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wXQYIKwYBBQUHMAKG
# UWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRp
# bWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNydDBfBgNVHR8EWDBWMFSg
# UqBQhk5odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRU
# aW1lU3RhbXBpbmdSU0E0MDk2U0hBMjU2MjAyNUNBMS5jcmwwIAYDVR0gBBkwFzAI
# BgZngQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4ICAQBlKq3xHCcE
# ua5gQezRCESeY0ByIfjk9iJP2zWLpQq1b4URGnwWBdEZD9gBq9fNaNmFj6Eh8/Ym
# RDfxT7C0k8FUFqNh+tshgb4O6Lgjg8K8elC4+oWCqnU/ML9lFfim8/9yJmZSe2F8
# AQ/UdKFOtj7YMTmqPO9mzskgiC3QYIUP2S3HQvHG1FDu+WUqW4daIqToXFE/JQ/E
# ABgfZXLWU0ziTN6R3ygQBHMUBaB5bdrPbF6MRYs03h4obEMnxYOX8VBRKe1uNnzQ
# VTeLni2nHkX/QqvXnNb+YkDFkxUGtMTaiLR9wjxUxu2hECZpqyU1d0IbX6Wq8/gV
# utDojBIFeRlqAcuEVT0cKsb+zJNEsuEB7O7/cuvTQasnM9AWcIQfVjnzrvwiCZ85
# EE8LUkqRhoS3Y50OHgaY7T/lwd6UArb+BOVAkg2oOvol/DJgddJ35XTxfUlQ+8Hg
# gt8l2Yv7roancJIFcbojBcxlRcGG0LIhp6GvReQGgMgYxQbV1S3CrWqZzBt1R9xJ
# gKf47CdxVRd/ndUlQ05oxYy2zRWVFjF7mcr4C34Mj3ocCVccAvlKV9jEnstrniLv
# UxxVZE/rptb7IRE2lskKPIJgbaP5t2nGj/ULLi49xTcBZU8atufk+EMF/cWuiC7P
# OGT75qaL6vdCvHlshtjdNXOCIUjsarfNZzCCBrQwggScoAMCAQICEA3HrFcF/yGZ
# LkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNVBAoT
# DERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UE
# AxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoXDTM4
# MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGluZyBS
# U0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNqEY81
# FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fkHUil
# jNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EEbkC9
# +0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8NF5v
# p7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUUFREm
# DrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP9qh8
# SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKWxdCy
# QEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespYMQmU
# iote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrPV6/7
# umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+zoFp
# p4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGjggFd
# MIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK4pBW
# 9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8B
# Af8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBpMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKG
# NWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290
# RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQC
# MAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZxML2+
# C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97frPBtI
# j+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+NMvE
# uBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYAgwPy
# WLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA1WSj
# jf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+BFo+
# z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06VXxy
# KkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284NHNb
# oDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDezooIs
# 8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM9q7W
# P/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpSM9LH
# JmyrxaFtoza2zNaQ9k+5t1wwggWNMIIEdaADAgECAhAOmxiO+dAt5+/bUOIIQBha
# MA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lD
# ZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBaFw0zMTExMDky
# MzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAX
# BgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0
# ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAL/mkHNo
# 3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3EMB/zG6Q4FutW
# xpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKyunWZanMylNEQ
# RBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsFxl7sWxq868nP
# zaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU15zHL2pNe3I6P
# gNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJBMtfbBHMqbpEB
# fCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObURWBf3JFxGj2T3
# wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6nj3cAORFJYm2
# mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxBYKqxYxhElRp2
# Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5SUUd0viastkF1
# 3nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+xq4aLT8LWRV+d
# IPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIBNjAPBgNVHRMB
# Af8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwPTzAfBgNVHSME
# GDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMCAYYweQYIKwYB
# BQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20w
# QwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2Vy
# dEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0aHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDARBgNV
# HSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0NcVec4X6CjdBs9
# thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnovLbc47/T/gLn4
# offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65ZyoUi0mcudT6cG
# AxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFWjuyk1T3osdz9
# HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPFmCLBsln1VWvP
# J6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9ztwGpn1eqXiji
# uZQxggN8MIIDeAIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBp
# bmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJ
# YIZIAWUDBAIBBQCggdEwGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMBwGCSqG
# SIb3DQEJBTEPFw0yNjAzMDUxMjQ1MDJaMCsGCyqGSIb3DQEJEAIMMRwwGjAYMBYE
# FN1iMKyGCi0wa9o4sWh5UjAH+0F+MC8GCSqGSIb3DQEJBDEiBCDweXMKPJsYuOPd
# 0EGUjywI0YrbnsDjZ/aS+tr1X1mRoTA3BgsqhkiG9w0BCRACLzEoMCYwJDAiBCBK
# oD+iLNdchMVck4+CjmdrnK7Ksz/jbSaaozTxRhEKMzANBgkqhkiG9w0BAQEFAASC
# AgCH+d6dSLNt3KbUkmhoeQus+GfRvWiPoGoDnLPethusvf7aWhpbRignuvty2uwY
# ZAv5dhPPEeXdCUx4L8bppEgPGVjtJsHcm98dQo0JR6kN922QcFcAbmtSwUfB/A6D
# UnDKXHchgUYd9M+l5TjQ/n0qba1SQd3riXy7vPCfgjwU4K50AuZdKQ0dRiugZxRm
# XpwI43AvBxdYEJ/0qABhKSIe/QOwxro53F0+sUlA1p1t8BYp4vARYQZez/HT32ld
# v7djDsKkTgfd3wuDnC4b0dBYWK0Ek1Th8xQq64FyVrNoGGpfGmZZJCMId9miIdeH
# QkH1dIPMZw+aharfCuWgd6Pu3zmYUIVMUSB81i/7a0d58jYlyxZaCABsZ2rSPMRq
# V/A5hk8rGo2reSpTlmJBQt0fpOHDpIswCUaV5eD+Zwds4mQYH22j9ZFXcIXp2jiF
# 2dBAHbAuIQzZYGYWw/LOpvb2cHoOpBXZPrQJUDSfIO3lnUKMQ4nDlnDYXj7/EVan
# 8+HG9VDU5VJPeQGC+fBRc2jCh3w7/2+DnwtfAKrgM72jAMIRxVn1k9qUgv1o9t+B
# frT7+IwHd3PZxL0hlechiPFQmY+DQX/V/42Csd9oB/VunuKzeW0Y6JKgks77qibz
# 5QIC8GPrwh2lWyAhZFWy54EFR++e5ub2UxJD6YU6qFpgPg==
# SIG # End signature block
