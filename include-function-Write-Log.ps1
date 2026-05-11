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
