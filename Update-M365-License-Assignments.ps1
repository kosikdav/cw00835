#######################################################################################################################
# Update-M365-License-Assignments
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "lic-mgmt-M365"
$LogFilePrefix		= "update-license-assignments"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile 	= New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"
$LogFileMin = New-OutputFile -RootFolder $ROF -Folder $LogFolder -Prefix $LogFilePrefix -Ext "log"

$AllowedDirectSKUs = @(
	"f30db892-07e9-47e9-837c-80727f46fd3d",
	"a403ebcc-fae0-4ca2-8c8c-7a907fd6c235",
	"3f9f06f5-3c31-472c-985f-62d9c10ec167",
	"606b54a9-78d8-4298-ad8b-df6ef4481c80"
)
#######################################################################################################################

. $IncFile_StdLogStartBlock

$SKU_DB = Import-CSVToHashDB -Path $DBFileLicensingInfoSKUs -KeyName "skuId"

Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
$UriResource = "users"
$UriSelect = "id,userPrincipalName,displayName,assignedLicenses"
$UriFilter = "assignedLicenses/`$count+ne+0&`$count=true"
$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource -Top 999 -Select $UriSelect -Filter $UriFilter
write-host $Uri -ForegroundColor Green
$AADUsers = Get-GraphOutputREST -Uri $Uri -AccessToken $AuthDB[$AppReg_LOG_READER].AccessToken -ContentType $ContentTypeJSON -ConsistencyLevel "eventual" -ProgressDots -Text "licensed AAD users"
write-Log "Licensed AAD users: $($AADUsers.Count)"

foreach ($User in $AADUsers) {
	Request-MSALToken -AppRegName $AppReg_LOG_READER -TTL 30
	$UriResource = "users/$($User.id)/licenseAssignmentStates"
	$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
	try {
		$Query = Invoke-RestMethod -Headers $AuthDB[$AppReg_LOG_READER].AuthHeaders -Uri $Uri -Method "GET" -ContentType $ContentTypeJSON
		$licenseAssignmentStates = $Query.Value
	}
	Catch {
		$licenseAssignmentStates = $null
	}
	write-host "$($User.userPrincipalName) $($licenseAssignmentStates.Count)" -ForegroundColor Green
	foreach ($state in $licenseAssignmentStates) {
		if (($null -eq $state.assignedByGroup) -and ( -not ($AllowedDirectSKUs -contains $state.skuid))) {
			Request-MSALToken -AppRegName $AppReg_USR_MGMT -TTL 30
			if ($SKU_DB.ContainsKey($state.skuid)) {
				$sku = $SKU_DB[$state.skuid]
			}
			$UriResource = "users/$($User.id)/assignLicense"
			$Uri = New-GraphUri -Version "v1.0" -Resource $UriResource
			$GraphBody = [PSCustomObject]@{
				addLicenses = @();
				removeLicenses = @($state.skuid)
			} | ConvertTo-Json
			Try{
				$ResultRemove = Invoke-RestMethod -Headers $AuthDB[$AppReg_USR_MGMT].AuthHeaders -Uri $Uri -Method "POST" -ContentType $ContentTypeJSON -Body $GraphBody
				$LogEntry = "$($User.userPrincipalName): $($state.skuid) $($sku.skupartnumber) $($sku.skuDisplayName) (state:$($state.state)) license removed"
				Write-Log $LogEntry
				Write-Log $LogEntry -AlternateLogfile $LogFileMin
			}
			Catch {
				write-log $_.Exception.Message -MessageType "Error"
			}
		}
	}
}

#######################################################################################################################

. $IncFile_StdLogEndBlock

# SIG # Begin signature block
# MIIpXgYJKoZIhvcNAQcCoIIpTzCCKUsCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAlSPu/a3g3oOUW
# XmTibM+Jt/5mM3K7l3W5Qx8+giqZqqCCDhMwggawMIIEmKADAgECAhAIrUCyYNKc
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
# AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgPVcedC2bPP8ATQmn
# h0LM0725e6AvDWaFaVN4gco4UegwDQYJKoZIhvcNAQEBBQAEggIAWo0UfQDroL/v
# Ae27acndymSxFCPBfT6n92ZYt7brFPHy63q2bjRJvXFwADW7E+f5q9BfwTVq7Y1E
# 0XhFSMQTuf3vk2wHMzCzxixZA20FHB+73j/oBEsqgO/t5LjoHhQGWa/Sg6MySchY
# NB9NF0HKWSV6GS/70aeVvw+aDM7isgsKihiliur94VVM4MchNVqpBvEd82Xf7FGD
# 6WwHFiXBfXGPC2wbsPkPzyN7Rgi92PuqVR1ajqTikbinDLuUr82KMtgI7Azt6LyM
# BKgI7FNU/tnQ9qjjAdDxVNfZP3mgJ6mRfr6wqj3d7JSJEywp+ovo5EuOBjNB5WQJ
# jtz/dzomj94XSBxXUNm4plIv6hfAnoMu/fsm2swyEF8xwjdtRTHKtKN4TotUCjvH
# qzo0YMHzfsPUtGZYDpnZCe/8aHtgaFzIMA7XFsr7cRKVf6XDf20WJSeWCaIqo2Ni
# b1+Xiu6WK5nrhUWvdMwW8YFq6CqvXU0CrIq9VxZhyDqPv9YjVAMWLC5zHC328f0+
# s+49kT1c/imvMRfdM+Qsk1yGU/MWl90v+eehBZfWokirdPb4NbAExMSe4nHQNcGi
# 4Ctdnk7YqIuGQ9NooOG3uyelSLQgL4/vCtoji8aVUdkirHhv61Uvr/KOFqHH2udz
# tLYS5eucDTAyTmOZ41AM3reYIBZi5Kyhghd3MIIXcwYKKwYBBAGCNwMDATGCF2Mw
# ghdfBgkqhkiG9w0BBwKgghdQMIIXTAIBAzEPMA0GCWCGSAFlAwQCAQUAMHgGCyqG
# SIb3DQEJEAEEoGkEZzBlAgEBBglghkgBhv1sBwEwMTANBglghkgBZQMEAgEFAAQg
# Sz6efEqufRbf5MmbsFH5iTvNpz/bxO3Cxnrv9GqM/5kCEQDKKyeC73TcObZCMXaO
# MmEIGA8yMDI2MDMxOTEwMzAwMlqgghM6MIIG7TCCBNWgAwIBAgIQCoDvGEuN8QWC
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
# SIb3DQEJBTEPFw0yNjAzMTkxMDMwMDJaMCsGCyqGSIb3DQEJEAIMMRwwGjAYMBYE
# FN1iMKyGCi0wa9o4sWh5UjAH+0F+MC8GCSqGSIb3DQEJBDEiBCDhSY29DKNahoBm
# U4c6KQSWyv5nZCJ1Bq6DY/tM6Mj4NDA3BgsqhkiG9w0BCRACLzEoMCYwJDAiBCBK
# oD+iLNdchMVck4+CjmdrnK7Ksz/jbSaaozTxRhEKMzANBgkqhkiG9w0BAQEFAASC
# AgCOHYG4HzxJ/tXz1EFRprIJa2cdvpVOC2TABURRyw9WrLtwKc4zDg9ddJW1MbNN
# L5DRs8le9kzLgQceJ3kqq7GpWa/AmKr3tiERx3gfgT1MSkvudVLX3LsZP4hRPpfV
# 1D3EupOcdsNcD4OQZc+CqeuRTX6fpWyVL4LH2jRoCnopuoszv1zsAkLnicr7Cex0
# +BuNj1dqZwO3i3b7Og1k92Zt6HTOO98AsqwIk4mBNmdDcnvzOJrZIUzCBNM2GMWL
# AGycYK9a0GTZToplKIWdJalzGn6ZQih8G0zoAdLqJs5cFKj1MEjp4/0yR1GdPoU8
# dO8rnsqHFsa9+TeyxTSpEjRNuk/BEgseRq8H4xpWPlZv7+OpNoiLTxqIqxsgblX4
# 7pCUiL8Z+0zNnzEvX8fBp+u6+N8kyu8ugRV4/aXGV8BVY4kJS0OFseEIupaYEmEc
# J5fotWpKm6kKw3U8wlx/0lqqeZKdErxpiefepcN0Z+/bnnSZttHNpANCevH5sb26
# hjPM/VcMsYLg+6WOv4h0K7xaL6CyCUkaPRtJpaP+VnJuWU7xO58/zQPCRBA9dExv
# 53FZVSCCw43lpc2y2E1Q3BDbn3wYfbtTO+yiPGX9zPgr39dtoHC06X6b71lyF1h9
# KX9WP2Y8IExCwBVjQXr62tkQuySomCuDx4GpXhDHei7j1Q==
# SIG # End signature block
