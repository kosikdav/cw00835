. $ScriptPath\include-Script-StartLog-generic.ps1

Write-Log -String "AAD app name: $($script:AppName) ($($script:ClientId))"
Write-Log -String "Tenant Id: $($script:TenantId)"
Write-Log -String "Tenant name: $($script:TenantName)"
Write-Log -String "Certificate thumbprint: $($script:Thumbprint) (expires in $([math]::Round(($script:ClientCertificate.NotAfter - (Get-Date)).TotalDays)) days)"
