<#
################################################################################################
 2020-07-01

 Program Reason: 
 Backuping AIP UL Settings Labels and Policy and RMS Templates to files
 

# Zajisteni prav pro spusteni skriptu:
#################################################################################################

Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope LocalMachine


# Příklad jak se vygeneruje Password HASH do souboru
# Tento 
Read-Host "Enter Password" -AsSecureString | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File "D:\Password.txt"

# Příklad odhlášení
Disconnect-PSSession

#>

CLS
Clear-Host
Remove-Variable * -ErrorAction SilentlyContinue
$Error.clear()


<#################################################################################################
#                DEFINOVANI PROMMENNYCH PRO BEH PROGRAMU
#################################################################################################>
$datum = get-date -Format yyyy-MM-dd_HH-mm
# CESTA KDE JE SPUSTEN SKRIP
$scriptpath = Split-Path -parent $MyInvocation.MyCommand.Definition

$backuppath = $scriptpath

$backuppath = $backuppath + "\AIPBackup_$datum" 
write-host "Umístění zálohy: " $backuppath -ForegroundColor Yellow

<#################################################################################################
#                         Overeni zalogování do AZURE - pri New nebo SET hodnot  
#################################################################################################> 


# Když není adresář pro soubory zalohy, tak se vytvoří
if (!(Test-Path $backuppath)) {
New-Item -ItemType "directory" -Path $backuppath -Force -ErrorAction Stop
}
 
 
# 1 - ZAPISOVACI REZIM !!!

$jsemzalogovan = Get-PSSession | select name | FT
# Kontrola prihlašení do Azure

$hesloje = get-content "$scriptpath\Password.txt" -ErrorAction SilentlyContinue

<#
if ($hesloje -eq $null -and $jsemzalogovan.count -eq 0 ) {
# Read-Host "Enter Password" -AsSecureString | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File "D:\Password.txt"

#  PROHLASENI DO EOL SERVICE
$credential = New-Object System.Management.Automation.PSCredential ('$Username', $SecurePassword)
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session 
}
#>

. "D:\scripts-m365\cezdata\include-appreg-CEZDATA_EXO_MGMT.ps1"

Connect-IPPSSession
Connect-AipService

<#################################################################################################
# 
#                                      START Programu 
#   
#################################################################################################> 


#
#  Zalohovani MIP Sabon
#
$sablony = Get-Label | select Name, DisplayName, Disabled, Comment 

write-host "Export nalezených šablon:   `t`t" $sablony.count -foregroundColor Yellow 
$outcesta = "$backuppath\List Labels.bck"
$outcesta
Get-Label | select Name, DisplayName, Disabled, Comment | ft > $outcesta
 
write-host "Export Detailu nalezených šablon " -foregroundColor Yellow -BackgroundColor Black
$outcesta = "$backuppath\List Labels full details.bck"
$outcesta
Get-Label | fl > $outcesta -force




#
#  Zalohovani MIP Politik
#
$politik = Get-LabelPolicy | select name, CreatedBy, Enabled, Comment

write-host "Export nalezených politik  `t`t" $politik.count -foregroundColor Yellow 
$outcesta = "$backuppath\List Policies.bck"
$outcesta
Get-LabelPolicy | select name, CreatedBy, Enabled, Comment | Sort-Object -Descending | FT > $outcesta

write-host "Export Detailu nalezených šablon " -foregroundColor Yellow -BackgroundColor Black
$outcesta = "$backuppath\List Policies full details.bck"
$outcesta
Get-LabelPolicy | fl > $outcesta



#
#  Zalohovani RMS Sabon
#
$AIPExistujiciRMSTemplates = Get-AipServiceTemplate 


write-host "Export nalezených RMS šablon `t`t" $AIPExistujiciRMSTemplates.count -foregroundColor Yellow 
foreach ($item in $AIPExistujiciRMSTemplates)
{
# $item.LabelId  Nelze snadno načíst jméno Labelu RMS, takže se tu převádí.
$jmeno = $item.Names[1] -split("_")
$jmeno2 = $jmeno[0] -split(" -> ")
$jmeno = $jmeno2[1]

if ($jmeno.count -eq 0) {$jmeno = "No name defined in cs-CZ"}

write-host "Item: " $jmeno  -ForegroundColor Yellow

$fullname = $backuppath + "\"+$jmeno + "-" + $item.LabelId + ".xml"

Export-AipServiceTemplate -TemplateId $item.templateId.Guid -path $fullname -Force -ErrorAction Continue
}

