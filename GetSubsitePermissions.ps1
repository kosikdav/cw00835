#Parameters
$SiteURL = "https://cezdata.sharepoint.com/sites/df/uu"
$ReportOutput = "C:\Temp\SitePermissionRpt.csv"
 
#Connect to Site
Connect-PnPonline -Url $SiteURL -Interactive
 
#Get the web
$Web = Get-PnPWeb -Includes RoleAssignments
 
#Loop through each permission assigned and extract details
$PermissionData = @()
ForEach ($RoleAssignment in $Web.RoleAssignments)
{
    #Get the Permission Levels assigned and Member
    Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
     
    #Get the Permission Levels assigned
    $PermissionLevels = ($RoleAssignment.RoleDefinitionBindings | Select -ExpandProperty Name | Where {$_ -ne "Limited Access"}) -join ","
    $PermissionType = $RoleAssignment.Member.PrincipalType
 
    #Leave Principals with no Permissions
    If($PermissionLevels.Length -eq 0) {Continue}
     
    #Collect Permission Data
    $Permissions = New-Object PSObject
    $Permissions | Add-Member NoteProperty Name($RoleAssignment.Member.Title)
    $Permissions | Add-Member NoteProperty Type($PermissionType)
    $Permissions | Add-Member NoteProperty PermissionLevels($PermissionLevels)
    $PermissionData += $Permissions
}
 
$PermissionData
$PermissionData | Export-csv -path $ReportOutput -NoTypeInformation