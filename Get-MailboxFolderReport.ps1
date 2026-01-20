#######################################################################################################################
# Set-MaiboxProperties
#######################################################################################################################
param(
    [Alias("Definitions","IniFile")][string]$VariableDefinitionFile,
	$Identity,
	$FolderScope = "All"
)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
. $ScriptPath\include-Script-StdStartBlock.ps1

#######################################################################################################################

$LogFolder			= "mbxmgmt"
$LogFilePrefix		= "mbx-folder-report"
$LogFileFreq		= "Y"

#######################################################################################################################

. $ScriptPath\include-Script-StdIncBlock.ps1

$LogFile = New-OutputFile -RootFolder $RLF -Folder $LogFolder -Prefix $LogFilePrefix -Freq $LogFileFreq -Ext "log"
$OutputFile = New-OutputFile -RootFolder $ROF -Folder $null -Prefix "mailbox-folder-report" -Suffix $Identity -Ext "csv"

Connect-EXOService -AppRegName $AppReg_EXO_MGMT -TTL 120

[array]$folderReport = @()
$folderStatistics = Get-MailboxFolderStatistics $Identity -FolderScope $Folderscope -ResultSize Unlimited

foreach ($folderStatistic in $folderStatistics) {
	$FolderSizeBytes = $null
	$encoding= [System.Text.Encoding]::GetEncoding("us-ascii")
	$nibbler= $encoding.GetBytes("0123456789ABCDEF")     
	$folderIdBytes = [Convert]::FromBase64String($folderStatistic.FolderId)
	$indexIdBytes = New-Object byte[] 48
	$indexIdIdx=0
	$folderIdBytes | Select-Object -Skip 23 -First 24 | ForEach-Object {
		$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -shr 4]
		$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -band 0xF]
	}
	$folderQuery = "folderid:$($encoding.GetString($indexIdBytes))"

	if ($folderStatistic.FolderSize) {
        $strg = $folderStatistic.FolderSize
        $FolderSizeBytes = ($strg.Substring($strg.IndexOf("(")+1,$strg.Length-$strg.IndexOf("(")-7)).Replace(",","")
    }

	if ( $folderStatistic.FolderAndSubfolderSize) {
		$strg = $folderStatistic.FolderAndSubfolderSize
		$FolderAndSubfolderSizeBytes = ($strg.Substring($strg.IndexOf("(")+1,$strg.Length-$strg.IndexOf("(")-7)).Replace(",","")
	}

	$folderReport += [PSCustomObject]@{
		Name		= $folderStatistic.Name
		Type		= $folderStatistic.FolderType
		FolderPath 	= $folderStatistic.FolderPath
		FolderQuery = $folderQuery
		FolderId 	= $folderStatistic.FolderId
		ParentFolderId = $folderStatistic.ParentFolderId
		ContentFolder = $folderStatistic.ContentFolder
		MailboxGuid	= $folderStatistic.ContentMailboxGuid
		CreationTime = $folderStatistic.CreationTime
		LastModifiedTime = $folderStatistic.LastModifiedTime
		ItemsInFolder = $folderStatistic.ItemsInFolder
		FolderSize = $folderStatistic.FolderSize
		FolderSizeBytes = $FolderSizeBytes
		Movable = $folderStatistic.Movable
		RecoverableItemsFolder = $folderStatistic.RecoverableItemsFolder
		ContainerClass = $folderStatistic.ContainerClass
		TargetQuota = $folderStatistic.TargetQuota
		StorageQuota = $folderStatistic.StorageQuota
		StorageWarningQuota = $folderStatistic.StorageWarningQuota
		VisibleItemsInFolder = $folderStatistic.VisibleItemsInFolder
		HiddenItemsInFolder = $folderStatistic.HiddenItemsInFolder
		DeletedItemsInFolder = $folderStatistic.DeletedItemsInFolder
		ItemsInFolderAndSubfolders = $folderStatistic.ItemsInFolderAndSubfolders
		DeletedItemsInFolderAndSubfolders = $folderStatistic.DeletedItemsInFolderAndSubfolders
		FolderAndSubfolderSize = $folderStatistic.FolderAndSubfolderSize	
		FolderAndSubfolderSizeBytes = $FolderAndSubfolderSizeBytes
		TopSubjectSize = $folderStatistic.TopSubjectSize
		TopSubjectCount = $folderStatistic.TopSubjectCount
		TopClientInfoCountForSubject = $folderStatistic.TopClientInfoCountForSubject
		LowLatencyContainerQuota = $folderStatistic.LowLatencyContainerQuota
		SearchFolder = $folderStatistic.SearchFolder
		Identity = $folderStatistic.Identity
		ConversationNamespace = $folderStatistic.ConversationNamespace
		WhenLabeled = $folderStatistic.WhenLabeled
		IsValid = $folderStatistic.IsValid
		ObjectState = $folderStatistic.ObjectState
	}
}

Export-Report -Report $folderReport -Path $OutputFile -SortProperty "FolderPath"
