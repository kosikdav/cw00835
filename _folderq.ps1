
$mail = "pavel.kriz@cez.cz"
$folderQueries = @()
$folderStatistics = Get-MailboxFolderStatistics $mail
foreach ($folderStatistic in $folderStatistics) {
    $folderId = $folderStatistic.FolderId;
    $folderPath = $folderStatistic.FolderPath;
    if ($folderPath.contains("ARCHIV")) {
        Continue
    }
    $encoding= [System.Text.Encoding]::GetEncoding("us-ascii")
    $nibbler= $encoding.GetBytes("0123456789ABCDEF");
    $folderIdBytes = [Convert]::FromBase64String($folderId);
    $indexIdBytes = New-Object byte[] 48;
    $indexIdIdx=0;
    $folderIdBytes | select -skip 23 -First 24 | %{$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -shr 4];$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -band 0xF]}
    $folderQuery = "folderid:$($encoding.GetString($indexIdBytes))";

    $folderQueries += [PSCustomObject]@{
        FolderPath = $folderPath;
        FolderQuery = $folderQuery
    }
}

Write-Host "-----Exchange Folders----- $($mail)"
$folderQueries | Format-Table

Exit


$user = "martin.myska@cez.cz"
$searchname = "$($user)_Deleted_items"
New-ComplianceSearch $searchname -ExchangeLocation $user -ContentMatchQuery "folderid:473DC1797C261F46827D1E039EE67686000000FF782E0000"
Start-ComplianceSearch $searchname
Get-ComplianceSearch $searchname |fl status, *time*

<#
#Martin Myška, kpjm: myskamar, martin.myska@cez.cz
/Odstraněná pošta                                       folderid:473DC1797C261F46827D1E039EE67686000000FF782E0000
/Deletions                                              folderid:AC3096542C1BB74FB530CFCCAB16436B00000000A4210000
/DiscoveryHolds                                         folderid:A9DD0CC724FEA34B896C2845FE51AD800000315EBB860000
/Purges                                                 folderid:AC3096542C1BB74FB530CFCCAB16436B00000000A62D0000

#Jiří Gruber, kpjm: gruberjir, jiri.gruber@cez.cz	
#Pavel Kříž, kpjm: krizpav, pavel.kriz@cez.cz
#>