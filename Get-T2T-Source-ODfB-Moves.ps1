
$Report = @()
$OutFile = "d:\temp\dst-odfb-moves.csv"
$srcodfb = Get-SPOCrossTenantUserContentMoveState -PartnerCrossTenantHostURL "https://cezdata-my.sharepoint.com/"
$srcodfb.count
foreach ($move in $srcodfb) {
    $Report += [PSCustomObject]@{
        SourceSiteURL = $move.SourceSiteURL
        TargetSiteURL = $move.TargetSiteURL
        MoveState = $move.MoveState
    }   
}
$Report | Export-Csv -Path $OutFile -NoTypeInformation -Encoding UTF8

