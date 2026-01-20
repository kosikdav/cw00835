$products = Get-MSCommerceProductPolicies -PolicyId AllowSelfServicePurchase
write-host "BEFORE:" -ForegroundColor Yellow
foreach ($product in $products){
    $prodStr = $product.ProductName + " (" + $product.ProductId + ")"
    write-host "$($prodStr.PadRight(75," "))" -NoNewline
    if ($product.PolicyValue -eq "Enabled") {
        write-host "Enabled" -foregroundcolor Red
    } else {
        write-host "Disabled" -ForegroundColor Green
    }
}
Write-Host
Write-Host "#######################################################################################"
Write-Host
$products = Get-MSCommerceProductPolicies -PolicyId AllowSelfServicePurchase | Where { $_.PolicyValue -eq "Enabled"}
foreach ($product in $products){
    #Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId $product.ProductId -Enabled $False
}
$products = Get-MSCommerceProductPolicies -PolicyId AllowSelfServicePurchase
write-host "AFTER:" -ForegroundColor Yellow
foreach ($product in $products){
    $prodStr = $product.ProductName + " (" + $product.ProductId + ")"
    write-host "$($prodStr.PadRight(75," "))" -NoNewline
    if ($product.PolicyValue -eq "Enabled") {
        write-host "Enabled" -foregroundcolor Red
    } else {
        write-host "Disabled" -ForegroundColor Green
    }
}
