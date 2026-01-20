$Url = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
$Path = "d:\data\m365-licensing\ms-licensing-product-names-service-plans.csv"
Invoke-WebRequest $Url -OutFile $Path
