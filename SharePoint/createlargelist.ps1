$siteUrl = ""
$categoryChoices = @("Sales", "Marketing", "Operations", "Technology") -as [string[]]
$prefixList = @("UUG", "TPM", "ACB", "RRE", "CBB", "RNA", "PST") -as [string[]]

Connect-PnPOnline -Url $siteUrl

New-PnPList -Title "Large List" -Template GenericList
Add-PnPField -List "Large List" -DisplayName "Category" -Type Choice -Choices $categoryChoices

for ([int]$counter = 1; $counter -lt 6780; $counter++) {
    Add-PnPListItem -List "Large List" -Values @{
        "Title" = ("{0}{1:0000}" -f (Get-Random -InputObject $prefixList), $counter);
        "Category" = (Get-Random -InputObject $categoryChoices);
    }
}
