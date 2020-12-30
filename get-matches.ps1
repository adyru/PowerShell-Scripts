<#
takes search criteria from a file in $items
looks in the files from and outputs the match count
#>
$folder = "E:\Scripts"
$files = "*text*.txt"
$Searchfiles = $folder + "\" + $files
$Items = Get-ChildItem $Searchfiles
$SearchTerms = get-content Searches.txt
ForEach ($item in $items)
    {
    $Search = Get-Content $item
    Write-Host -ForegroundColor Green  $Item
    ForEach ($SearchTerm in $SearchTerms)
        {
        $MatchCount = [regex]::matches($search,$SearchTerm).count
        write-host "Total of " $SearchTerm "is " $MatchCount
        #$MatchCount
        } 
    }