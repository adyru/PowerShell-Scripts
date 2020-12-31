<#
Input file is searches2.csv with 2 columns
Search,SearchPhrase - one for the thing being looked for 
2nd is short version of this for report

Loops though files in $basefolders between 2 date ranges
And file filter $files

Searches files for matches and then groups then into folder output and match count

#>

$MatchedItems = @()
$GoBack = (get-Date).adddays(-1)
$Goforward = (get-Date).adddays(-14)
$Basefolders =  get-childitem "D:\Program Files\BlackBerry\UEM\Logs\"  |?{($_.mode -like "*d*") -and ($_.LastWriteTime -lt $GoBack) -and ($_.LastWriteTime -gt $GoForward)} # | select -First 1
ForEach($Basefolder in $Basefolders)
    {
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers();
        $files = "*_bp_*.txt"
        $Items = Get-ChildItem $Basefolder.fullname -Filter $files
        #$items
        $SearchTerms = import-csv Searches2.csv
        ForEach ($item in $items)
            {
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers();
            
            #$TotalItem = Get-Content $item.fullname
            ForEach ($SearchTerm in $SearchTerms)
                {
                #write-host $SearchTerm.Search
                write-host "Searching $item........"
                $LogParser = [System.IO.File]::OpenText($item.fullname)
                while($null -ne ($line = $LogParser.ReadLine())) {
                   If( $line -like $SearchTerm.Search)
                    {
                    $myObject = New-Object System.Object
                    $myObject | Add-Member -type NoteProperty -name Folder -Value $Basefolder
                    $myObject | Add-Member -type NoteProperty -name MatchedItem -Value $MatchCount.Value
                    $myObject | Add-Member -type NoteProperty -name Phrase -Value $SearchTerm.SearchPhrase
                    $MatchedItems += $myObject
                    }

                            }
                } 
            #$AllItems += $TotalItems

            }
            
    }
$MatchedItems.Count
#$MatchedItems
$MatchedItems | Group-Object folder,Phrase | select count,name
#$MatchedItems 
