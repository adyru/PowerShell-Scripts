<#
Takes input of records from a text file of names
Does look up writes to screen records
or that it cant be found on server

run command and then run teh below to find exception
$Error[0].Exception.GetType().FullName
System.ComponentModel.Win32Exception

#>
$Records = Get-Content servers.txt
ForEach ($record in $Records)
    {
    try
            {$Lookups = Resolve-DnsName $record -ErrorAction Stop}
   
   catch [System.ComponentModel.Win32Exception] 
            {Write-host -ForegroundColor red $record " was not found"}
    
    Finally
            {
               forEach ($lookup in $Lookups)
                {
                write-host $lookup.name,$lookup.IPAddress,$lookup.Type
                }
            }
        
    }