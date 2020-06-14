<#
Takes input of records from a text file of names
Does look up writes to screen records
or that it cant be found on server

run command and then run teh below to find exception
$Error[0].Exception.GetType().FullName
System.ComponentModel.Win32Exception

#>

$Matches = @()

Function Global:DNS
    {
        try
            {$Script:Lookups = Resolve-DnsName $record -ErrorAction Stop}
   
   catch [System.ComponentModel.Win32Exception] 
            {Write-host -ForegroundColor red $record " was not found"}
    
    Finally
            {
               forEach ($lookup in $Lookups)
                {
                write-host $lookup.name,$lookup.IPAddress,$lookup.Type
                $DNSObject = New-Object System.Object
                $DNSObject | Add-Member -type NoteProperty -name Name -Value $lookup.name
                $DNSObject | Add-Member -type NoteProperty -name IP -Value $lookup.IPAddress
                $DNSObject | Add-Member -type NoteProperty -name Type -Value $lookup.Type
                #add object to array
                #Matches$Global:Matches +=  $DNSObject
                $Script:Matches = $Matches +  $DNSObject
                }
            }
    }
 

$date = get-date -Format dd-MM-yyyy-HH-mm-ss
$CSVOut =  $PSScriptRoot + "\" + $Date + "-DNS.csv"


$Records = Get-Content servers.txt
ForEach ($record in $Records)
    {
    DNS record
    }
$Matches | Export-Csv -NoClobber -NoTypeInformation -path $CSVOut
$Matches = $null