
<#
this looks up what process has an open connection to a IP
then outputs this to a csv
#>
#Change to ip you are after
$IP = "*10*"
$date = get-date -Format dd-MM-yyyy-HH-mm-ss
$CSVOut =  $PSScriptRoot + "\" + $Date + "-Process.csv"
#initialise array
$PIDMatches = @()
#Amend to how many ties to loop
foreach($i in 1..5)
    {
    Write-Host $i
    #how long to pause in each loop
    Start-Sleep -Seconds 1
    #here we look at what connections there are to an IP
    $connections = Get-NetTCPConnection |? {$_.RemoteAddress -like $IP}
        ForEach($connection in $connections)
            {
            #then we will look up the process that has that ip open
            $processes = get-process -id $connection.OwningProcess
            ForEach($process in $processes)
                {
                #for each of these will input to an array
                    # we input what is useful into this
                $process.ProcessName 
                $process.ProcessName, $process.ID 
                $PIDObject = New-Object System.Object
                $PIDObject | Add-Member -type NoteProperty -name ProcessName -Value $process.ProcessName
                $PIDObject | Add-Member -type NoteProperty -name PID  -Value $process.ID
                $PIDObject | Add-Member -type NoteProperty -name RemoteAddress -Value $connection.RemoteAddress
                $PIDObject | Add-Member -type NoteProperty -name RemotePort -Value $connection.RemotePort
                $PIDObject | Add-Member -type NoteProperty -name CreationTime -Value $connection.CreationTime 
                #add object to array
                $Script:PIDMatches +=  $PIDObject
                }
            }

    }
#output this to a csv and null
$PIDMatches | Export-Csv -NoClobber -NoTypeInformation -path $CSVOut
$PIDMatches = $null




