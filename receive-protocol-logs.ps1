<#
This script searches the protocol logs of exchange servers
need to amend $searchtime to the UTC hour and date you wish to search
and amend $search to what you are looking for
it will output to a csv in the same directory as you ran the script
#>
# best to use server or IP as the search as the script is setup to pick up 10 lines after the search
# if you change it to sender recipit you will need to change the context number to pick up all the smtp conversation
$Search = "servername"
#time period you are interested in
#YearMonthDateHour - in UTC time
$SearchTime = "*2020060513*"
#Date and out vars
$date = get-date -Format hh-mm-dd-MMM-yyyy
$out = $date + "-receive-output.csv"
#Create Array
$Matches = @()

#Get exchane servers to search
$Servers = Get-ExchangeServer servers* | sort name

ForEach ($server in $servers)

    {
    #amend te path of servers protocl log location
    $path = "\\" + $server.name + "\h$\TransportRoles\Logs\Hub\ProtocolLog\SmtpReceive\" + $SearchTime
    $Files = Get-ChildItem $Path

     ForEach ($File in $Files)

        {
        write-host "Processing " $file
        #search file - conext numbers are lines above and below a match to return
        $lines = Get-Content $file | Select-String -Pattern  $search  -Context 0,9 

        ForEach ($line in $lines)

            {
            # remove leading characters and end one as not needed
            $line = $line -Replace "^> ", "" -Replace  "  ","" -Replace  "$,",""
            #Split on new line character
            $Relines  =$Line.Split([Environment]::NewLine)
            ForEach ($Reline in $Relines)
                {
                #if line is empty ignore
                If ([STRING]::IsNullOrWhitespace($Reline))
                    {
                    Write-Host "Ignoring Line"
                    }
                else{
                    #split lime into coresponding variables on ,
                    $DateTime,$ConnectorID,$Sessionid,$SequenceeNumber,$LocalEndpoint,$RemoteEndpoint,$Event,$Data,$Context = $ReLine.split(',')
                    Add these to a new object and add ones for server and logname
                        $myObject = New-Object System.Object
                        $myObject | Add-Member -type NoteProperty -name DateTime -Value $DateTime
                        $myObject | Add-Member -type NoteProperty -name ConnectorID -Value $ConnectorID
                        $myObject | Add-Member -type NoteProperty -name Sessionid -Value $Sessionid
                        $myObject | Add-Member -type NoteProperty -name SequenceeNumber -Value $SequenceeNumber
                        $myObject | Add-Member -type NoteProperty -name LocalEndpoint -Value $LocalEndpoint
                        $myObject | Add-Member -type NoteProperty -name RemoteEndpoint -Value $RemoteEndpoint
                        $myObject | Add-Member -type NoteProperty -name Event -Value $Event
                        $myObject | Add-Member -type NoteProperty -name Data -Value $Data
                        $myObject | Add-Member -type NoteProperty -name Context -Value $Context
                        $myObject | Add-Member -type NoteProperty -name server -Value $server.Name
                        $myObject | Add-Member -type NoteProperty -name File -Value $File.Name
                        #add object to array
                        $Matches += $myObject
                        }
                }

            }

        }

    }
#output array to csv
$Matches | Export-Csv -NoClobber -NoTypeInformation -path $Out

 

