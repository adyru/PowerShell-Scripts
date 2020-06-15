

Function Script:BBUser
    {
    #voiud previous vvalues
    $BBUSer = $null
    $BBAppGroups = $null
    $BBAppGroup = $null
    $BBGroups = $null
    $BBGroup = $null

    #Garbage collection to stop memory going offpiste
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers();

    $BBUSer = get-aduser $user.Username -properties Sector,BU,GivenName,Surname,Title,Office
    $BBAppGroups = Get-QADMemberOf $user.Username | ? {$_.name -like "*app*"}
    $BBAppGroup =  $BBAppGroups.name -join ","
    $BBGroups = Get-QADMemberOf $user.Username | ? {$_.name -like "*bes*"}
    $BBGroup =  $BBGroups.name -join ","
            switch($BBUSer) 
                {
                # USer in Sector1
                {$BBUSer.Sector -eq "Sector1"}
	                {
                    write-host -foregroundcolor Red "MB is in Sector1"
                    $BBObject = New-Object System.Object
                    $BBObject | Add-Member -type NoteProperty -name 4x4 -Value $user.Username
                    $BBObject | Add-Member -type NoteProperty -name FirstName -Value $BBUSer.GivenName
                    $BBObject | Add-Member -type NoteProperty -name Surname -Value $BBUSer.Surname
                    $BBObject | Add-Member -type NoteProperty -name Postion -Value $BBUSer.Title
                    $BBObject | Add-Member -type NoteProperty -name Office -Value $BBUSer.Office
                    $BBObject | Add-Member -type NoteProperty -name Sector -Value $BBUSer.Sector
                    $BBObject | Add-Member -type NoteProperty -name BU -Value $BBUSer.BU
                    $BBObject | Add-Member -type NoteProperty -name OS -Value $USer."OS version"
                    $BBObject | Add-Member -type NoteProperty -name Model -Value $USer.Model
                    $BBObject | Add-Member -type NoteProperty -name Serial -Value $USer."Serial number"
                    $BBObject | Add-Member -type NoteProperty -name LastContact -Value $USer."Last contact"
                    $BBObject | Add-Member -type NoteProperty -name MDMAppGroups -Value $BBAppGroup
                    $BBObject | Add-Member -type NoteProperty -name MDMGroups -Value $BBGroup
                    #add object to array
                    $Script:AVBB +=  $BBObject
                    } 
                # USer in Sector2
                {$BBUSer.Sector -eq "Sector2"}
	                {
                    write-host -foregroundcolor Red "MB is in Sector2"
                    $BBObject = New-Object System.Object
                    $BBObject | Add-Member -type NoteProperty -name 4x4 -Value $user.Username
                    $BBObject | Add-Member -type NoteProperty -name FirstName -Value $BBUSer.GivenName
                    $BBObject | Add-Member -type NoteProperty -name Surname -Value $BBUSer.Surname
                    $BBObject | Add-Member -type NoteProperty -name Postion -Value $BBUSer.Title
                    $BBObject | Add-Member -type NoteProperty -name Office -Value $BBUSer.Office
                    $BBObject | Add-Member -type NoteProperty -name Sector -Value $BBUSer.Sector
                    $BBObject | Add-Member -type NoteProperty -name BU -Value $BBUSer.BU
                    $BBObject | Add-Member -type NoteProperty -name OS -Value $USer."OS version"
                    $BBObject | Add-Member -type NoteProperty -name Model -Value $USer.Model
                    $BBObject | Add-Member -type NoteProperty -name Serial -Value $USer."Serial number"
                    $BBObject | Add-Member -type NoteProperty -name LastContact -Value $USer."Last contact"
                    $BBObject | Add-Member -type NoteProperty -name MDMAppGroups -Value $BBAppGroup
                    $BBObject | Add-Member -type NoteProperty -name MDMGroups -Value $BBGroup
                    #add object to array
                    $Script:MTBB +=  $BBObject
                    
                    }
                # USer in Sector3
                {$BBUSer.Sector -eq "Sector3"}
	                {
                    write-host -foregroundcolor Red "MB is in Sector3 "
                    $BBObject = New-Object System.Object
                    $BBObject | Add-Member -type NoteProperty -name 4x4 -Value $user.Username
                    $BBObject | Add-Member -type NoteProperty -name FirstName -Value $BBUSer.GivenName
                    $BBObject | Add-Member -type NoteProperty -name Surname -Value $BBUSer.Surname
                    $BBObject | Add-Member -type NoteProperty -name Postion -Value $BBUSer.Title
                    $BBObject | Add-Member -type NoteProperty -name Office -Value $BBUSer.Office
                    $BBObject | Add-Member -type NoteProperty -name Sector -Value $BBUSer.Sector
                    $BBObject | Add-Member -type NoteProperty -name BU -Value $BBUSer.BU
                    $BBObject | Add-Member -type NoteProperty -name OS -Value $USer."OS version"
                    $BBObject | Add-Member -type NoteProperty -name Model -Value $USer.Model
                    $BBObject | Add-Member -type NoteProperty -name Serial -Value $USer."Serial number"
                    $BBObject | Add-Member -type NoteProperty -name LastContact -Value $USer."Last contact"
                    $BBObject | Add-Member -type NoteProperty -name MDMAppGroups -Value $BBAppGroup
                    $BBObject | Add-Member -type NoteProperty -name MDMGroups -Value $BBGroup
                    #add object to array
                    $Script:LABB  +=  $BBObject
                    }
                # USer in Sector4
                {$BBUSer.Sector -eq "Sector4"}
	                {
                    write-host -foregroundcolor Red "MB is in Sector4"
                    $BBObject = New-Object System.Object
                    $BBObject | Add-Member -type NoteProperty -name 4x4 -Value $user.Username
                    $BBObject | Add-Member -type NoteProperty -name FirstName -Value $BBUSer.GivenName
                    $BBObject | Add-Member -type NoteProperty -name Surname -Value $BBUSer.Surname
                    $BBObject | Add-Member -type NoteProperty -name Postion -Value $BBUSer.Title
                    $BBObject | Add-Member -type NoteProperty -name Office -Value $BBUSer.Office
                    $BBObject | Add-Member -type NoteProperty -name Sector -Value $BBUSer.Sector
                    $BBObject | Add-Member -type NoteProperty -name BU -Value $BBUSer.BU
                    $BBObject | Add-Member -type NoteProperty -name OS -Value $USer."OS version"
                    $BBObject | Add-Member -type NoteProperty -name Model -Value $USer.Model
                    $BBObject | Add-Member -type NoteProperty -name Serial -Value $USer."Serial number"
                    $BBObject | Add-Member -type NoteProperty -name LastContact -Value $USer."Last contact"
                    $BBObject | Add-Member -type NoteProperty -name MDMAppGroups -Value $BBAppGroup
                    $BBObject | Add-Member -type NoteProperty -name MDMGroups -Value $BBGroup
                    #add object to array
                    $Script:COBB +=  $BBObject
                    }
                # USer in Sector5
                {$BBUSer.Sector -eq "Sector5"}
	                {
                    write-host -foregroundcolor Red "MB is in Sector5"
                    $BBObject = New-Object System.Object
                    $BBObject | Add-Member -type NoteProperty -name 4x4 -Value $user.Username
                    $BBObject | Add-Member -type NoteProperty -name FirstName -Value $BBUSer.GivenName
                    $BBObject | Add-Member -type NoteProperty -name Surname -Value $BBUSer.Surname
                    $BBObject | Add-Member -type NoteProperty -name Postion -Value $BBUSer.Title
                    $BBObject | Add-Member -type NoteProperty -name Office -Value $BBUSer.Office
                    $BBObject | Add-Member -type NoteProperty -name Sector -Value $BBUSer.Sector
                    $BBObject | Add-Member -type NoteProperty -name BU -Value $BBUSer.BU
                    $BBObject | Add-Member -type NoteProperty -name OS -Value $USer."OS version"
                    $BBObject | Add-Member -type NoteProperty -name Model -Value $USer.Model
                    $BBObject | Add-Member -type NoteProperty -name Serial -Value $USer."Serial number"
                    $BBObject | Add-Member -type NoteProperty -name LastContact -Value $USer."Last contact"
                    $BBObject | Add-Member -type NoteProperty -name MDMAppGroups -Value $BBAppGroup
                    $BBObject | Add-Member -type NoteProperty -name MDMGroups -Value $BBGroup
                    #add object to array
                    $Script:CNBB +=  $BBObject
                    }

                default
                    {
                    write-host -foregroundcolor Red "MB is in Sector5"
                    $BBObject = New-Object System.Object
                    $BBObject | Add-Member -type NoteProperty -name 4x4 -Value $user.Username
                    $BBObject | Add-Member -type NoteProperty -name FirstName -Value $BBUSer.GivenName
                    $BBObject | Add-Member -type NoteProperty -name Surname -Value $BBUSer.Surname
                    $BBObject | Add-Member -type NoteProperty -name Postion -Value $BBUSer.Title
                    $BBObject | Add-Member -type NoteProperty -name Office -Value $BBUSer.Office
                    $BBObject | Add-Member -type NoteProperty -name Sector -Value $BBUSer.Sector
                    $BBObject | Add-Member -type NoteProperty -name BU -Value $BBUSer.BU
                    $BBObject | Add-Member -type NoteProperty -name OS -Value $USer."OS version"
                    $BBObject | Add-Member -type NoteProperty -name Model -Value $USer.Model
                    $BBObject | Add-Member -type NoteProperty -name Serial -Value $USer."Serial number"
                    $BBObject | Add-Member -type NoteProperty -name LastContact -Value $USer."Last contact"
                    $BBObject | Add-Member -type NoteProperty -name MDMAppGroups -Value $BBAppGroup
                    $BBObject | Add-Member -type NoteProperty -name MDMGroups -Value $BBGroup
                    #add object to array
                    $Script:UNBB +=  $BBObject
                    }
                }
        
    }

Function $Script:Get-BBGroups
    {

    }

$COBB = @()
$MTBB = @()
$AVBB = @()
$LABB = @()
$CNBB = @()
$UNBB = @()

#import BB users
$users = Import-Csv export.csv | select -First 100


# Format the date so that we can create different output files
$date = get-date -Format dd-MM-yyyy--hh-mm
$COBBCSVOut =  $date + "COBB-Blue-BB-Users.csv"
$MTBBCSVOut =  $date + "MTBB-Blue-BB-Users.csv"
$AVBBCSVOut =  $date + "AVBB-Blue-BB-Users.csv"
$LABBCSVOut =  $date + "LABB-Blue-BB-Users.csv"
$CNBBCSVOut =  $date + "CNBB-Blue-BB-Users.csv"
$UNBBCSVOut =  $date + "UNBB-Blue-BB-Users.csv"
ForEach ($user in $users)
    {
    BBUSer $user.Username
    }

$MTBB | Export-Csv -NoClobber -NoTypeInformation -path $MTBBCSVOut
$COBB| Export-Csv -NoClobber -NoTypeInformation -path $COBBCSVOut
$AVBB | Export-Csv -NoClobber -NoTypeInformation -path $AVBBCSVOut
$LABB | Export-Csv -NoClobber -NoTypeInformation -path $LABBCSVOut
$CNBB | Export-Csv -NoClobber -NoTypeInformation -path $CNBBCSVOut
$UNBB | Export-Csv -NoClobber -NoTypeInformation -path $UNBBCSVOut 
