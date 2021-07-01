<#
This will convert a list of mailboxes
into MEU
#>


# First of we will say that this script will not run without .....
#Requires -Modules ActiveDirectory
#requires -Modules sqlserver

#Check Exchange PS commands are avaiable andexit if not
$ExchangePS = get-command "*MailboxDatabaseCopy*"
If(!($ExchangePS))
    {
    write-host -ForegroundColor Red "Exiting as exchange powershell module not found"
    exit
    }

# Below are the functions that we will be using in this script
Function Write-Log {
    <#
        . Notes
        =======================================
        v1  Created on: 25/05/2021
            Created by AU             
        =======================================
        . Description
        Params are Level and Message - it will use these as the output along with the date 
        Depending on how $OutputScreenInfo and  $OutputScreenNotInfo are set it will also output to screen
        You need to define the output file  $logfile  in the main script block
                write-log -level info -message "Something"

        #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $False)]
        [ValidateSet("INFO", "WARN", "ERROR", "FATAL", "DEBUG")]
        [String]
        $Level = "INFO",

        [Parameter(Mandatory = $True)]
        [string]
        $Message
    
    )
    # Check if we have a logfile defined and quit script if not with error to screen
    If (!($logfile)) {
        write-host -foregroundcolor red "Logfile not defind so exiting - create a variable called $logfile  and try again"
        exit

    }
    # Set these to $true or $false - switches on and off output to screen
    # One if for info eventss the other for anything but info
    $OutputScreenInfo = $false
    $OutputScreenNotInfo = $true
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    # Create the line from the timestamp, level and message
    $Line = "$Stamp $Level $Message"
    # Create seperate log file for errors
    $ErrorLog = $logfile -replace ".txt", "-errors.txt"
    #write-host "Output is $($outputscreen)"
    # Check if the level isnt info and then write to error log
    If ($level -ne "info") {
        Add-Content $ErrorLog  -Value $Line
        # Check if we want this outputted to screen, if so it is unexpected so error to screen
        If ($OutputScreenNotInfo -eq $true) {
            write-host -ForegroundColor Red $Message 
        }
    }
    # Write to log for info events
    Else {
        Add-Content $logfile -Value $Line
        # And if required output these to screen
        If ($OutputScreenInfo -eq $true) {
            write-host  $Message 
        }
    }

}

Function Export-File {
    <#
        . Notes
        =======================================
        v1  Created on: 13/05/2021
            Created by AU  
        V1.1 - 21/05/2021 Added in a check so that if files exists it exports with a random digit inserted           
        =======================================
        . Description
        Params are variable and filename 
        it then exports the variable to a csv if it isnt null
        If files exists adds in some random digits to makes sure there is an export
            export-file $var $varout
        #>
    [CmdletBinding()]
        
    Param(
        [Parameter(Mandatory = $false)]
        [Array]
        $OutVar,

        [Parameter(Mandatory = $True)]
        [string]
        $OutFile

    )
    Begin {
        $checkFile = test-path $outfile
    }
    Process {
        If ($checkFile -eq $true) {
            # if the files exists we will use  a random number into the filename 
            write-host -ForegroundColor Red "$($outfile) already exists"

            If (!($OutVar)) {
                write-host "Variable is null"
                return;
            }
            Else {
                $Random = Get-Random -Minimum 1 -Maximum 100
                $newOut = $outfile -replace ".csv", "$($Random).csv"
                $OutVar | Export-Csv -NoClobber -NoTypeInformation -path $newout -Encoding UTF8
                $OutcheckFile = test-path $outfile
                If ($OutcheckFile -eq $true)
                { write-host -ForegroundColor Green "Variable exported to $($newout)" }
                Else
                { write-host -ForegroundColor red "Variable not exported to $($newout)" }
            }
            return;
        }
        Else {
            If (!($OutVar)) {
                write-host "Variable is null"
                return;
            }
            Else {
                $OutVar | Export-Csv -NoClobber -NoTypeInformation -path $OutFile -Encoding UTF8
                $OutcheckFile = test-path $outfile
                If ($OutcheckFile -eq $true)
                { write-host -ForegroundColor Green "Variable exported to $($OutFile)" }
                Else
                { write-host -ForegroundColor red "Variable not exported to $($OutFile)" }
            }
        }
            
 

    }
    End {

    }
}



Function Nettestserver {
    <#
        .Description
        test-server server port.Params are server and port - accepts either common ports or numbers
        it then trys to resolve the DNS name and catches the error if cant 
        if it succeeds it goes on to test the connection to a port 
        returns values back to script block in $tcpclient -  Timeout for connection is $Timer in MS
        #>
    [CmdletBinding()]

    Param(
        [Parameter(Mandatory = $True)]
        [string]
        $servertest,

        [Parameter(Mandatory = $True)]
        [string]
        $Porttest
    )
    try {
        # Check DNS and just keep the first one
        $lookups = $null
        $Lookups = (Resolve-DnsName $servertest -DnsOnly -ErrorAction Stop).IP4Address
        $DNSCheck = $Lookups | Select-Object -First 1
    }
    #Catch the error in the DNS record doesnt exist
    catch [System.ComponentModel.Win32Exception] {
        Write-host -ForegroundColor red $servertest " was not found"
        exit
    }
    
    # Null out array
    If ([STRING]::IsNullOrWhitespace($DNSCheck)) {
    }
    Else {
        # If it is a numerical port do a check
        if ($porttest -match "^\d{1,3}") {
            try {
                $tcpclient = New-Object System.Net.Sockets.TCPClient
                $Timer = 1500
                $StartConnection = $tcpclient.BeginConnect($servertest, $PortTest, $null, $null)
                $wait = $StartConnection.AsyncWaitHandle.WaitOne($timer, $false)
                return  Write-Output -NoEnumerate  $tcpclient
            }

            catch {
                return  Write-Output -NoEnumerate  $tcpclient
            }
        }
                    
        Else {
            write-host -ForegroundColor Red "You have entered and incorrect port $($porttest) - it needs to either be a number"
        }
    }
}







######################
# Start of Variables #
######################

# Static DC to use - needs to be FQDN this will be checked later on
$DC = "BDCADOMR11.blue01.Babcockgroup.co.uk"

#Arrays
$NotQCSUser = @()
$NotQCSMailbox = @()
$QCSMailboxUser = @()
$MailUser = @()
$ErrorMEU = @()
$NoRecipient = @()
$NotMB = @()
$NoinAD = @()
$MBNotDelete = @()

#Get the date for Output
$date = get-date -Format dd-MM-yyyy--hh-mm


#Create new folder for output files if it doesnt exist using todays date
$FolderDate = Get-Date -Format dd-MM-yyyy--hh-mm
$newFolder = "$($PSScriptRoot)\output\$($FolderDate)"
$TestFolder = test-path $newFolder
If ($testfolder -eq $false) { New-Item -Path $NewFolder -ItemType directory }
$OutputFolder = $NewFolder

#Output Files that will use to check what has gone on
$NotQCSUserOut = "$OutputFolder\NotQCSUser-$($date).csv"
$NotQCSMailboxOut = "$OutputFolder\NotQCSMailbox-$($date).csv"
$QCSMailboxUserOut = "$OutputFolder\QCSMailboxUser-$($date).csv"
$MailUserOut = "$OutputFolder\MailUser-$($date).csv"
$ErrorMEUOut = "$OutputFolder\MailEU-Error-$($date).csv"
$NoRecipientOut = "$OutputFolder\Recipient-Error-$($date).csv"
$NotMBOut = "$OutputFolder\Recipient-Not-MB-$($date).csv"
$NotinADOut = "$OutputFolder\Recipient-Not-in-AD-$($date).csv"
$MBNotDeleteOut = "$OutputFolder\MB-Not_deleted-$($date).csv"
$NoRecipientOut = "$OutputFolder\No-recipient-$($date).csv"

#Need to create a logfile for the write-log function as the script will error out without that
$logfile = "$OutputFolder\Convert-MEU-Log-$($date).txt"

#SQL
$SQLServer = "bdcasqlr241ARS.blue01.babcockgroup.co.uk"
$Database = "AzureGuestReports"
$table = "QCSMigration"

#Counters to see where we are
[int]$USerProcessedCounter = 0
[int]$NotQCSUserCounter = 0
[int]$NotQCSMailboxCounter = 0
[int]$QCSMailboxUserCounter = 0
[int]$MailUserCounter = 0
[int]$ErrorMEUCounter = 0
[int]$NoRecipientCounter = 0
[int]$NotMBCounter = 0
[int]$NoinADCounter = 0
[int]$MBNotDeleteCounter = 0

#Inputfiles
$UsersFile = "users.txt"

# We are going to do a garbage collection every 2 mins so 
# need to kick of a timer
#Start Stopwatch
$sw = [diagnostics.stopwatch]::StartNew()

#Email to remove
$EMailDomain = "*@blue01.babcockgroup.co.uk"

###################
# End of Vaiables #
###################




#DC Connectivity Check - Exits on failure, user is given 10 secs to check DC
#write-host "Domain Controller to use is  $($DC)"
#Start-Sleep -Seconds 10

do { $answer = Read-Host "Domain Controller to use is  $($DC) - yes or no" }
until ("Yes", "No", "Y", "N" -contains $answer) write-host "$answer"
If ($answer -eq 'Yes' -or $answer -eq 'Y' -or $answer -eq 'yes' -or $answer -eq 'y') {
    write-host "Pressing On"
}
Else {
    write-host "Exiting script"
    exit
}


Write-Log -Message "Setting Domain Controller to  $($DC )" #write-host
Write-Log -Message "Checking connectivity to Domain Controller $($DC )" #write-host

# Check using the nettestserver function 
$CheckCOnnection = Nettestserver $DC 135
$CheckCOnnection

#$CheckCOnnection = test-netconnection $DC -port 135
# Check what the status of the connection is
If ($CheckCOnnection.Connected -eq $false) {
    # It failed - report and exit script
    write-host -ForegroundColor red "Connection is $($CheckCOnnection.Connected) so exiting script, amend DC variable and re run" 
    Write-Log -Message "Connectivity to Domain Controller $($DC ) is $($CheckCOnnection.Connected ) so exiting script " #write-host -Level "Fatal"
    exit
}
Else {
    #All good so carry on and report to screen
    write-host -ForegroundColor Green "Connectivity to Domain Controller $($DC ) Succdded - check connection is  $($CheckCOnnection.Connected ) "
    Write-Log -Message "Connectivity to Domain Controller $($DC ) Succdded - check connection is  $($CheckCOnnection.Connected ) " #write-host
}


#Check SQL Connection 
$conn = New-Object System.Data.SqlClient.SqlConnection                                      
$conn.ConnectionString = "Server=$SQLServer;Database=$Database;Integrated Security=True;"                                                                        
$conn.Open()
IF($conn.State -ne "Open")
    {
    Write-Log -Message "Failed to connect to database $($Database) on server $($SQLServer) so exiting" -Level "Error"
    exit
    }
Else
    {
    Write-Log -Message "Connected to database $($Database) on server $($SQLServer)"
    }

#Check import file and import users if present
Write-Log -Message "Checking $($UsersFile ) is available"
$CheckUserImport = Test-Path $UsersFile
If ($CheckUserImport -eq $false) {
    #Input file doesnt exist so exit script
    #write-host "$($UsersFile) doesnt exist so exiting script, amend $UsersFile variable and re run" -ForegroundColor red
    Write-Log -Message "$($UsersFile) doesnt exist so exiting script" #write-host -Level "Fatal"
    exit
}
Else {
    #Input found so carrying on
    Write-Log -Message "$($UsersFile) exists" #write-host 
}

# Import the users file
$USers = get-content $UsersFile # | select -First 40

#Report on the number of users to process
$UsersCount = $USers | Measure-Object
write-host "$($UsersCount.Count) to be processed"
Write-Log -Message "$($UsersCount) to be processed" #write-host 


#here we start the main loop
ForEach ($user in $users) {
    #Do garbage collection every couple of minutes to stop memory going off piste
    # so check if it is around 2 mins
    if ( $Sw.Elapsed.minutes -eq 2) {
        # it is over 2 mins so start garbage collection
        Write-Log -Message "Doing Garbage Collection after $($Sw.Elapsed.minutes ) minutes"  #write-host 
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers();
        #Reset timer by stopping and starting a new one
        $Sw.Stop()
        $sw = [diagnostics.stopwatch]::StartNew()

    }

    #Check where we are and report to screen
    $USerProcessedCounter ++
    Write-Log -Message "Processing $($USerProcessedCounter) of $($UsersCount.Count) mailboxes" #write-host 
    Write-host -ForegroundColor Green "Processing $($USerProcessedCounter) of $($UsersCount.Count) mailboxes" 
    # Check recipient type
    $RecipientType = get-recipient $user -erroraction silentlycontinue -DomainController $DC
    #check recipient exists and report if it cant be found
    If ([STRING]::IsNullOrWhitespace($RecipientType)) {
        # no match so log this  and move on  to next user
        #write-host -ForegroundColor Red "Cannot find a recipient matching $($user)"
        Write-Log -Message "Cannot find a recipient matching $($user)" #write-host -Level "Error"
        $NoRecipientObj = New-Object System.Object
        $NoRecipientObj | Add-Member -type NoteProperty -name Recipient -Value $User
        $NoRecipient += $NoRecipientObj
        $NoRecipientCounter++
    }
    # check if recipient is a mailbox
    Elseif ($RecipientType.RecipientType -ne "UserMailbox") {
        # recipeint isnt a mailbox so log this and move on  to next user
        #write-host -ForegroundColor Red "Recipient is not a mailbox - recipient type is matching $($RecipientType.RecipientType)"
        Write-Log -Message "Recipient is not a mailbox - recipient type is matching $($RecipientType.RecipientType)" #write-host -Level "Error"
        $NotMBObj = New-Object System.Object
        $NotMBObj | Add-Member -type NoteProperty -name Recipient -Value $User
        $NotMBObj | Add-Member -type NoteProperty -name Type -Value $RecipientType.RecipientType 
        $NotMB += $NotMBObj
        $NotMBCounter++
    }
    Else
    { # Mailbox found so we will process further
        Write-Log -Message "$($User) is $($RecipientType.RecipientType)" #write-host
        Write-Log -Message "Searching for $($User) in AD" #write-host  
        $ADUser = Get-ADUser -Filter 'mail -eq $user' -properties * -server $DC  
        # We cant find AD object with that email address
        If ([STRING]::IsNullOrWhitespace($ADUser)) {
            #write-host -ForegroundColor Red  -Message "Found $($ADUser)in AD"
            Write-Log -Message "$($User) not found in AD" #write-host -Level "Error"
            $NoinADObj = New-Object System.Object
            $NoinADObj  | Add-Member -type NoteProperty -name Recipient -Value $User
            $NoinAD += $NoinADObj 
            $NoinADCounter++
        }
        # USer found so we will continue with processing
        Else {
            Write-Log -Message "Found $($ADUser.samaccountname)in AD" #write-host  
            #If($ADUser.DistinguishedName -like "*OU=Quest Collaboration Services Objects,OU=QCS*")
            # Check if the AD User is in the QCS OU, has a 19 character samaccountname and also has a blank UPN
            If ($ADUser.DistinguishedName -like "*OU=Quest Collaboration Services Objects,OU=QCS*" -and $ADUser.samaccountname -match "\w{19}" -and ([STRING]::IsNullOrWhitespace($ADUser.UserPrincipalName))) {
                #write-host -ForegroundColor Green "Users is okay"
                Write-Log -Message "$($ADUser.Samaccountname) is in the QCS OU" #write-host 
                # Get the mailbox objects attributes including email addresses
                $MB = get-mailbox  $user -DomainController $DC | Select-Object @{L = "NewEmailAddresses"; E = { $_.EmailAddresses } }, *
                $JoinEmailAddresses = $MB.EmailAddresses -join ","
                #write-host "Emailaddresses are $($JoinEmailAddresses)"
                # Check if the Mailbox is on the QCS Exhange Database
                If ($MB.Database -like "*QCS*") {
                    #write-host -ForegroundColor Green "Database is $($MB.Database)"
                    # First put all the attributes found in an array in case we need to backout
                    Write-Log -Message "$($MB.PrimarySmtpAddress) Database is $($MB.Database)" #write-host 
                    $QCSMailboxUserObj = New-Object System.Object
                    $QCSMailboxUserObj | Add-Member -type NoteProperty -name samaccountname -Value $ADUser.samaccountname
                    $QCSMailboxUserObj | Add-Member -type NoteProperty -name DN -Value $ADUser.DistinguishedName
                    $QCSMailboxUserObj | Add-Member -type NoteProperty -name Database -Value $MB.Database
                    $QCSMailboxUserObj  | Add-Member -type NoteProperty -name PrimaryEmailAddress -Value $CheckMailUser.PrimarySmtpAddress.Address
                    $QCSMailboxUserObj  | Add-Member -type NoteProperty -name Alias -Value $MB.Alias
                    $QCSMailboxUserObj  | Add-Member -type NoteProperty -name Targetaddress -Value $ADUser.targetAddress 
                    $QCSMailboxUserObj  | Add-Member -type NoteProperty -name EmailAddresses -Value $JoinEmailAddresses 
                    $QCSMailboxUserObj  | Add-Member -type NoteProperty -name LegacyDN -Value $ADUser.LegacyExchangeDN
                    $QCSMailboxUserObj | Add-Member -type NoteProperty -name ExchangeGUID -Value $MB.ExchangeGuid
                    $QCSMailboxUser += $QCSMailboxUserObj
                    $QCSMailboxUserCounter++
                    $MB.PrimarySmtpAddress
                    Write-Log -Message "Disabling mailbox$($MB.PrimarySmtpAddress) on $($DC) " #write-host
                    # First thing we need to do is disable the mailbox - so remove it from the AD object so we can convert it to a MEU
                    Disable-Mailbox -Identity $MB.PrimarySmtpAddress -DomainController $DC -confirm:$false
                    # wait 1 sec for this to replicate
                    Start-Sleep -seconds 1
                    #Then check to see if this has worked
                    $checkMB = get-mailbox $MB.PrimarySmtpAddress  -DomainController $DC  -ErrorAction silentlycontinue
                    # If there is no such mailbox then we process the object further
                    If ([STRING]::IsNullOrWhitespace($checkMB)) {
                        Write-Log -Message "Mailbox  $($MB.PrimarySmtpAddress) on $($DC) doesnt exist" #write-host
                        Write-Log -Message "Enabling Mailuser $($adUser.samaccountname) on $($DC) with External Address $($ADUser.targetAddress)" #write-host
                        # First we need to enable the AD Object as an MEU - confirm needs to be removed
                        Enable-MailUser -Identity $adUser.samaccountname -ExternalEmailAddress $ADUser.targetAddress -alias $MB.alias -DomainController $DC -confirm:$false
                        Set-MailUser -Identity $adUser.samaccountname -EmailAddressPolicyEnabled $false -DomainController $DC
                        #Set-MailUser -Identity $adUser.samaccountname  -EmailAddresses @{remove=$EmailToRemove}
                        # wait 1 sec for this to replicate
                        Start-sleep -seconds 1
                        #Then check to see if this has worked
                        $CheckMailUser = Get-mailuser -Identity $adUser.samaccountname -DomainController $DC -ErrorAction silentlycontinue
                        # It hasnt so we need to take note of that and save the details for later use
                        If ([STRING]::IsNullOrWhitespace($CheckMailUser)) {
                            write-host -ForegroundColor red "Error finding mailuser $($adUser.samaccountname)"
                            Write-Log -Message "Cannot find MEU $($ADUser.Samaccountname)" #write-host -Level "Error"
                            $ErrorMEUObj = New-Object System.Object
                            $ErrorMEUObj | Add-Member -type NoteProperty -name samaccountname -Value $ADUser.samaccountname
                            $ErrorMEUObj | Add-Member -type NoteProperty -name DN -Value $ADUser.DistinguishedName
                            $ErrorMEUObj | Add-Member -type NoteProperty -name Database -Value $MB.Database
                            $ErrorMEUObj  | Add-Member -type NoteProperty -name Alias -Value $MB.Alias
                            $ErrorMEUObj  | Add-Member -type NoteProperty -name PrimaryEmailAddress -Value $MB.PrimarySmtpAddress
                            $ErrorMEUObj  | Add-Member -type NoteProperty -name Targetaddress -Value $ADUser.targetAddress 
                            $ErrorMEUObj  | Add-Member -type NoteProperty -name EmailAddresses -Value $JoinEmailAddresses 
                            $ErrorMEUObj  | Add-Member -type NoteProperty -name LegacyDN -Value $ADUser.LegacyExchangeDN
                            $ErrorMEUObj | Add-Member -type NoteProperty -name ExchangeGUID -Value $MB.ExchangeGuid
                            $ErrorMEU += $ErrorMEUObj
                            $ErrorMEUCounter++

                            # Add to SQL Table for reporting
                            [int]$Random = Get-Random -Maximum 100000 -Minimum 100
                            $Random
                            $SQLQuery  = "INSERT INTO QCSMigration  (Samaccountname,DistinguishedName,ExchangeDatabase,PrimarySmtpAddress,Alias,targetAddress,EmailAddresses,LegacyExchangeDN,ExchangeGuid,RandomNumber,RecordAdded,Succeeded,FailureReason) VALUES `
                                                ('$($ADUser.samaccountname)','$($ADUser.DistinguishedName)','$($MB.Database)','$($CheckMailUser.PrimarySmtpAddress)' `
                                                ,'$($MB.Alias)','$($ADUser.targetAddress)','$($JoinEmailAddresses)','$($ADUser.LegacyExchangeDN)'`
                                                ,'$($MB.ExchangeGuid)','$($Random)','$((get-date).ToString('yyyy-MM-dd'))','False','MEU Samaccountname Not Found'`
                                                )"
                                    Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $database  -Query $SQLQuery 

                        }
                        # MEU found so we will not continue processing
                        Else {
                            Write-Log -Message "Found MEU $($ADUser.Samaccountname)" #write-host 
                            Write-Log -Message "Setting Mailuser $($adUser.samaccountname) on $($DC) with External Address $($ADUser.targetAddress)" #write-host
                            # We will now add the rest of the Mail attributes to the object
                            Set-MailUser -Identity $adUser.samaccountname -EmailAddressPolicyEnabled $false -ExternalEmailAddress $ADUser.targetAddress  -EmailAddresses $MB.EMAILADDRESSES -ExchangeGuid $MB.ExchangeGuid -DomainController $DC -WindowsEmailAddress $MB.PrimarySmtpAddress # -erroraction silentlycontinue
                            # wait 1 sec for this to replicate
                            Start-sleep -seconds 1
                            #Then check to see if this has worked via looking up against the ExchangeGUID
                            $CheckMailUser2 = Get-mailuser -Identity $adUser.samaccountname -DomainController $DC  | Where-Object { $_.ExchangeGuid -eq $MB.ExchangeGuid }
                            # MEU with that GUID cant be found - so we will report on this
                            If ([STRING]::IsNullOrWhitespace($CheckMailUser2)) {
                                write-host -ForegroundColor red "Error finding mailuser $($adUser.samaccountname) with ExchangeGUID $($MB.ExchangeGuid)"
                                Write-Log -Message "Error finding mailuser $($adUser.samaccountname) with ExchangeGUID $($MB.ExchangeGuid)" #write-host -Level "Error"
                                $ErrorMEUObj = New-Object System.Object
                                $ErrorMEUObj | Add-Member -type NoteProperty -name samaccountname -Value $ADUser.samaccountname
                                $ErrorMEUObj | Add-Member -type NoteProperty -name DN -Value $ADUser.DistinguishedName
                                $ErrorMEUObj | Add-Member -type NoteProperty -name Database -Value $MB.Database
                                $ErrorMEUObj  | Add-Member -type NoteProperty -name PrimaryEmailAddress -Value $CheckMailUser.PrimarySmtpAddress
                                $ErrorMEUObj   | Add-Member -type NoteProperty -name Alias -Value $MB.Alias
                                $ErrorMEUObj  | Add-Member -type NoteProperty -name Targetaddress -Value $ADUser.targetAddress 
                                $ErrorMEUObj  | Add-Member -type NoteProperty -name EmailAddresses -Value $JoinEmailAddresses
                                $ErrorMEUObj  | Add-Member -type NoteProperty -name LegacyDN -Value $ADUser.LegacyExchangeDN
                                $ErrorMEUObj | Add-Member -type NoteProperty -name ExchangeGUID -Value $MB.ExchangeGuid
                                $ErrorMEU += $ErrorMEUObj

                                # Add to SQL Table for reporting
                                [int]$Random = Get-Random -Maximum 100000 -Minimum 100
                                $Random
                                $SQLQuery  = "INSERT INTO QCSMigration  (Samaccountname,DistinguishedName,ExchangeDatabase,PrimarySmtpAddress,Alias,targetAddress,EmailAddresses,LegacyExchangeDN,ExchangeGuid,RandomNumber,RecordAdded,Succeeded,FailureReason) VALUES `
                                                ('$($ADUser.samaccountname)','$($ADUser.DistinguishedName)','$($MB.Database)','$($CheckMailUser.PrimarySmtpAddress)' `
                                                ,'$($MB.Alias)','$($ADUser.targetAddress)','$($JoinEmailAddresses)','$($ADUser.LegacyExchangeDN)'`
                                                ,'$($MB.ExchangeGuid)','$($Random)','$((get-date).ToString('yyyy-MM-dd'))','False','ExchangeGUID Not Found'`
                                                )"
                                    Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $database  -Query $SQLQuery 
                            }
                            #MEU found with the same GUID so we will progress to the last stage
                            Else {
                                Write-Log -Message "Found mailuser $($adUser.samaccountname) with ExchangeGUID $($MB.ExchangeGuid)" #write-host 
                                Write-Log -Message "Setting AD User $($adUser.samaccountname) on $($DC) with legacyDN $($ADUser.legacyExchangeDN)" #write-host -erroraction silentlycontinue
                                # Change the legacyDN to the same value as the Mailbox had to stop there being NDRs when users email via Outlook
                                Set-ADUser $adUser.samaccountname -Replace @{legacyExchangeDN = $ADUser.legacyExchangeDN } -server $DC -erroraction SilentlyContinue
                                Start-sleep -seconds 1
                                #Then check to see if this has worked via looking up against the LegacyDN
                                $CheckLegacyDN = get-aduser $adUser.samaccountname -Properties * -server $DC  | Where-Object { $_.legacyExchangeDN -eq $ADUser.legacyExchangeDN }
                                # We cant find a user with that legacyDn so take note of this
                                If ([STRING]::IsNullOrWhitespace($CheckLegacyDN)) {
                                    write-host -ForegroundColor red "Error finding AD User $($adUser.samaccountname) with legacyExchangeDN $($ADUser.legacyExchangeDN)"
                                    Write-Log -Message "Error finding mailuser $($adUser.samaccountname) with legacyExchangeDN $($ADUser.legacyExchangeDN)" #write-host -Level "Error"
                                    $ErrorMEUObj = New-Object System.Object
                                    $ErrorMEUObj | Add-Member -type NoteProperty -name samaccountname -Value $ADUser.samaccountname
                                    $ErrorMEUObj | Add-Member -type NoteProperty -name DN -Value $ADUser.DistinguishedName
                                    $ErrorMEUObj | Add-Member -type NoteProperty -name Database -Value $MB.Database
                                    $ErrorMEUObj  | Add-Member -type NoteProperty -name PrimaryEmailAddress -Value $CheckMailUser.PrimarySmtpAddress
                                    $ErrorMEUObj  | Add-Member -type NoteProperty -name Alias -Value $MB.Alias
                                    $ErrorMEUObj  | Add-Member -type NoteProperty -name Targetaddress -Value $ADUser.targetAddress 
                                    $ErrorMEUObj  | Add-Member -type NoteProperty -name EmailAddresses -Value $JoinEmailAddresses 
                                    $ErrorMEUObj  | Add-Member -type NoteProperty -name LegacyDN -Value $ADUser.LegacyExchangeDN
                                    $ErrorMEUObj | Add-Member -type NoteProperty -name ExchangeGUID -Value $MB.ExchangeGuid
                                    $ErrorMEU += $ErrorMEUObj
                                    
                                    # Add to SQL Table for reporting
                                    [int]$Random = Get-Random -Maximum 100000 -Minimum 100
                                    $Random
                                    $SQLQuery  = "INSERT INTO QCSMigration  (Samaccountname,DistinguishedName,ExchangeDatabase,PrimarySmtpAddress,Alias,targetAddress,EmailAddresses,LegacyExchangeDN,ExchangeGuid,RandomNumber,RecordAdded,Succeeded,FailureReason) VALUES `
                                                ('$($ADUser.samaccountname)','$($ADUser.DistinguishedName)','$($MB.Database)','$($CheckMailUser.PrimarySmtpAddress)' `
                                                ,'$($MB.Alias)','$($ADUser.targetAddress)','$($JoinEmailAddresses)','$($ADUser.LegacyExchangeDN)'`
                                                ,'$($MB.ExchangeGuid)','$($Random)','$((get-date).ToString('yyyy-MM-dd'))','False','LegacyExchangeDN Not Found'`
                                                )"
                                    Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $database  -Query $SQLQuery 
                                }
                                #User found with that legacyDN so report on success
                                Else {
                                    Write-Log -Message "Found mailuser $($adUser.samaccountname) with legacyExchangeDN $($ADUser.legacyExchangeDN)" #write-host
                                    #Report this is finished
                                    Write-Log -Message "Finished Processing $($adUser.samaccountname) on $($DC)" #write-host
                                    $MailUserobj = New-Object System.Object
                                    $MailUserobj | Add-Member -type NoteProperty -name samaccountname -Value $ADUser.samaccountname
                                    $MailUserobj | Add-Member -type NoteProperty -name DN -Value $ADUser.DistinguishedName
                                    $MailUserobj | Add-Member -type NoteProperty -name Database -Value $MB.Database
                                    $MailUserobj  | Add-Member -type NoteProperty -name PrimaryEmailAddress -Value $CheckMailUser.PrimarySmtpAddress
                                    $MailUserobj  | Add-Member -type NoteProperty -name Alias -Value $MB.Alias
                                    $MailUserobj  | Add-Member -type NoteProperty -name Targetaddress -Value $ADUser.targetAddress 
                                    $MailUserobj | Add-Member -type NoteProperty -name EmailAddresses -Value $JoinEmailAddresses 
                                    $MailUserobj  | Add-Member -type NoteProperty -name LegacyDN -Value $ADUser.LegacyExchangeDN
                                    $MailUserobj | Add-Member -type NoteProperty -name ExchangeGUID -Value $MB.ExchangeGuid
                                    $MailUser += $MailUserobj
                                    $MailUserCounter++
                                    
                                    # Add to SQL Table for reporting
                                    [int]$Random = Get-Random -Maximum 100000 -Minimum 100
                                    $Random
                                    $SQLQuery  = "INSERT INTO QCSMigration  (Samaccountname,DistinguishedName,ExchangeDatabase,PrimarySmtpAddress,Alias,targetAddress,EmailAddresses,LegacyExchangeDN,ExchangeGuid,RandomNumber,RecordAdded,Succeeded,FailureReason) VALUES `
                                                ('$($ADUser.samaccountname)','$($ADUser.DistinguishedName)','$($MB.Database)','$($CheckMailUser.PrimarySmtpAddress)' `
                                                ,'$($MB.Alias)','$($ADUser.targetAddress)','$($JoinEmailAddresses)','$($ADUser.LegacyExchangeDN)'`
                                                ,'$($MB.ExchangeGuid)','$($Random)','$((get-date).ToString('yyyy-MM-dd'))','True','NA'`
                                                )"
                                    Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $database  -Query $SQLQuery 
                                }
                            }
                        }


 
                    }
                    # The mailbox still exists so we cannot proceedd further so report on this and move on to the next user
                    Else {
                        #write-host -ForegroundColor red "Mailbox  $($MB.PrimarySmtpAddress) on $($DC) exists so cant process user $($adUser.samaccountname)"
                        Write-Log -Message "Mailbox  $($MB.PrimarySmtpAddress) on $($DC) exists so cant process user $($adUser.samaccountname)" #write-host -Level "Error"
                        $MBNotDeleteObj = New-Object System.Object
                        $MBNotDeleteObj  | Add-Member -type NoteProperty -name samaccountname -Value $ADUser.samaccountname
                        $MBNotDeleteObj  | Add-Member -type NoteProperty -name MB-Value $MB.PrimarySmtpAddress
                        $MBNotDeleteObj  | Add-Member -type NoteProperty -name Database -Value $MB.Database
                        $MBNotDelete += $MBNotDeleteObj 
                        $MBNotDeleteCounter++
                    }


                }
                # Mailbox is not on QCS database so make note an move on to the next user
                Else {
                    #write-host -ForegroundColor red "Not a QCS Mailbox as Database is $($MB.Database)"
                    Write-Log -Message "$($MB.PrimarySmtpAddress) Database is not in the QCS database $($MB.Database)" #write-host -Level "Error"
                    $NotQCSMailboxObj = New-Object System.Object
                    $NotQCSMailboxObj | Add-Member -type NoteProperty -name samaccountname -Value $ADUser.samaccountname
                    $NotQCSMailboxObj | Add-Member -type NoteProperty -name DN -Value $ADUser.DistinguishedName
                    $NotQCSMailboxObj | Add-Member -type NoteProperty -name Database -Value $MB.Database
                    $NotQCSMailbox += $NotQCSMailboxObj
                    $NotQCSMailboxCounter++
                }

            }
            # User object is either bit in the QCS OU, doesnt have a samaccountname that matches exactly 19 characters or it doesnt have a blank UPN
            Else {
                #write-host -ForegroundColor red "User is not a QCS stub object"
                Write-Log -Message "$($ADUser.Samaccountname) is not in the QCS OU  - $($Aduser.CanonicalName)" #write-host -Level "Error"
                $NotQCSUserObj = New-Object System.Object
                $NotQCSUserObj | Add-Member -type NoteProperty -name User -Value $ADUser.samaccountname
                $NotQCSUserObj | Add-Member -type NoteProperty -name DN -Value $ADUser.DistinguishedName
                $NotQCSUserObj | Add-Member -type NoteProperty -name UPN -Value $ADUser.UserPrincipalName
                $NotQCSUser += $NotQCSUserObj
                $NotQCSUserCounter++
            }
        }
    }
}

# Write a summary of what happened during the run to file and screen
Write-Log -Message "$($USerProcessedCounter) users processed)" #write-host 
#Write-host -ForegroundColor yellow  "$($USerProcessedCounter) users processed" 

Write-Log -Message "$($NoinADCounter) could not be found in Active Directory" #write-host
#Write-host -ForegroundColor yellow  "$($NoinADCounter) could not be found in Active Directory"

Write-Log -Message "$($NoRecipientCounter) had email addresses that could not be resolved" #write-host
#Write-host -ForegroundColor yellow  "$($NoRecipientCounter) had email addresses that could not be resolved"

Write-Log -Message "$($NotMBCounter) email addresses were not attached to mailboxes so couldnt be processed" #write-host
#Write-host -ForegroundColor yellow  "$($NotMBCounter) email addresses were not attached to mailboxes so couldnt be processed" 

Write-Log -Message "$($NotQCSUserCounter) were not in the QCS OU)" #write-host 
#Write-host -ForegroundColor yellow  "$($NotQCSUserCounter) were not in the QCS OU"  


Write-Log -Message "$($NotQCSMailboxCounter) had mailboxes that weren't on the QCS mailboxes database" #write-host
#Write-host -ForegroundColor yellow  "$($NotQCSMailboxCounter) had mailboxes that weren't on the QCS mailboxes database"

Write-Log -Message "$($MBNotDeleteCounter) had mailboxes that weren't deleted" #write-host
#Write-host -ForegroundColor yellow  "$($MBNotDeleteCounter) had mailboxes that weren't deleted"

Write-Log -Message "$($QCSMailboxUserCounter) had mailboxes on the QCS mailboxes database" #write-host
#Write-host -ForegroundColor yellow  "$($QCSMailboxUserCounter) had mailboxes on the QCS mailboxes database" 

Write-Log -Message "$($ErrorMEUCounter) had an issue with configuration of the MEU" 
#Write-host -ForegroundColor yellow  "$($ErrorMEUCounter) had an issue with configuration of the MEU" 

Write-Log -Message "$($MailUserCounter) were successfully processed" 
#Write-host -ForegroundColor yellow  "$($MailUserCounter) were successfully processed"



# We will output all the arrays to a csv if via the export-file function
export-file $NotQCSUser $NotQCSUserout 
export-file $NotQCSMailbox $NotQCSMailboxout 
export-file $QCSMailboxUser $QCSMailboxUserout
export-file $MailUser $MailUserout 
export-file $ErrorMEU  $ErrorMEUout 
export-file $NoRecipient $NoRecipientout 
export-file $NotMB  $NotMBout 
export-file $NointAD $NotinADout 
export-file $MBNotDelete $MBNotDeleteout 




#lastly open the logfile
notepad $logfile
