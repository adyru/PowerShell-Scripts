<#
This will convert a list of mailboxes
into MEU
#>



Function Write-Log {
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$False)]
    [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
    [String]
    $Level = "INFO",

    [Parameter(Mandatory=$True)]
    [string]
    $Message,

    [Parameter(Mandatory=$False)]
    [string]
    $logfile
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp $Level $Message"
    If($logfile) {
        Add-Content $logfile -Value $Line
    }
    Else {
        Write-Output $Line
    }
}


######################
# Start of Variables #
######################

# Static DC to use 
$DC = "RPWDOMC01"

#Arrays
$NotQCSUser= @()
$NotQCSMailbox = @()
$QCSMailboxUser =  @()
$MailUser =  @()
$ErrorMEU  =  @()
$NoRecipient =  @()
$NotMB =  @()
$NoinAD =  @()
$MBNotDelete =  @()
#Get the date for Output
$date = get-date -Format dd-MM-yyyy--hh-mm

#Output Files
$NotQCSUserOut = "$($PSScriptRoot)\Output\NotQCSUser-$($date).csv"
$NotQCSMailboxOut  = "$($PSScriptRoot)\Output\NotQCSMailbox-$($date).csv"
$QCSMailboxUserOut  =  "$($PSScriptRoot)\Output\QCSMailboxUser-$($date).csv"
$MailUserOut  =  "$($PSScriptRoot)\Output\MailUser-$($date).csv"
$ErrorMEUOut =  "$($PSScriptRoot)\Output\MailEU-Error-$($date).csv"
$NoRecipient  =  "$($PSScriptRoot)\Output\Recipient-Error-$($date).csv"
$NotMBOut   =  "$($PSScriptRoot)\Output\Recipient-Not-MB-$($date).csv"
$NoinAD  =  "$($PSScriptRoot)\Output\Recipient-Not-in-AD-$($date).csv"
$MBNotDeleteOut = "$($PSScriptRoot)\output\MB-Not_deleted-$($date).csv"

$logfile = "$($PSScriptRoot)\output\Convert-MEU-Log-$($date).txt"


#Counters
[int]$USerProcessedCounter = 0
[int]$NotQCSUserCounter = 0
[int]$NotQCSMailboxCounter = 0
[int]$QCSMailboxUserCounter = 0
[int]$MailUserCounter = 0
[int]$ErrorMEUCounter  =  0
[int]$NoRecipientCounter  = 0
[int]$NotMBCounter =  0
[int]$NoinADCounter =  0
[int]$MBNotDeleteCounter = 0

#Inputfiles
$UsersFile = users.txt

###################
# End of Vaiables #
###################

#DC Connectivity Check - Exits on failure
write-host "Domain Controller to use is  $($DC)"
Start-Sleep -Seconds 10

Write-Log -Message "Setting Domain Controller to  $($DC )" -logfile $logfile
Write-Log -Message "Checking connectivity to Domain Controller $($DC )" -logfile $logfile
$CheckCOnnection = test-netconnection $DC -port 135
If($CheckCOnnection.TcpTestSucceeded -eq $false)
    {
    write-host -ForegroundColor red "Connection is $($CheckCOnnection.TcpTestSucceeded) so exiting script, amend DC variable and re run" 
    Write-Log -Message "Connectivity to Domain Controller $($DC ) is $($CheckCOnnection.TcpTestSucceeded) so exiting script " -logfile $logfile -Level "Fatal"
    exit
    }
Else
    {
    write-host -ForegroundColor Green "Connectivity to Domain Controller $($DC ) Succdded - check connection is  $($CheckCOnnection.TcpTestSucceeded) "
    Write-Log -Message "Connectivity to Domain Controller $($DC ) Succdded - check connection is  $($CheckCOnnection.TcpTestSucceeded) " -logfile $logfile
    }


#Check import file and import users if present
Write-Log -Message "Checking $($UsersFile ) is available"
$CheckUserImport = Test-Path $UsersFile
If($CheckCOnnection.TcpTestSucceeded -eq $false)
    {
    write-host "$($UsersFile) doesnt exist so exiting script, amend $UsersFile variable and re run" -ForegroundColor red
    Write-Log -Message "$($UsersFile) doesnt exist so exiting script" -logfile $logfile -Level "Fatal"
    exit
    }
Else
    {
    write-host "$($UsersFile) exists"
    Write-Log -Message "$($UsersFile) exists" -logfile $logfile 
    }

Import the users file
$USers = get-content $UsersFile | select -First 1

$UsersCount = $USers | measure $user
write-host "$($UsersCount.Count) to be processed"
Write-Log -Message "$($UsersCount) to be processed" -logfile $logfile 

ForEach($user in $users)
    {
    #First of always remember do garbage collection to stop memory going offpiste
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers();
    $USerProcessedCounter ++
    Write-Log -Message "Processing $($USerProcessedCounter) of $($UsersCount.Count) mailboxes" -logfile $logfile 
    # Check recipient type
    $RecipientType = get-recipient $user
    #check recipient exists
    If ([STRING]::IsNullOrWhitespace($RecipientType))
        {
        # no match so log this  and move on  to next user
        write-host -ForegroundColor Red "Cannot find a recipient matching $($user)"
        Write-Log -Message "Cannot find a recipient matching $($user)" -logfile $logfile -Level "Error"
        $NoRecipientObj = New-Object System.Object
        $NoRecipientObj | Add-Member -type NoteProperty -name Recipient -Value $User
        $NoRecipient  += $NoRecipientObj
        $NoRecipientCounter++
        }
    # check if recipient is a mailbox
    Elseif($RecipientType.RecipientType -ne "UserMailbox")
        {
        # recipeint isnt a mailbox so log this and move on  to next user
        write-host -ForegroundColor Red "Recipient is not a mailbox - recipient type is matching $($RecipientType.RecipientType)"
        Write-Log -Message "Recipient is not a mailbox - recipient type is matching $($RecipientType.RecipientType)" -logfile $logfile -Level "Error"
        $NotMBObj = New-Object System.Object
        $NotMBObj | Add-Member -type NoteProperty -name Recipient -Value $User
        $NotMBObj | Add-Member -type NoteProperty -name Type -Value $RecipientType.RecipientType 
        $NotMB  += $NotMBObj
        $NotMBCounter++
        }
    Else
    # Mailbox found
        {
        Write-Log -Message "$($User) is $($RecipientType.RecipientType)" -logfile $logfile
        Write-Log -Message "Searching for $($User) in AD" -logfile $logfile  
                $ADUser = Get-ADUser -Filter 'mail -eq $user' -properties * 
                If ([STRING]::IsNullOrWhitespace($ADUser))
                    {
                    Write-host -ForegroundColor Red  -Message "Found $($ADUser)in AD"
                    Write-Log -Message "$($User) not found in AD" -logfile $logfile -Level "Error"
                    $NoinADObj = New-Object System.Object
                    $NoinADObj  | Add-Member -type NoteProperty -name Recipient -Value $User
                    $NoinAD  += $NoinADObj 
                    $NoinADCounter++
                    }
                Else
                    {
                        Write-Log -Message "Found $($ADUser.samaccountname)in AD" -logfile $logfile  
                        If($ADUser.DistinguishedName -like "*OU=Quest Collaboration Services Objects,OU=QCS*")
                            {
                            write-host -ForegroundColor Green "Users is okay"
                            Write-Log -Message "$($ADUser.Samaccountname) is in the QCS OU" -logfile $logfile 
                            $MB = get-mailbox  $user | Select @{L="NewEmailAddresses";E={$_.EmailAddresses}},*
                                If($MB.Database -like "*QCS*")
                                    {
                                    write-host -ForegroundColor Green "Database is $($MB.Database)"
                                    Write-Log -Message "$($MB.PrimarySmtpAddress) Database is $($MB.Database)" -logfile $logfile 
                                    $QCSMailboxUserObj = New-Object System.Object
                                    $QCSMailboxUserObj | Add-Member -type NoteProperty -name User -Value $ADUser.samaccountname
                                    $QCSMailboxUserObj | Add-Member -type NoteProperty -name DN -Value $ADUser.DistinguishedName
                                    $QCSMailboxUserObj | Add-Member -type NoteProperty -name Database -Value $MB.Database
                                    $QCSMailboxUserObj  | Add-Member -type NoteProperty -name PrimaryEmailAddress -Value $CheckMailUser.PrimarySMTPAddress
                                    $QCSMailboxUserObj  | Add-Member -type NoteProperty -name Targetaddress -Value $ADUser.targetAddress 
                                    $QCSMailboxUserObj  | Add-Member -type NoteProperty -name EmailAddresses -Value $MB.NewEmailAddresses
                                    $QCSMailboxUserObj  | Add-Member -type NoteProperty -name LegacyDN -Value $ADUser.LegacyExchangeDN
                                    $QCSMailboxUserObj | Add-Member -type NoteProperty -name ExchangeGUID -Value $MB.ExchangeGuid
                                    $QCSMailboxUser  += $QCSMailboxUserObj
                                    $QCSMailboxUserCounter++
                                    Write-Log -Message "Disabling mailbox$($MB.PrimarySmtpAddress) on $($DC) " -logfile $logfile
                                    Disable-Mailbox -Identity $MB.PrimarySmtpAddress -DomainController $DC
                                    Start-Sleep -Milliseconds 100
                                    $checkMB = get-mailbox $MB.PrimarySmtpAddress
                                    If ([STRING]::IsNullOrWhitespace($checkMB))
                                        {
                                        Write-Log -Message "Mailbox  $($MB.PrimarySmtpAddress) on $($DC) doesnt exist" -logfile $logfile
                                        Write-Log -Message "Enabling Mailuser $($adUser.samaccountname) on $($DC) with External Address $($ADUser.targetAddress)" -logfile $logfile
                                        Enable-MailUser -Identity $adUser.samaccountname -ExternalEmailAddress $ADUser.targetAddress -alias $MB.alias -DomainController $DC 
                                        Start-Sleep -Milliseconds 100
                                        $CheckMailUser = Get-mailuser -Identity $adUser.samaccountname
                                             If ([STRING]::IsNullOrWhitespace($CheckMailUser))
                                                {
                                                write-host -ForegroundColor red "Error finding mailuser $($adUser.samaccountname)"
                                                Write-Log -Message "Cannot find MEU $($ADUser.Samaccountname)" -logfile $logfile -Level "Error"
                                                $ErrorMEUObj = New-Object System.Object
                                                $ErrorMEUObjObj | Add-Member -type NoteProperty -name User -Value $ADUser.samaccountname
                                                $ErrorMEUObjObj | Add-Member -type NoteProperty -name DN -Value $ADUser.DistinguishedName
                                                $ErrorMEUObjObj | Add-Member -type NoteProperty -name Database -Value $MB.Database
                                                $ErrorMEUObjObj  | Add-Member -type NoteProperty -name PrimaryEmailAddress -Value $CheckMailUser.PrimarySMTPAddress
                                                $ErrorMEUObjObj  | Add-Member -type NoteProperty -name Targetaddress -Value $ADUser.targetAddress 
                                                $ErrorMEUObjObj  | Add-Member -type NoteProperty -name EmailAddresses -Value $MB.NewEmailAddresses
                                                $ErrorMEUObjObj  | Add-Member -type NoteProperty -name LegacyDN -Value $ADUser.LegacyExchangeDN
                                                $ErrorMEUObjObj | Add-Member -type NoteProperty -name ExchangeGUID -Value $MB.ExchangeGuid
                                                $ErrorMEU += $ErrorMEUObj
                                                $ErrorMEUCounter++
                                                }
                                            Else
                                                {
                                                Write-Log -Message "Found MEU $($ADUser.Samaccountname)" -logfile $logfile 
                                                Write-Log -Message "Setting Mailuser $($adUser.samaccountname) on $($DC) with External Address $($ADUser.targetAddress)" -logfile $logfile
                                                Set-MailUser -Identity $adUser.samaccountname -EmailAddressPolicyEnabled $false -ExternalEmailAddress $ADUser.targetAddress  -EmailAddresses $MB.EMAILADDRESSES -ExchangeGuid $MB.GUID -DomainController $DC -WindowsEmailAddress $MB.PrimarySmtpAddress
                                                Start-Sleep -Milliseconds 100
                                                $CheckMailUser2 = Get-mailuser -Identity $adUser.samaccountname | ? {$_.ExchangeGuid -eq $MB.GUID}
                                                If ([STRING]::IsNullOrWhitespace($CheckMailUser2))
                                                    {
                                                    write-host -ForegroundColor red "Error finding mailuser $($adUser.samaccountname) with ExchangeGUID $($MB.GUID)"
                                                    Write-Log -Message "Error finding mailuser $($adUser.samaccountname) with ExchangeGUID $($MB.GUID)" -logfile $logfile -Level "Error"
                                                    $ErrorMEUObj = New-Object System.Object
                                                    $ErrorMEUObjObj | Add-Member -type NoteProperty -name User -Value $ADUser.samaccountname
                                                    $ErrorMEUObjObj | Add-Member -type NoteProperty -name DN -Value $ADUser.DistinguishedName
                                                    $ErrorMEUObjObj | Add-Member -type NoteProperty -name Database -Value $MB.Database
                                                    $ErrorMEUObjObj  | Add-Member -type NoteProperty -name PrimaryEmailAddress -Value $CheckMailUser.PrimarySMTPAddress
                                                    $ErrorMEUObjObj  | Add-Member -type NoteProperty -name Targetaddress -Value $ADUser.targetAddress 
                                                    $ErrorMEUObjObj  | Add-Member -type NoteProperty -name EmailAddresses -Value $MB.NewEmailAddresses
                                                    $ErrorMEUObjObj  | Add-Member -type NoteProperty -name LegacyDN -Value $ADUser.LegacyExchangeDN
                                                    $ErrorMEUObjObj | Add-Member -type NoteProperty -name ExchangeGUID -Value $MB.ExchangeGuid
                                                    $ErrorMEU += $ErrorMEUObj
                                                    }
                                                Else
                                                    {
                                                    Write-Log -Message "Found mailuser $($adUser.samaccountname) with ExchangeGUID $($MB.GUID)" -logfile $logfile -Level "Error"
                                                    Write-Log -Message "Setting AD User $($adUser.samaccountname) on $($DC) with legacyDN $($ADUser.legacyExchangeDN)" -logfile $logfile
                                                    Set-ADUser $adUser.samaccountname -Replace @{legacyExchangeDN=$ADUser.legacyExchangeDN} -server $DC 
                                                    $CheckLegacyDN = $adUser.samaccountname | ? {$_.legacyExchangeDN -eq $ADUser.legacyExchangeDN}
                                                    If ([STRING]::IsNullOrWhitespace($CheckLegacyDN))  
                                                        {
                                                        write-host -ForegroundColor red "Error finding AD User $($adUser.samaccountname) with legacyExchangeDN $($ADUser.legacyExchangeDN)"
                                                        Write-Log -Message "Error finding mailuser $($adUser.samaccountname) with ExchangeGUID $($MB.GUID)" -logfile $logfile -Level "Error"
                                                        $ErrorMEUObj = New-Object System.Object
                                                        $ErrorMEUObjObj | Add-Member -type NoteProperty -name User -Value $ADUser.samaccountname
                                                        $ErrorMEUObjObj | Add-Member -type NoteProperty -name DN -Value $ADUser.DistinguishedName
                                                        $ErrorMEUObjObj | Add-Member -type NoteProperty -name Database -Value $MB.Database
                                                        $ErrorMEUObjObj  | Add-Member -type NoteProperty -name PrimaryEmailAddress -Value $CheckMailUser.PrimarySMTPAddress
                                                        $ErrorMEUObjObj  | Add-Member -type NoteProperty -name Targetaddress -Value $ADUser.targetAddress 
                                                        $ErrorMEUObjObj  | Add-Member -type NoteProperty -name EmailAddresses -Value $MB.NewEmailAddresses
                                                        $ErrorMEUObjObj  | Add-Member -type NoteProperty -name LegacyDN -Value $ADUser.LegacyExchangeDN
                                                        $ErrorMEUObjObj | Add-Member -type NoteProperty -name ExchangeGUID -Value $MB.ExchangeGuid
                                                        $ErrorMEU += $ErrorMEUObj
                                                        }
                                                    Else
                                                        {
                                                        Write-Log -Message "Found mailuser $($adUser.samaccountname) with ExchangeGUID $($MB.GUID)" -logfile $logfile
                                                        Write-Log -Message "Finished Processing $($adUser.samaccountname) on $($DC)" -logfile $logfile
                                                        $MailUserobj = New-Object System.Object
                                                        $MailUserobj | Add-Member -type NoteProperty -name User -Value $ADUser.samaccountname
                                                        $MailUserobj | Add-Member -type NoteProperty -name DN -Value $ADUser.DistinguishedName
                                                        $MailUserobj | Add-Member -type NoteProperty -name Database -Value $MB.Database
                                                        $MailUserobj  | Add-Member -type NoteProperty -name PrimaryEmailAddress -Value $CheckMailUser.PrimarySMTPAddress
                                                        $MailUserobj  | Add-Member -type NoteProperty -name Targetaddress -Value $ADUser.targetAddress 
                                                        $MailUserobj | Add-Member -type NoteProperty -name EmailAddresses -Value $MB.NewEmailAddresses
                                                        $MailUserobj  | Add-Member -type NoteProperty -name LegacyDN -Value $ADUser.LegacyExchangeDN
                                                        $MailUserobj | Add-Member -type NoteProperty -name ExchangeGUID -Value $MB.ExchangeGuid
                                                        $MailUser += $MailUserobj
                                                        $MailUserCounter++
                                                        }
                                                    }
                                                }


 
                                            }
                                    Else
                                        {
                                        write-host -ForegroundColor red "Mailbox  $($MB.PrimarySmtpAddress) on $($DC) exists so cant process user $($adUser.samaccountname)"
                                        Write-Log -Message "Mailbox  $($MB.PrimarySmtpAddress) on $($DC) exists so cant process user $($adUser.samaccountname)" -logfile $logfile -Level "Error"
                                        $MBNotDeleteObj = New-Object System.Object
                                        $MBNotDeleteObj  | Add-Member -type NoteProperty -name User -Value $ADUser.samaccountname
                                        $MBNotDeleteObj  | Add-Member -type NoteProperty -name MB-Value $MB.PrimarySmtpAddress
                                        $MBNotDeleteObj  | Add-Member -type NoteProperty -name Database -Value $MB.Database
                                        $MBNotDelete   += $MBNotDeleteObj 
                                        $MBNotDeleteCounter++
                                        }


                                    }
                                Else
                                    {
                                    write-host -ForegroundColor red "Not a QCS Mailbox as Database is $($MB.Database)"
                                    Write-Log -Message "$($MB.PrimarySmtpAddress) Database is not in the QCS database $($MB.Database)" -logfile $logfile -Level "Error"
                                    $NotQCSMailboxObj = New-Object System.Object
                                    $NotQCSMailboxObj | Add-Member -type NoteProperty -name User -Value $ADUser.samaccountname
                                    $NotQCSMailboxObj | Add-Member -type NoteProperty -name DN -Value $ADUser.DistinguishedName
                                    $NotQCSMailboxObj | Add-Member -type NoteProperty -name Database -Value $MB.Database
                                    $NotQCSMailbox  += $NotQCSMailboxObj
                                    $NotQCSMailboxCounter++
                                    }

                            }
                        Else
                            {
                            write-host -ForegroundColor red "User is not stub object"
                            Write-Log -Message "$($ADUser.Samaccountname) is not in the QCS OU  - $($Aduser.CanonicalName)" -logfile $logfile -Level "Error"
                            $NotQCSUserObj = New-Object System.Object
                            $NotQCSUserObj | Add-Member -type NoteProperty -name User -Value $ADUser.samaccountname
                            $NotQCSUserObj | Add-Member -type NoteProperty -name DN -Value $ADUser.DistinguishedName
                            $NotQCSUser += $NotQCSUserObj
                            $NotQCSUserCounter++
                            }
        }
        }
    }

# Write a summary of what happened during the run to file and screen
Write-Log -Message "$($USerProcessedCounter) users processed)" -logfile $logfile 
Write-host -ForegroundColor yellow  "$($USerProcessedCounter) users processed)" 

Write-Log -Message "$($NoinADCounter) could not be found in Active Directory)" -logfile $logfile
Write-host -ForegroundColor yellow  "$($NoinADCounter) could not be found in Active Directory)"

Write-Log -Message "$($NoRecipientCounter) had email addresses that could not be resolved)" -logfile $logfile
Write-host -ForegroundColor yellow  "$($NoRecipientCounter) had email addresses that could not be resolved)"

Write-Log -Message "$($NotMBCounter) email addresses were not attached to mailboxes so couldnt be processed)" -logfile $logfile
Write-host -ForegroundColor yellow  "$($NotMBCounter) email addresses were not attached to mailboxes so couldnt be processed)" 

Write-Log -Message "$($NotQCSUserCounter) were not in the QCS OU)" -logfile $logfile 
Write-host -ForegroundColor yellow  "$($NotQCSUserCounter) were not in the QCS OU)"  


Write-Log -Message "$($NotQCSMailboxCounter) had mailboxes that weren't on the QCS mailboxes database)" -logfile $logfile
Write-host -ForegroundColor yellow  "$($NotQCSMailboxCounter) had mailboxes that weren't on the QCS mailboxes database)"

Write-Log -Message "$($MBNotDeleteCounter) had mailboxes that weren't deleted)" -logfile $logfile
Write-host -ForegroundColor yellow  "$($MBNotDeleteCounter) had mailboxes that weren't deleted)"

Write-Log -Message "$($QCSMailboxUserCounter) had mailboxes on the QCS mailboxes database)" -logfile $logfile
Write-host -ForegroundColor yellow  "$($QCSMailboxUserCounter) had mailboxes on the QCS mailboxes database)" 

Write-Log -Message "$($ErrorMEUCounter) had an issue with configuration of the MEU)" -logfile $logfile
Write-host -ForegroundColor yellow  "$($ErrorMEUCounter) had an issue with configuration of the MEU)" 

Write-Log -Message "$($MailUserCounter) were successfully processed)" -logfile $logfile
Write-host -ForegroundColor yellow  "$($MailUserCounter) were successfully processed)"



#Write-Output
$NotQCSUser  | Export-Csv -NoClobber -NoTypeInformation -path $NotQCSUserOut
$NotQCSMailbox  | Export-Csv -NoClobber -NoTypeInformation -path $NotQCSMailboxOut
$QCSMailboxUser | Export-Csv -NoClobber -NoTypeInformation -path $QCSMailboxUserOut 
$MailUser | Export-Csv -NoClobber -NoTypeInformation -path $MailUserOut 
$ErrorMEU | Export-Csv -NoClobber -NoTypeInformation -path $ErrorMEUOut 
$NoRecipient | Export-Csv -NoClobber -NoTypeInformation -path $NoRecipientOut 
$NotMB | Export-Csv -NoClobber -NoTypeInformation -path $NotMBOut 
$NoinAD | Export-Csv -NoClobber -NoTypeInformation -path $NoinADOut 
$MBNotDelete   | Export-Csv -NoClobber -NoTypeInformation -path $MBNotDeleteOut


##open logfile
notepad $logfile
