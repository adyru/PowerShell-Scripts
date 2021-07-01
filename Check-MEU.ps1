
$users = import-csv users.csv -Encoding UTF8 # | select -first 1

$DC = "BDCADOMR11.blue01.Babcockgroup.co.uk"

$UsersArray =  @()
$Found =  @()
$NotFound =  @()
$EmailAddressFound =  @()
$EmailAddressNotFound =  @()

#Get the date for Output
$date = get-date -Format dd-MM-yyyy--hh-mm

#Output Files
$Out = "$($PSScriptRoot)\Output\User-$($date).csv"


#Start Stopwatch
$sw = [diagnostics.stopwatch]::StartNew()

ForEach ($user in $users)
    {
    #Do garbage collection every couple of minutes to stop memory going offpiste
    if( $Sw.Elapsed.minutes -eq 2)
        {
        Write-host "Doing Garbage Collection after $($Sw.Elapsed.minutes ) minutes"
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers();
        #Reset timer by stopping and starting a new one
        $Sw.Stop()
        $sw = [diagnostics.stopwatch]::StartNew()

        }
        #$EmailaddressNotFoundToAdd =  @()
        #$EmailAddressFound =  @()
        #$EmailaddressFoundToAdd  =  @()
        #$EmailaddressNotFoundToAdd =  @()
        #$EmailaddressFoundToAdd  =  @()
    $adUser = Get-ADUser $user.samaccountname -server $DC
    If ([STRING]::IsNullOrWhitespace($adUser))
        {
        $ADUserObj = "Not Found"
        write-host -ForegroundColor Red "User $($user.samaccountname) was not found"
        }
    Else
        {
        $ADUserObj = $adUser.Samaccountname
        }


    $alias = Get-MailUser $user.samaccountname -DomainController $DC |? {$_.alias -eq $user.alias}
    If ([STRING]::IsNullOrWhitespace($alias))
        {
        $aliasObj = "Not Found"
        write-host -ForegroundColor Red "alias $($user.alias) was not found"
        }
    Else
        {
        $aliasObj = $alias.alias
        }

    $AddressPolicy  = Get-MailUser $user.samaccountname  -DomainController $DC |? {$_.EmailAddressPolicyEnabled -eq $false}
    If ($AddressPolicy -Eq $true)
        {
        $AddressPolicyObj = "True"
        }
    Else
        {
        $AddressPolicyObj = "False"
        }

    $ExchangeGuid  = Get-MailUser $user.samaccountname -DomainController $DC | ? {$_.ExchangeGuid -eq $user.ExchangeGuid}
    If ([STRING]::IsNullOrWhitespace($ExchangeGuid))
        {
        $ExchangeGuidObj = "Not Found"
        write-host -ForegroundColor Red "ExchangeGuid $($user.ExchangeGuid) was not found"
        }
    Else
        {
        $ExchangeGuidObj = $ExchangeGuid.ExchangeGuid
        }

    $WindowsEmailAddress = Get-MailUser $user.samaccountname -DomainController $DC | ? {$_.WindowsEmailAddress -eq $user.WindowsEmailAddress}
    If ([STRING]::IsNullOrWhitespace($WindowsEmailAddress ))
        {
        $WindowsEmailAddressObj = "Not Found"
        write-host -ForegroundColor Red "WindowsEmailAddress $($user.WindowsEmailAddress) was not found"
        }
    Else
        {
        $WindowsEmailAddressObj = $WindowsEmailAddress.WindowsEmailAddress
        }

    $legacyDN = Get-MailUser $user.samaccountname -DomainController $DC| ? {$_.LegacyExchangeDN -eq $user.LegacyExchangeDN}
    If ([STRING]::IsNullOrWhitespace($legacyDN))
        {
        $legacyDNobj = "Not Found"
        write-host -ForegroundColor Red "LegacyDN $($user.LegacyExchangeDN) was not found"
        }
    Else
        {
        $legacyDNobj = $legacyDN.LegacyExchangeDN 
        }

    $PrimarySmtpAddress = Get-MailUser $user.samaccountname -DomainController $DC | ? {$_.PrimarySmtpAddress -eq $user.PrimarySmtpAddress}
    If ([STRING]::IsNullOrWhitespace($PrimarySmtpAddress))
        {
        $PrimarySmtpAddressobj = "Not Found"
        write-host -ForegroundColor Red "Primary SMTP address $($user.PrimarySmtpAddress) was not found"
        }
    Else
        {
        $PrimarySmtpAddressobj = $PrimarySmtpAddress.PrimarySmtpAddress
        }
    $targetAddressChange = "SMTP:$($user.targetAddress)"
    #write-host "Target address to use is $($targetAddressChange)"
    $targetAddress = Get-MailUser $user.samaccountname -DomainController $DC | ? {$_.ExternalEmailAddress -eq $targetAddressChange}
    If ([STRING]::IsNullOrWhitespace($targetAddress))
        {
        $targetAddressobj = "Not Found"
        write-host -ForegroundColor Red "Target address $($targetAddressChange) was not found"
        }
    Else
        {
        $targetAddressobj  = $targetAddress.ExternalEmailAddress
        }
    
    $Emailaddresses = $user.EmailAddresses -split ","
    ForEach ($Emailaddress in $Emailaddresses)
            {
            #$Emailaddress 
           $CheckEmailAddressesInput = (Get-MailUser $user.samaccountname -DomainController $DC).EmailAddresses
           $inputCheck = $CheckEmailAddressesInput -contains $Emailaddress
           #write-host "Check is $($inputCheck )"
           If($inputCheck  -eq $true)
                {
                #write-host -foregroundcolor "yellow"  "Here"
                $EmailaddressFoundToAdd = $EmailaddressFoundToAdd + "," + $Emailaddress
                #Write-host "Hmm $($EmailaddressFoundToAdd)"
                }
            Else
                {
                $EmailaddressNotFoundToAdd = $EmailaddressNotFoundToAdd + "," + $Emailaddress
                write-host -ForegroundColor Red " Email address $($Emailaddress) was not found"
                }
                #Write-host "End"
            $EmailaddressFoundToAdd = $EmailaddressFoundToAdd  -replace "^,",""
            $EmailaddressNotFoundToAdd = $EmailaddressNotFoundToAdd -replace "^,",""
            #$EmailaddressFoundToAdd 
           }

        $UsersArrayObj = New-Object System.Object
        $UsersArrayObj | Add-Member -type NoteProperty -name Samaccountname -Value $ADUserObj
        $UsersArrayObj | Add-Member -type NoteProperty -name Alias -Value $aliasObj
        $UsersArrayObj | Add-Member -type NoteProperty -name PrimarySMTPAddress  -Value $PrimarySmtpAddressobj
        $UsersArrayObj | Add-Member -type NoteProperty -name TargetAddress -Value $targetAddressobj
        $UsersArrayObj | Add-Member -type NoteProperty -name WindowsEmailAddress  -Value $WindowsEmailAddressObj  
        $UsersArrayObj | Add-Member -type NoteProperty -name FoundEmailAddresses -Value $EmailaddressFoundToAdd
        $UsersArrayObj | Add-Member -type NoteProperty -name NotFoundEmailAddresses  -Value $EmailaddressNotFoundToAdd
        $UsersArrayObj | Add-Member -type NoteProperty -name LegacyExchangeDN -Value $legacyDNobj
        $UsersArrayObj | Add-Member -type NoteProperty -name EmailAddressPolicyEnabled -Value $AddressPolicyObj
        $UsersArrayObj | Add-Member -type NoteProperty -name ExchangeGuid -Value $ExchangeGuidObj
        $UsersArray  += $UsersArrayObj
        $EmailaddressFoundToAdd = ""
        $EmailaddressNotFoundToAdd = ""
    }



$UsersArray | Export-Csv -NoClobber -NoTypeInformation -path $out -Encoding UTF8
