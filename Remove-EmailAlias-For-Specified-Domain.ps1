<#
Searches for all users with a certain alias
then removes only this email address from the users
#>

#What domain are we removing
$EMailDomain = "*@inaer.es"
#Search for these users
$MBs = Get-Mailbox -Filter {emailaddresses -like $EMailDomain}

#use this for testing
#$MBs = Get-mailbox "2utte*" | select -first 1

#Get the date in a format ot use later
$date = get-date -Format dd-MM-yyyy--hh-mm

#Output Files
$RemovedAddressesOut  = $PSScriptRoot +"\ChangedUsers" +$date + ".csv"
$transcript = $PSScriptRoot +"\remove-emailaddressess-" +$date + ".txt"
start-transcript $transcript

#Some Arrays for later
$RemovedAddresses = @()

#loop through users 
ForEach ($MB in $MBs)
    {
    $EmailToRemove = $MB.EmailAddresses | Where {$_ -like $EmailDomain}
    write-host "Email address is " $EmailToRemove
    Write-Host "Account is " $MB.samaccountname 
    Start-Sleep -Seconds 10
    Set-Mailbox $MB.samaccountname  -EmailAddresses @{remove=$EmailToRemove} -confirm # -whatif
    $RemovedAddressesObj = New-Object System.Object
    $RemovedAddressesObj | Add-Member -type NoteProperty -name Samaccountname -Value $MB.samaccountname
    $RemovedAddressesObj| Add-Member -type NoteProperty -name EmailAddressRemoved -Value $EmailToRemove
    $RemovedAddresses += $RemovedAddressesObj
    }
    
$RemovedAddresses   | Export-Csv -NoClobber -NoTypeInformation -path $RemovedAddressesOut
stop-transcript