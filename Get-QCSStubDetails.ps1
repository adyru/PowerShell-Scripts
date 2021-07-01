$UsersArray =  @()

#Get the date for Output
$date = get-date -Format dd-MM-yyyy--hh-mm

#Output Files
$Out = "$($PSScriptRoot)\Output\User-$($date).csv"

$users = Get-Content users.txt
ForEach($user in $users)
    {
    $MB = get-mailbox $user
    $ADUser = get-aduser $MB.samaccountname -Properties targetAddress
    $TArget = $ADUser.targetAddress -replace "^SMTP:",""
    #$MB.Emailaddresses
    $JoinedEmailAddresses = $MB.Emailaddresses -join ","
    $UsersArrayObj = New-Object System.Object
    $UsersArrayObj | Add-Member -type NoteProperty -name Samaccountname -Value $MB.Samaccountname
    $UsersArrayObj | Add-Member -type NoteProperty -name Alias -Value $MB.Alias
    $UsersArrayObj | Add-Member -type NoteProperty -name PrimarySMTPAddress   -Value $MB.PrimarySMTPAddress 
    $UsersArrayObj | Add-Member -type NoteProperty -name TargetAddress -Value $TArget 
    $UsersArrayObj | Add-Member -type NoteProperty -name WindowsEmailAddress  -Value $MB.WindowsEmailAddress  
    $UsersArrayObj | Add-Member -type NoteProperty -name EmailAddresses -Value $JoinedEmailAddresses  #$MB.$Emailaddresses
    $UsersArrayObj | Add-Member -type NoteProperty -name LegacyExchangeDN -Value $MB.LegacyExchangeDN
    $UsersArrayObj | Add-Member -type NoteProperty -name EmailAddressPolicyEnabled -Value $MB.EmailAddressPolicyEnabled
    $UsersArrayObj | Add-Member -type NoteProperty -name ExchangeGuid -Value $MB.ExchangeGuid
    $UsersArray  += $UsersArrayObj
    }
$UsersArray | Export-Csv -NoClobber -NoTypeInformation -path $out -Encoding UTF8