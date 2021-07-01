$DC = "BDCADOMR11.blue01.Babcockgroup.co.uk"
$users = import-csv users.csv -Encoding UTF8 # | select -first 1
foreach($user in $users)
    {
    $CheckUser = get-mailuser -Identity $User.samaccountname -erroraction silentlycontinue
    Start-Sleep -Seconds 1
    If ([STRING]::IsNullOrWhitespace($CheckUser ))
        {
        Enable-MailUser -Identity $User.samaccountname -ExternalEmailAddress $User.targetAddress -alias $user.alias -DomainController $DC 
        }
    Else
        {
        write-host -ForegroundColor Green "$($CheckUser.Name) exists"
        }
    
    Set-ADUser $User.samaccountname -Replace @{legacyExchangeDN=$User.LegacyExchangeDN} -server $DC 
    Set-MailUser -Identity $user.samaccountname  -ExternalEmailAddress $User.targetAddress  -ExchangeGuid $User.ExchangeGUID -WindowsEmailAddress $User.PrimarySMTPAddress -EmailAddressPolicyEnabled $false   -DomainController $DC
    $Emailaddresses = $User.EMAILADDRESSES  -split ","
       ForEach ($Emailaddress in $Emailaddresses)
           {
           $currentEmails = (get-MailUser -Identity $user.samaccountname -DomainController $DC).EMAILADDRESSES
           $Check = $currentEmails -contains $Emailaddress
           If($Check -eq $False)
                {
                write-host "adding $($Emailaddress) to $($user.samaccountname )"
                Set-MailUser $user.samaccountname  -EmailAddresses @{Add=$Emailaddress} -DomainController $DC 
                }
            Else
                {
                write-host -ForegroundColor yellow  "ignoring $($Emailaddress) as it is already on object $($user.displayname )"
                }
           }
   
    }