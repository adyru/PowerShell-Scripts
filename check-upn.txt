
<#

This is to fix an issue where a users upn matches another users
primary email address
EWS fails to work for the user


#>
#Create Table object
$ALLUsers = New-Object system.Data.DataTable "ALLUsers"
#Define Columns
$ALLUserscol1 = New-Object system.Data.DataColumn Samaccountname,([string])
$ALLUserscol2 = New-Object system.Data.DataColumn oldUPN,([string]) #Add some Columns
$ALLUserscol3 = New-Object system.Data.DataColumn newUPN,([string]) #Add some Columns
$ALLUserscol4 = New-Object system.Data.DataColumn OU,([string]) #Add some Columns
$ALLUsers.columns.add($ALLUserscol1)
$ALLUsers.columns.add($ALLUserscol2)
$ALLUsers.columns.add($ALLUserscol3)
$ALLUsers.columns.add($ALLUserscol4)

#out file vars
$date = get-date -Format dd-MM-yyyy--hh-mm-ss
$CSVOut =  $Date + "-leaver-smtp.csv"

#Searck leavers OU amd pull out email addresses
$Leavers = Get-ADUser -Filter * -SearchBase "OU=Leavers,DC=YourDomain,DC=co,DC=uk"  | select UserPrincipalName -first 20000 | ? {$_.UserPrincipalName -match "^([^! '~\(\)]+@domain1.com)" `
    -or $_.UserPrincipalName -match "^([^! '~\(\)]+@domain2.com)" -or $_.UserPrincipalName -match "^([^! '~\(\)]+@domain3.com)" -or $_.UserPrincipalName -match "^([^! '~\(\)]+@domain4.com)" }

#test on one user
#$Leavers = Get-ADUser BROU1404

ForEach ($leaver in $leavers)
    {
        If ($leaver.UserPrincipalName -match ".*[^.]@.*")
        {
                $LeaverChanges = $null
                #write-host "Checking " $leaver.UserPrincipalName
                #$LeaverEmail = $leaver.emailaddress
                #UPN
                $LeaverEmail = $leaver.UserPrincipalName
                #WRITE-hOST $LeaverEmail
                #$LeaverChanges = Get-mailbox -filter "primarysmtpaddress -eq '$LeaverEmail'"  -OrganizationalUnit "OU=User Accounts,DC=YourDomain,DC=co,DC=uk" -ErrorAction SilentlyContinue |select primarysmtpaddress
                # due to having users with ' in the email address need to escape this
                #$LeaverChanges = Get-mailbox -filter "primarysmtpaddress -eq '$($LeaverEmail -Replace "'","''")'" -OrganizationalUnit "OU=User Accounts,DC=YourDomain,DC=co,DC=uk" -ErrorAction SilentlyContinue |select primarysmtpaddress
                #$LeaverChanges = Get-mailbox -filter "primarysmtpaddress -eq '$($LeaverEmail -Replace "'","''" -replace "[\() ~]","-")'" -OrganizationalUnit "OU=User Accounts,DC=YourDomain,DC=co,DC=uk" -ErrorAction SilentlyContinue |select primarysmtpaddress
                $LeaverChanges = Get-aduser -filter "emailaddress -eq '$($LeaverEmail -Replace "'","''" -replace "[\() ~]","-")'" -SearchBase  "OU=User Accounts,DC=YourDomain,DC=co,DC=uk" -ErrorAction SilentlyContinue -Properties emailaddress |select emailaddress
                #$LeaverChanges 
                #write-host "leaverchanges is " $LeaverChanges.primarysmtpaddress
    
                ForEach ($LeaverChange in $LeaverChanges)
                                                                                                                                {
                iF($LeaverChange.emailaddress -like "*@*")
                    {
                    $useremail = $LeaverChange.emailaddress
                    #write-host $useremail
                    #$LeaverToChange = Get-ADUser -filter "emailaddress -eq '$useremail'"   -SearchBase "OU=Leavers,DC=YourDomain,DC=co,DC=uk" -Properties emailaddress,samaccountname,CanonicalName | select -first 1
                    # due to having users with ' in the email address need to escape this
                    $LeaverToChange = Get-ADUser -filter "UserPrincipalName -eq '$($useremail -Replace "'","''" )'"   -SearchBase "OU=Leavers,DC=YourDomain,DC=co,DC=uk" -Properties emailaddress,samaccountname,CanonicalName  | select -first 1
                    write-host $LeaverToChange.UserPrincipalName -ForegroundColor Green
                    #Get a ramdom number to add
                    $Random = get-random -min 100 -maximum 199
                    #create new email var
                    $oldemail = $LeaverToChange.UserPrincipalName
                    $newemail = $oldemail -replace "@","$random@"
                    set-ADUser $LeaverToChange.samaccountname -UserPrincipalName $newemail -Confirm # -WhatIf
                    #Add data to row so we can check and revert if needed
            
                    $row = $ALLUsers.NewRow()
                    #Enter data in the row
                    $Row.samaccountname = $LeaverToChange.samaccountname
                    $row.oldUPN = $oldemail
                    $row.newUPN = $newemail
                    $row.OU = $LeaverToChange.CanonicalName
                    #Add the row to the table
                    $ALLUsers.Rows.Add($row)
                    $LeaverChange = $null
                    }
                }
                $LeaverChange = $null
                
           }
        }
        $LeaverChanges = $null
    write-host "LEavers are " $LeaverChange
   $ALLUsers | export-csv  $CSVOut -NoClobber -NoTypeInformation
   
