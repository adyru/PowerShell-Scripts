<#
This script is to fix the potential issue in QCS
Where it has issus opening up mailboxes
with tildas or exclamaion marks in the legacyDN
it will add the exisitng legacyDN as an x500 address
and remove the special character
outputs the values to csv in case of issis
#>

#Variables for later on
$Tilda = "/o=***/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=~*"
$Eclamation = "/o=***/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=!*"
$OU = "OU=OU,DC=DomaiName,DC=co,DC=uk"
$date = get-date -Format dd-MM-yyyy--hh-mm
$CSVOut =  $Date + "-legacyDN.csv"

#Create Table object
$ALLUsers = New-Object system.Data.DataTable "ALLUsers"
#Define Columns
$ALLUserscol1 = New-Object system.Data.DataColumn Samaccountname,([string])
$ALLUserscol2 = New-Object system.Data.DataColumn oldLegacyDN,([string]) #Add some Columns
$ALLUserscol3 = New-Object system.Data.DataColumn newLegacyDN,([string]) #Add some Columns
$ALLUserscol4 = New-Object system.Data.DataColumn PrimarySMTP,([string]) #Add some Columns
$ALLUsers.columns.add($ALLUserscol1)
$ALLUsers.columns.add($ALLUserscol2)
$ALLUsers.columns.add($ALLUserscol3)
$ALLUsers.columns.add($ALLUserscol4)

#get the users in scope
$TildaUsers = Get-ADUser -searchbase $OU -properties legacyExchangeDN,EmailAddress  -filter {legacyExchangeDN -like  $tilda} | select -first 10
$EclamationUsers = Get-ADUser -searchbase $OU -properties legacyExchangeDN,EmailAddress  -filter {legacyExchangeDN -like  $Eclamation} | select -first 10

#Process the Eclamation MArk Users
ForEach ($EclamationUser in $EclamationUsers )
    {
    $NewDN = $EclamationUser.legacyExchangeDN -replace "/cn=Recipients/cn=!","/cn=Recipients/cn="
    $x500 = "X500:" + $EclamationUser.legacyExchangeDN
    write-host -ForegroundColor Red "Will add x500 address " $ $x500
    write-host -ForegroundColor Green " New DN is "$NewDN
    #get-Mailbox $EclamationUser.SamAccountName
    #Set-Mailbox $EclamationUser.SamAccountName -EmailAddresses @{Add=$x500} -confirm #-whatif
    #Set-ADUser $EclamationUser.SamAccountName -Replace @{legacyExchangeDN=$NewDN} -confirm # -whatif
    $row = $ALLUsers.NewRow()
    #Enter data in the row
    $Row.samaccountname = $EclamationUser.SamAccountName
    $row.oldLegacyDN = $EclamationUser.legacyExchangeDN 
    $row.newLegacyDN = $NewDN
    $row.PrimarySMTP = $EclamationUser.EmailAddress
    #Add the row to the table
    $ALLUsers.Rows.Add($row)


    }

#Process the Eclamation Tilda Users
ForEach ($TildaUsers in $TildaUsers)
    {
    $NewDN = $TildaUsers.legacyExchangeDN -replace "/cn=Recipients/cn=~","/cn=Recipients/cn="
    $x500 = "X500:" + $TildaUsers.legacyExchangeDN
    write-host -ForegroundColor Red "Will add x500 address " $x500
    write-host -ForegroundColor Green " New DN is " $NewDN
    #get-Mailbox $TildaUsers.SamAccountName
    Set-Mailbox $TildaUsers.SamAccountName -EmailAddresses @{Add=$x500}  -confirm #-whatif
    Set-ADUser $TildaUsers.SamAccountName -Replace @{legacyExchangeDN=$NewDN}  -confirm  #-whatif
    $row = $ALLUsers.NewRow()
    #Enter data in the row
    $Row.samaccountname = $TildaUsers.SamAccountName
    $row.oldLegacyDN = $TildaUsers.legacyExchangeDN 
    $row.newLegacyDN = $NewDN
    $row.PrimarySMTP = $TildaUsers.EmailAddress
    #Add the row to the table
    $ALLUsers.Rows.Add($row)
    }
$ALLUsers | export-csv  $CSVOut -NoClobber -NoTypeInformation