<#
This script is to fix the potential issue in QCS
Where it has issus opening up mailboxes
with tildas or exclamaion marks in the legacyDN
it will add the exisitng legacyDN as an x500 address
and remove the special character
outputs the values to csv in case of issis
#>

#Variables for later on
$Ampersand = ".*&.*@.*"
$Backslash = ".*\/.*@.*"
$OU = "OU=OU,DC=DomainName,DC=co,DC=uk"
$date = get-date -Format dd-MM-yyyy--hh-mm
$CSVOut =  $Date + "-PrimarySMTP.csv"

#Create Table object
$ALLUsers = New-Object system.Data.DataTable "ALLUsers"
#Define Columns
$ALLUserscol1 = New-Object system.Data.DataColumn Samaccountname,([string])
$ALLUserscol2 = New-Object system.Data.DataColumn OldPrimary,([string]) #Add some Columns
$ALLUserscol3 = New-Object system.Data.DataColumn NewPrimary,([string]) #Add some Columns
$ALLUsers.columns.add($ALLUserscol1)
$ALLUsers.columns.add($ALLUserscol2)
$ALLUsers.columns.add($ALLUserscol3)

#get the users in scope
#For exchange 2016
$AmpersandUsers = Get-Mailbox -OrganizationalUnit $OU -resultsize unlimited | ? {($_.PrimarySmtpAddress -match $Ampersand) -and ($_.database -like "*DB*")}| select -first 1
#For exchange 210
#$AmpersandUsers = Get-Mailbox -OrganizationalUnit $OU | ? {($_.PrimarySmtpAddress -match $Ampersand) -and ($_.database -like "*MDB*")}| select -first 1
#$BackslashUsers = Get-Mailbox -OrganizationalUnit $OU | ? {$_.PrimarySmtpAddress -match $Backslash}| select -first 1

#Process the Eclamation MArk Users
ForEach ($AmpersandUser in $AmpersandUsers )
    {
    $AmpersandUserOld = $AmpersandUser.PrimarySmtpAddress
    $AmpersandUserNew = $AmpersandUser.PrimarySmtpAddress -replace "&","-"
    write-host -ForegroundColor Green " New Primary is " $AmpersandUserNew
    start-sleep -Seconds 5
    Set-Mailbox $AmpersandUser.SamAccountName -EmailAddresses @{Add=$AmpersandUserNew} -confirm -whatif
    Set-Mailbox $AmpersandUser.SamAccountName -PrimarySmtpAddress $AmpersandUserNew -mailAddressPolicyEnabled $false -confirm -whatif
  
    $row = $ALLUsers.NewRow()
    #Enter data in the row
    $Row.samaccountname = $AmpersandUser.SamAccountName
    $row.OldPrimary = $AmpersandUserOld
    $row.NewPrimary = $AmpersandUserNew
    #Add the row to the table
    $ALLUsers.Rows.Add($row)


    }

$ALLUsers | export-csv  $CSVOut -NoClobber -NoTypeInformation