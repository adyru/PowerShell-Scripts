<#
Script will check a users primary smtp address and see if it matches their UPN
If it doesnt it will
    1. Check if the domain part of the email address is a valid UPN Suffix in the forest
    2. If the UPN is already in use
    3. If it is a valid UPN suffix and it isnt currently used it will amend the UPN 

it will output to for the above to csv

V2 - checks to see if the primary smtp address is set, ignores if it isnt and outputs to file
#>

#####################
# Variables 
#####################

#Format date for output files
$date = get-date -Format dd-MMM-yyyy-hh-mm-ss

#Arrays
$Changes = @()
$Matches = @()
$InvalidUPNs = @()
$ExistingUPNs = @()
$NullPrimary = @()

#Output Files
$MatchesNotOut = $date + "-UPN-Doesnt-Match.csv"
$MatchesOut = $date + "-UPN-Matches-PrimaryEmail.csv"
$InvalidUPNsOut  = $date + "-Invalid-UPN.csv"
$ExistingUPNsOut = $date + "-Existing-UPN.csv"
$ChangesOut = $date + "-UPN-Changed.csv"
$NullPrimaryOut = $date + "-Null-Primary.csv"

#Search AD for users
#$users = Get-QADUser  -OrganizationalUnit "OU=Something,DC=Something,DC=co,DC=uk" -properties samaccountname,UserPrincipalName,PrimarySMTPAddress,ParentContainer -SizeLimit 100
$users = Get-QADUser  -OrganizationalUnit "OU=SomethingElse,DC=Something,DC=co,DC=uk" -properties samaccountname,UserPrincipalName,PrimarySMTPAddress,ParentContainer -SizeLimit 20000
write-host "There are " $users.count " to process"

#Get UPNs
$validUPNs = Get-UserPrincipalNamesSuffix 

########################
# Script Block
########################

ForEach ($user in $users)
        {
            if([string]::IsNullOrWhiteSpace($user.PrimarySMTPAddress))
            {
            $myObject = New-Object System.Object
            $myObject | Add-Member -type NoteProperty -name 4x4 -Value $User.SamAccountName
            $myObject | Add-Member -type NoteProperty -name UserPrincipalName -Value $User.UserPrincipalName
            $myObject | Add-Member -type NoteProperty -name PrimarySMTPAddress -Value $User.PrimarySMTPAddress
            $myObject | Add-Member -type NoteProperty -name ParentContainer -Value $User.ParentContainer
            $NullPrimary += $myObject
            }
            Else
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            {
    $count ++
    write-host $count 
    # Garbage Collection to stop memory going off piste
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers();
    #Write-host "UPN is " $user.UserPrincipalName
    #Write-host "email is " $user.PrimarySMTPAddress
    $checkUPN = $null
    #First Check if the users UPN matches its Primary Email Address
    If($user.UserPrincipalName -eq $user.PrimarySMTPAddress)
            {
            #write-host "i am here"
            # It does - we arent interested but we will output to file anyway
            $myObject1 = New-Object System.Object
            $myObject1 | Add-Member -type NoteProperty -name 4x4 -Value $user.SamAccountName
            $myObject1 | Add-Member -type NoteProperty -name UserPrincipalName -Value $user.UserPrincipalName
            $myObject1 | Add-Member -type NoteProperty -name PrimarySMTPAddress -Value $user.PrimarySMTPAddress
            $myObject1 | Add-Member -type NoteProperty -name ParentContainer -Value $user.ParentContainer
            $Matches += $myObject 
            }
    Else
            {
            #The UPN doesnt match the Primary Emnail address
                #lets get just the suffix
            $PrimaryDomain = $user.PrimarySMTPAddress -replace "^.*?(?=@)","" -replace "@",""
            #write-host "Check UPN is " $CheckUPN
            #write-host "UPN is " $PrimaryDomain
            #write-host "Email is " $user.PrimarySMTPAddress
            #Check if the domain used in teh primary smtp address is a valid
                #UPN Suffix and just Output if it is
            If ($CheckUPN = $ValidUPNs -contains $PrimaryDomain -eq $true)
                {
                #So it is a valid UPN Suffix and they dont match
                    #So we need to check if UPN is in use
                    #write-host -ForegroundColor red  "i am here" $user.PrimarySMTPAddress
                    # Search AD for a user with the users email address set as their UPN
                    $UPNExist = Get-QADUser -UserPrincipalName $user.PrimarySMTPAddress -properties samaccountname,UserPrincipalName,PrimarySMTPAddress,ParentContainer
                    $UPNExistCheck = $UPNExist.UserPrincipalName
                    #Write-host "UPN exists" $UPNExistCheck
                    # SO there is a user with that UPN
                    If($UPNExistCheck  -like "*@*")
                        {
                        #Store it in an array for later outputting
                        #write-host "here"
                        Write-host "UPN exists" $UPNExistCheck
                        $myObject2 = New-Object System.Object
                        $myObject2|  Add-Member -type NoteProperty -name 4x4 -Value $UPNExist.SamAccountName
                        $myObject2 | Add-Member -type NoteProperty -name UserPrincipalName -Value $UPNExist.UserPrincipalName
                        $myObject2 | Add-Member -type NoteProperty -name PrimarySMTPAddress -Value $UPNExist.PrimarySMTPAddress
                        $myObject2 | Add-Member -type NoteProperty -name ParentContainer -Value $UPNExist.ParentContainer
                        $ExistingUPNs += $myObject2
                        }
                    Else
                        {
                        # So it is a valid UPN, the UPN doesnt exist
                            # and the upn doesnt match the primary email address
                            # So we will change the UPN and store details for output
                        #Set-QADuser $user.samaccountname -UserPrincipalName $user.PrimarySMTPAddress -confirm -whatif
                        $myObject3 = New-Object System.Object
                        $myObject3 | Add-Member -type NoteProperty -name 4x4 -Value $user.SamAccountName
                        $myObject3 | Add-Member -type NoteProperty -name OldUserPrincipalName -Value $user.UserPrincipalName
                        $myObject3 | Add-Member -type NoteProperty -name NewUserPrincipalName -Value $user.PrimarySMTPAddress
                        $myObject3 | Add-Member -type NoteProperty -name PrimarySMTPAddress -Value $user.PrimarySMTPAddress
                        $myObject3 | Add-Member -type NoteProperty -name ParentContainer -Value $user.ParentContainer
                        $Changes += $myObject3
                        }
                }
            Else
                {
                # We will save users where the domain part of the email address 
                    #is not a valid UPN Suffic
                $myObject4 = New-Object System.Object
                $myObject4 | Add-Member -type NoteProperty -name 4x4 -Value $user.SamAccountName
                $myObject4 | Add-Member -type NoteProperty -name UserPrincipalName -Value $user.UserPrincipalName
                $myObject4 | Add-Member -type NoteProperty -name PrimarySMTPAddress -Value $user.PrimarySMTPAddress
                $myObject4 | Add-Member -type NoteProperty -name ParentContainer -Value $user.ParentContainer
                $InvalidUPNs += $myObject4
                }
                    

            }

        }
    }


# We will go through all the arrays
    # and if there is some entries output to file

If ($Matches -ne $null)
    {$Matches | Export-Csv -NoClobber -NoTypeInformation -path $MatchesOut}


If ($InvalidUPNs -ne $null)
    {$InvalidUPNs | Export-Csv -NoClobber -NoTypeInformation -path $InvalidUPNsOut}

 
If ($ExistingUPNs  -ne $null)
    {$ExistingUPNs | Export-Csv -NoClobber -NoTypeInformation -path $ExistingUPNsOut}


If ($Changes  -ne $null)
    {$Changes  | Export-Csv -NoClobber -NoTypeInformation -path $ChangesOut}

If ($NullPrimary -ne $null)
    {$NullPrimary  | Export-Csv -NoClobber -NoTypeInformation -path $NullPrimaryOut}
