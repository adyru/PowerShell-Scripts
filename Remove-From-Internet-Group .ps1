
<# 
Script to add remove users from the listed groups

It will email out the results

It uses ARS PS modules so that actions 
#>

#connect to ARS
Connect-QADService -Proxy "server"

#Get the date in a format ot use later
$date = get-date -Format dd-MM-yyyy--hh-mm



#Inpu file
$Input = "\\NAS1.your.domain.dc\share\Remove-Users.txt"

#$Vars for mail send at the end
$SD = "auser@somewhere.com"
$Sender = "InternetAccess@somewhere.com"
$SmtpServer = "smtprelay.your.domain.dc"
$subject = "Internet Access (White)"
$CC = "auser@somewhere.com"
$Admin = "auser@somewhere.com"



$RemovedFromGroupOut = $PSScriptRoot +"\RemovedFromGroups\RemovedFromGroup-" +$date + ".txt"
$NotRemovedFromGroupOut = $PSScriptRoot +"\NotRemovedFromGroup\NotRemovedFromGroup-" +$date + ".txt"
$NotMemberOut = $PSScriptRoot +"\NotMember\NotMember-" +$date + ".txt"
$Not4x4Out   = $PSScriptRoot +"\Not4x4\Not4x4-" +$date + ".txt"
$NotFoundOut = $PSScriptRoot +"\NotFound\NotFound -" +$date + ".txt"
$Old = $PSScriptRoot + "\Old-User-Files\Old-Remove-User-" +$date + ".txt"

# Transcript details
$transcript = $PSScriptRoot +"\transcript\remove-InternetGroups-" +$date + ".txt"
start-transcript $transcript

#Some Arrays for later
$RemovedFromGroup = @()
$NotRemovedFromGroup = @()
$NotFound = @()
$Not4x4 =  @()
$NotMember = @()
$attachments = @()
$UsersremovedfromGroups = 0
$UsersremovedNotfromGroups = 0
$usersNotFound = 0
$UsersNot4x4 = 0
$UsersNotmember = 0

#Check the Input File exists and exit with email if not
$PathExist =  test-path $Input

If ($PathExist -eq $False)
    {
    write-host "File Doesnt Exist"
    send-mailmessage -to $Admin  -from $admin -subject "Remove Users From Internet Groups Input Not Found (White)"  -body "The input file was not found at $Date so the remove users from  internet groups script didnt run <BR> Please check for the input file on the server " -SmtpServer $SmtpServer -cc $CC -BodyAsHtml 
    Exit
    }

$users = get-content $Input
    ForEach ($user in $users)
        {
        $aduser = $null
        
        #Check if the user exists
        $ADuser = Get-QADUser -SamAccountName $user
        $aduser.samaccountname
        #matches samaccountname
      
        #check if valid - not what we use
            if ($ADUser.samaccountname -match "\d{2}[A-Za-z]{6}\d{6}$")
                    {
                    write-host  "Users " $ADUser.samaccountname " Exists"
                    #Check internet group group membership
                        $Adgroups = Get-QADMemberOf $ADUser.samaccountname | ? {$_.name -like "*Test-Internet-*"}
                           #firsof if there are none 
                            If ([STRING]::IsNullOrWhitespace($ADgroups))
                                {
                                write-host $ADUser.samaccountname "is not a member of any internet groups"
                                $NotMemberObj = New-Object System.Object
                                $NotMemberObj | Add-Member -type NoteProperty -name User -Value $ADUser.samaccountname
                                $NotMember += $NotMemberObj
                                $UsersNotmember ++
                                }
                            #So they are a member of groups so we press on
                            Else{
                                ForEach ($ADgroup in $ADgroups)
                                    {
                                    write-host $ADgroup
                                    Remove-QADGroupMember $ADgroup.name -member $ADUser.samaccountname -proxy #-whatif
                                    #Check they have been removed
                                    Start-Sleep -Seconds 1
                                       $check = Get-QADGroupMember $ADgroup.name|  ? {$_.samaccountname -eq $ADUser.samaccountname}
                                        #check if empty as that means success
                                        If ([STRING]::IsNullOrWhitespace($check))
                                            {
                                            #add to array to report later
                                            $RemovedFromGroupObj = New-Object System.Object
                                            $RemovedFromGroupObj | Add-Member -type NoteProperty -name User -Value $ADUser.samaccountname
                                            $RemovedFromGroupObj| Add-Member -type NoteProperty -name Group -Value $ADgroup.name
                                            $RemovedFromGroup += $RemovedFromGroupObj
                                            $UsersremovedfromGroups ++
                                            }
                                        #if this isnt empty we have an issue so add to array to report
                                        Else
                                            {
                                            write-host -ForegroundColor red  "user still in group"
                                            $NotRemovedFromGroupObj = New-Object System.Object
                                            $NotRemovedFromGroupObj | Add-Member -type NoteProperty -name User -Value $ADUser.samaccountname
                                            $NotRemovedFromGroupObj| Add-Member -type NoteProperty -name Group -Value $ADgroup.name
                                            $NotRemovedFromGroup += $NotRemovedFromGroupObj
                                            $UsersremovedNotfromGroups ++
                                            }

                                }
                        #loop through these


                    }
                                }
            # see if it is an empty string - ie not in AD
            If ([STRING]::IsNullOrWhitespace($ADUser.samaccountname))
                    {
                    write-host -ForegroundColor red  "Users " $User " doesnt exist in AD"
                    $NotFoundObj  = New-Object System.Object
                    $NotFoundObj | Add-Member -type NoteProperty -name User -Value $User
                    $NotFound  += $NotFoundObj
                    $usersNotFound ++
                    }
            #LAst account exists but isnt a valid user
            Elseif ($ADUser.samaccountname -notmatch "\d{2}[A-Za-z]{6}\d{6}$")
                    {
                   write-host -ForegroundColor Yellow  "Users " $User " exists in AD but not a 4x4"
                    $Not4x4Obj  = New-Object System.Object
                    $Not4x4Obj | Add-Member -type NoteProperty -name User -Value $User
                    $Not4x4  += $Not4x4Obj
                    $UsersNot4x4 ++

                    }

        }

#Output arrays
If ($RemovedFromGroup -ne $null)
    {$RemovedFromGroup  | Export-Csv -NoClobber -NoTypeInformation -path $RemovedFromGroupOut}

If ($NotRemovedFromGroup -ne $null)     
    {$NotRemovedFromGroup  | Export-Csv -NoClobber -NoTypeInformation -path $NotRemovedFromGroupOut}

If ($NotMember -ne $null)   
    {$NotMember | Export-Csv -NoClobber -NoTypeInformation -path $NotMemberOut}

If ($Not4x4 -ne $null)   
    {$Not4x4   | Export-Csv -NoClobber -NoTypeInformation -path $Not4x4Out}  
 
 If ($NotFound  -ne $null)   
    {$NotFound    | Export-Csv -NoClobber -NoTypeInformation -path $NotFoundOut}   

# not we have the outputs test if they exist as tey dont get created if null
 $TestRemovedFromGroupOut = test-path $RemovedFromGroupOut 
 $testNotRemovedFromGroupOut = test-path $NotRemovedFromGroupOut
 $TestNotMemberOut = test-path $NotMemberOut
 $TestNot4x4Out   = test-path $Not4x4Out   
 $TestNotFoundOut = test-path $NotFoundOut 
 
 #if they do we want to add them to the email so add to array we will use for this
 if ($TestRemovedFromGroupOut -eq $true)
    {
    $attachments += $RemovedFromGroupOut 
    }
 
 
 if ($testNotRemovedFromGroupOut  -eq $true)
    {
    $attachments += $NotRemovedFromGroupOut
    }

 
 if ($TestNotMemberOut -eq $true)
    {
    $attachments +=  $NotMemberOut
    }

 if ( $TestNot4x4Out -eq $true)
    {
    $attachments += $Not4x4Out 
    }
 
 if ($TestNotFoundOut -eq $true)
    {
    $attachments += $NotFoundOut
    }

#Prepare the email using the counters we have used to give an idea of what has been done
$Body = "ServiceDesk <BR> <BR> " + $UsersremovedfromGroups + " have been removed Internet Access Group. <br> <BR> `
" + $UsersremovedNotfromGroups  + " were not removed from groups <BR> <BR>  `
" + $UsersNotmember   + " were not in any Internet groups <BR> <BR> `
" + $usersNotFound + " were not found in AD<BR> <BR> `
" + $UsersNot4x4+ " were not 4x4 accounts<BR> <BR> `
Details Attached <BR> <BR> Regards <br> <BR> Babcock IS"   
#send the email out  
send-mailmessage -to "auser@somewhere.com" -from $sender -subject "Add To Internet groups (White)"  -body $Body -SmtpServer $SmtpServer -cc $CC -BodyAsHtml  -Attachments $attachments
#move-item "\\NAS1.your.domain.dc\SuccessFactor-InternetGroup-Import\Users.txt" $Old


#move input tile
move-item $input  $Old
#Stop transcript
Stop-Transcript
