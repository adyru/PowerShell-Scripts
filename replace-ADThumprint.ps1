<#
Script to take a unicode csv
with a hex photo and upload 
it to a AD Users ThumbnailPhoto

Relies on function Resize-Image - with below amendment as it locks the image files
(https://gallery.technet.microsoft.com/scriptcenter/Resize-Image-A-PowerShell-3d26ef68)


Need the addition of the .Dispose highlighted below to stop the above 

        $img2.Save($OutputFile);

    }
    	$img2.Dispose()
    	$img.Dispose()

V3   - Put picture resizing

V3  - check if hex is in use
    - add reporting email  

v4  - check the if the existing  photo in AD is the same as in the csv, just outputs if it does
    - amended how attachments are handled
#>

#Get the date in a format ot use later
$date = get-date -Format dd-MM-yyyy--hh-mm

#Input/Output Vars
$filename = $PSScriptRoot + "\TempPic.jpg"
$InputFile = "D:\Adrian\Scripts\Projects\IDM\Script\photo_hex2a.csv"

#Email Vars
$Domain = (Get-ADDomain).name
$SmtpServer = "smtprelay.$domain.domain1.co.uk"
$CC = "someone@domain1.com"
$Admin = "someone@domain1.com"
$ManagmentReportsDL = "someone@domain1.com"
$Sender = "ADPhotossomeone@domain1.com"
$subject = "AD Photos Error ($domain)"
#Arrays
$NotFound = @()
$DuplicateHex = @()
$PhotoSame = @()
$UsersChanged = @()
$attachments = @()

#Hash Table
$hash = $null
$hash = @{}

#OutputFiles
$DuplicateHexOut =$PSScriptRoot +"\Output\Duplicate-Hex-" +$date + ".csv"
$UsersChangedOut = $PSScriptRoot +"\Output\Changed-Users-" +$date + ".csv"
$NotFoundOut = $PSScriptRoot +"\Output\Not-Found-Users-" +$date + ".csv"
$PhotoSameOut = $PSScriptRoot +"\Output\Users-Same-Photo-" +$date + ".csv"

#Counters
$usersNotFound = 0
$DuplicateHashCount = 0
$UsersChangedCount = 0
$PhotoSameCount = 0

# Transcript details
$transcript = $PSScriptRoot +"\transcript\replace-add-thumpprint" +$date + ".txt"

#Kick Of transcript
start-transcript $transcript

#Check Input Exists - email if fails
If ($PathExist -eq $False)
    {
    write-host "File Doesnt Exist"
    send-mailmessage -to $Admin  -from $Sender -subject $Subject  -body "The input file was not found at $Date so the add user photos script didnt run<BR> `
        Please check for the input file on the server " -SmtpServer $SmtpServer -cc $CC -BodyAsHtml 
    Exit
    }

#Input exists so import the user file
$users = import-csv $InputFile  –Delimiter “;” -encoding unicode # | select -First 1

#Import resize
Import-Module .\Resize-Image.psm1

#test the above has loaded
$TestModule = Get-Module Resize-Image

#If it didnt load quit and email
IF($TestModule.Name -eq "Resize-Image")
    {
    write-host "Pressing on"
    }
Else
    {
    send-mailmessage -to $Admin  -from $Sender -subject $Subject  -body "Resize-Image Module failed on $Date so the add user photos script didnt run<BR> `
        Please check for the input file on the server " -SmtpServer $SmtpServer -cc $CC -BodyAsHtml 
    Exit
    }

#Start looping through the users
ForEach($user in $users)
    {
    #First of do garbage collection to stop memory going off piste
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers();

    #null the arrays again
    $checkADUser = $null
    $hex = $null
    $bytes = $null
    $photo = $null
    $CheckPhoto = $null
    $CompareBytes = $null

    #write-host "sleeping for 1 secs"
    Start-Sleep -Seconds 1
    $checkADUser = Get-QADUser $user.dn
    #User not found so will ignore
     If ([STRING]::IsNullOrWhitespace($checkADUser))
        {
        $usersNotFound++
        $NotFoundObj = New-Object System.Object
        $NotFoundObj | Add-Member -type NoteProperty -name User -Value $user.DN
        $NotFound  += $NotFoundObj
        }
    #user exists

    Else

        {
            #Check if the hash has been used this run
            $CheckHash = $hash.ContainsValue($USER.thumbnailPhoto)
            #write-host "Check hash is " $CheckHash 
                #It hasnt carry on as usual
                If ($CheckHash -eq $false)
                    {
                        #get the hex and convert it to a byte array
                        $hex = $USER.thumbnailPhoto
                        $bytes = New-Object -TypeName byte[] -ArgumentList ($hex.Length / 2)
                        for ($i = 0; $i -lt $hex.Length; $i += 2) {
                            $bytes[$i / 2] = [System.Convert]::ToByte($hex.Substring($i, 2), 16)
                        }

                            #Output the bytes to file and then resize
                            [IO.File]::WriteAllBytes($filename,$bytes)
                            Resize-Image -InputFile $filename -Width 96 -Height 96
                            #finally convert back to byte array and upload the file
                            $photo = [byte[]](Get-Content $filename -Encoding byte)

                            #Get Users current photo hex
                            $CheckPhoto = Get-ADUser $checkADUser.samaccountname -Properties * | select thumbnailPhoto
                            #Cant compare a null value so give a value for the comparision operator if it is null
                                #IF there is a photo we will compare the bytes array
                            IF($CheckPhoto.thumbnailPhoto -eq $null)
                                {
                                $CompareBytes -eq $false
                                }
                            Else
                                {
                                $CompareBytes = [System.Linq.Enumerable]::SequenceEqual($CheckPhoto.thumbnailPhoto,$photo)
                                
                                }
                            #The new photo will be the same as the current photo so we will just log this
                            If($CompareBytes  -eq $true )
                                {
                                #Add the hex to hash table
                                $Hash.Add($checkADUser.samaccountname, $USER.thumbnailPhoto)
                                #write-host -ForegroundColor yellow "here"
                                $checkADUser.samaccountname 
                                write-host -ForegroundColor Red   "Photo is the same for user " $checkADUser.samaccountname 
                                $PhotoSameCount++
                                $PhotoSameObj = New-Object System.Object
                                $PhotoSameObj | Add-Member -type NoteProperty -name User -Value $checkADUser.samaccountname
                                $PhotoSame  += $PhotoSameObj
                                }
                            #The photo doesnt match the current photo so we will update
                            Else
                                {
                                
                                Set-ADUser $checkADUser.samaccountname -Replace @{thumbnailPhoto=$photo} -confirm #-whatif
                        
                                #Add the hex to hash table
                                $Hash.Add($checkADUser.samaccountname, $USER.thumbnailPhoto)

                                #log to array and update counter
                                $UsersChangedCount++
                                $UsersChangedObj = New-Object System.Object
                                $UsersChangedObj | Add-Member -type NoteProperty -name User -Value $checkADUser.samaccountname
                                $UsersChanged += $UsersChangedObj
                                #delete temp file for next run
                                remove-item $filename -force #-confirm
                                }
                                
                            
                    }
                #Check if the hex has been used
                ElseIf ($CheckHash -eq $true)
                    {
                    #Output to screen and just log to array
                    write-host -ForegroundColor  Red  "Hash already used this run"
                    $DuplicateHashCount++
                    $DuplicateHexObj = New-Object System.Object
                    $DuplicateHexObj | Add-Member -type NoteProperty -name User -Value $checkADUser.samaccountname
                    $DuplicateHex  += $DuplicateHexObj
                    }

        }
}


#Output variables as long as they arent null

If ($DuplicateHex -ne $null)
    {$DuplicateHex | Export-Csv -NoClobber -NoTypeInformation -path $DuplicateHexOut}

If ($UsersChanged -ne $null)     
    {$UsersChanged  | Export-Csv -NoClobber -NoTypeInformation -path $UsersChangedOut}

If ($NotFound -ne $null)   
    {$NotFound | Export-Csv -NoClobber -NoTypeInformation -path $NotFoundOut}

If ($PhotoSame-ne $null)   
    {$PhotoSame | Export-Csv -NoClobber -NoTypeInformation -path $PhotoSameOut}
                       
                                
#test the output files exist
    $TestDuplicateHexOut = test-path $DuplicateHexOut
    $testUsersChangedOut = test-path $UsersChangedOut
    $TestNotFoundOut = test-path $NotFoundOut
    $TestPhotoSameOut = test-path $PhotoSameOut



#if the test above is successful add them to an array
    #we will use this to add the attachments to the email in a bit
     if ($TestDuplicateHexOut -eq $true)
        {
        $attachments += $DuplicateHexOut
        }
 
     if ($testUsersChangedOut  -eq $true)
        {
        $attachments += $UsersChangedOut
        }

 
     if ($TestNotFoundOut-eq $true)
        {
        $attachments +=  $NotFoundOut
        }
      if ($TestPhotoSameOut-eq $true)
        {
        $attachments +=  $PhotoSameOut
        }


#Management Report

# Create a DataTable
$table = New-Object system.Data.DataTable "Table"
$col1 = New-Object system.Data.DataColumn Name,([string])
$col2 = New-Object system.Data.DataColumn Number,([string])
$table.columns.add($col1)
$table.columns.add($col2)

$row = $table.NewRow()
$row.Name = "Users Photo Changed"
$row.Number = $UsersChangedCount
$table.Rows.Add($row)

$row = $table.NewRow()
$row.Name = "Users whos existing photo matched the one in csv"
$row.Number = $PhotoSameCount
$table.Rows.Add($row)

$row = $table.NewRow()
$row.Name = "Users who matched a photo hex in this run"
$row.Number = $DuplicateHashCount
$table.Rows.Add($row)

$row = $table.NewRow()
$row.Name = "Users Not found"
$row.Number = $usersNotFound
$table.Rows.Add($row)

# Create an HTML version of the DataTable
$html = "<table><tr><td>Name</td><td>Number</td></tr>"
foreach ($row in $table.Rows)
{ 
    $html += "<tr><td>" + $row[0] + "</td><td>" + $row[1] + "</td></tr>"
}
$html += "</table>"


#Send the report email
$body = "Here is the results from the last run of the $Domain " + $MyInvocation.MyCommand + " script which ran at $date :<br /><br />" + $html      
send-mailmessage -to $ManagmentReportsDL -from $sender -subject "AD Photo Script Report ($Domain)"  -body $Body -SmtpServer $SmtpServer -cc $CC -BodyAsHtml   -Attachments $attachments

# Delete attachments


     if ($TestDuplicateHexOut -eq $true)
        {
        remove-item $DuplicateHexOut  -force
        }
 
     if ($testUsersChangedOut  -eq $true)
        {
        remove-item $UsersChangedOut -force
        }

 
     if ($TestNotFoundOut-eq $true)
        {
        remove-item $NotFoundOut  -force
        }
    if ($TestPhotoSameOut-eq $true)
        {
        remove-item $PhotoSameOut  -force
        }
#Stop transcript is the last thing we need to do
Stop-Transcript

