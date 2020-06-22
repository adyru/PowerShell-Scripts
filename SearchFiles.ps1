<#
get all files on a computer
and then se
#>

#Variables to change as required
$date = get-date -Format dd-MM-yyyy-HH-mm-ss
$output = $PSScriptRoot + "\" + $Date + "-output.txt"
$errors = $PSScriptRoot + "\" + $Date + "-errors.txt"
$search = "bdcadomc"
$searchFolder = "C:\Users\adria\Documents\"

write-host -ForegroundColor green " Search will be for " $search " in " $searchFolder

#get files - exclude output,errors and scriptname so that it will only show ones we are intertested in
$files = Get-ChildItem  $searchFolder -Recurse -File -ErrorAction SilentlyContinue -Exclude $output,$errors,$MyInvocation.MyCommand.Name


foreach ($file in $files)
#search through the files
    {
    #write-host "file is "$file
        try
            #try to look for the patten in a file and output it to file
            {Select-String $file.fullname -Pattern $search -ErrorAction Stop | add-content $output}
        
   catch [System.ArgumentException] 
            {
            #if we get the locked file error output to host and write to file
            Write-host -ForegroundColor yellow "error on file " $file.fullname " probably locked, will out put to errors file"
            $file.fullname | add-content $errors
            }
    CATCH [System.Management.Automation.WildcardPatternException]
            {
            Write-host -ForegroundColor yellow "error on file " $file.fullname " generally because the file cant ne opened - will put in errorfile"
            $file.fullname | add-content $errors
            }
            # we arent going to carry on with anything this time but might as well keep things as they should be in case i expand this
    Finally
            {
            }
    }
