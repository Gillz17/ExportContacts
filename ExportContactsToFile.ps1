#Checks to see if the value is true or false, used to check if the values are emptyFunction IIf($If, $IfTrue, $IfFalse) {    If ($If) {If ($IfTrue -is "ScriptBlock") {&$IfTrue} Else {$IfTrue}}    Else {If ($IfFalse -is "ScriptBlock") {&$IfFalse} Else {$IfFalse}}}

#Outlook contacts to loop through
$csvPath = 'H:\Support\Contacts.CSV'
$name = @()

#Loops through CSV file
Import-Csv -Path $csvPath | Foreach-Object{
    #Checks to see if the value exists
    $Title = IIF $_."Title" -eq "" $_."Title" ""
    $firstName = IIF $_."First Name" -eq "" $_."First Name" ""
    $middleName = IIF $_."Middle Name" -eq "" $_."Middle Name" ""
    $lastName = IIF $_."Last Name" -eq "" $_."Last Name" ""
    
    #If the value exists add it to the file name
    $name = (IIF $Title -eq "" "$Title" "") + (IIF $firstName -eq "" "$firstName" "") + 
        (IIF $middleName -eq "" " $middleName " "") + (IIF $lastName -eq "" "$lastName" "")
    $name = $name.Replace("\","-").Replace("/","-").Replace("?","-")
    
    #Assign the notes variable the contents of the notes column
    $notes = $_.notes

    #Check if name is empty, if so assign the company instead
    if($name -eq ""){
        $name = $_."Company"
    }

    #Check if the name is empty, if not create the file and write the note to the file
    If($name -ne ""){
        $outputPath = "C:\Documents\$name.rtf"
        Set-Content -Path $outputPath -Value ($notes)
    }else{
        #Write the CSV object to the screen if name is empty
        Write-Host "Error Creating file: $_"
    }
}