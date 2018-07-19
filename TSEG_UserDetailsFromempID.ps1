# Get AD User email address from employee ID

Function CreateFile ($ReqPath, $ReqMode)
{
    <# Create the FileStream and StreamWriter objects, returns Stream Object
    Append will append to an exisitng file or create a new one if it does not exisit
    Create will create a new file or overwrite an existing one
    The file will have todays date appened to it #>

    $date = get-date -Format d
    $date = $date.split("/")
    $date = $date.item(2) + $date.item(1) + $date.item(0)
    $FileParts = $ReqPath.split(".")
    $ReqPath = $FileParts.item(0) + $date +"." + $FileParts.Item(1)

    $mode       = [System.IO.FileMode]::$ReqMode
    $access     = [System.IO.FileAccess]::Write
    $sharing    = [IO.FileShare]::Read
    $FilePath    = [System.IO.Path]::Combine($ReqPath)
    $fs = New-Object IO.FileStream($FilePath, $mode, $access, $sharing)
    Return $fs
}

function WriteData ($LineTxt,$Stream) 
{
    #$Date = get-date -Format G
    #$Date = $Date + "    : "  
    #$LineTxt = $date + $LineTxt  
    $Stream.writeline( $LineTxt )
}
Function CloseGracefully($Stream,$FileSystem)
{
    # Close all file streams, files and sessions. Call this for each FileStream and FileSystem pair.
    # $Line = $Date +  " PSSession and log file closed."
    #WriteLog $Line $stream 
    Write-Host $Date  +  " PSSession and log file closed."
    $Stream.Close()
    $FileSystem.Close() 
    # Close PS Sessions
    #Get-PSSession | Remove-PSSession
    $error.clear()
    #Exit
}

# Main
$Inputfile = "C:\temp\InputempID.csv"
$EmployeeIDs = import-csv -Path $Inputfile
$OutputFile = CreateFile "C:\temp\UserEmpID.csv" "Create"
$OutPutStream = New-Object System.IO.StreamWriter($OutputFile)

$targetAtt = "EmployeeNumber"
$DataLine = "EmployeeID,EmailAddress,Department"
WriteData $DataLine $OutPutStream

foreach ($EmployeeID in $EmployeeIDs)
{
    $TargetEmpID = ($EmployeeID."$targetAtt").trim()
    $TargetUser = get-aduser -Filter '$targetAtt -eq $TargetEmpID' -Properties emailaddress, department
    $DataLine =  $TargetEmpID + "," + $TargetUser.emailaddress + "," + $TargetUser.Department
    WriteData $DataLine $OutPutStream
}

CloseGracefully $OutputStream $OutputFile
