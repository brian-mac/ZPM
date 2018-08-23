<# 
.synopsis 
 This script will: 
*Find all users who match the attributes in the Inputfile.
*It will then list these users in the output file.
*At this stage the attribute is Department, but this script can easyly be modified to accept the attribute as a input parameter

.Description
Version 1.1.0
Find all users who match the attributes in the Inputfile.
*It will then list these users in the output file.

.Parameter InputFile
Optional the file name and location that stores the attribute to match.


.Outputs
System.io.FileStream Appendes to the log file C:\temp\UserDetails.log.
System.io.FileStream Creates a new file       C:\temp\InputFileNameResults.csv.

.Example
To call this script with an input file  : powershell .\TSEG_UserDetailsFromDept.ps1 -inputfile 'C:\temp\exampledept.csv'
#>
[CmdletBinding (DefaultParameterSetName="Set 2")]
param (
    [Parameter(Mandatory=$false,HelpMessage="File path and name",Parametersetname = "Set 2" )][string] $Inputfile
)
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
 #$inputfile = "C:\temp\DepartmentsFoodandBev1.csv"
if ($Inputfile.Length -eq 0) 
{
    $InputFile = read-host -prompt  "Please enter Filname and path e.g. C:\temp\DepartmentsFoodandBev1.csv"
}
$logPath = "C:\temp\DetailsFromDept.log" 
$path = test-path C:\Temp
if ($path -eq $false)
{
    New-Item C:\temp -ItemType Directory 
}
$LogFile = CreateFile $LogPath "Create"
$LogStream = New-Object System.io.StreamWriter($LogFile)
try
{
    $Departments = import-csv -Path $Inputfile
}
catch 
{
    WriteData "Error with input file, please check path" $LogStream  
    exit     
}
$Outpath = $Inputfile.split(".").item(0)
$OutPath = $Outpath + "Result.csv" 
$OutputFile = CreateFile  $OutPath "Create"
$OutPutStream = New-Object System.IO.StreamWriter($OutputFile)
$targetAtt = "Department"
$DataLine = "EmployeeNumber,EmailAddress,Department,Enabled"
WriteData $DataLine $OutPutStream
foreach ($Department in $Departments)
{
    Try
    {
        $TargetDept = ($Department."$targetAtt").trim()
        $TargetUsers = get-aduser -Filter '($targetAtt -eq $TargetDept) -and (EmailAddress -ne "*") -and (Enabled -eq $true)  ' -Properties  EmployeeNumber, department, EmailAddress, enabled
        foreach($TargetUser in $TargetUSers)
        {
            #Get-ADUser $targetuser -Properties emailaddress, department
            $DataLine =  $TargetUser.EmployeeNumber + "," + $TargetUser.emailaddress + "," + $TargetUser.Department + "," + $TargetUser.Enabled
            WriteData $DataLine $OutPutStream
        }
    }
    Catch 
    {
        WriteLine "No Users could be fond for that department, please check department name" $LogFile
    }
}
CloseGracefully $OutputStream $OutputFile
CloseGracefully $LogStream $LogFile
