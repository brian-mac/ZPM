<# 
.synopsis 
 This script will: Add the specified user/s to the Google MigratedGapp group.   It will log output to a log file.
Version 1.0

.Description
The script can either create act on a single user at a time or multiple users using a CSV file.
The script has two input parameters: Target User email address and SamAccountName.

.Parameter Email
The email address of the user to be migrated.

.Parameter SamAccountName
This is the SamAccounName of the user to be migrated..

.Parameter InputFile
If this parameter is present the script will read Email, O365ForwardingAddress; from a suitably headed CSV file.  A template is provided with the script.

.Outputs
System.io.FileStream Appendes to the log file C:\temp\OMigratedT2Tasks.txt.

.Example
To call this script in Windows 10 with an input file  : powershell .\AddAsiatoGapps.ps1 -inputfile 'C:\temp\MigratedUsersT2.csv'

.Example    
To call this script in windows 10 to migrate a user DL: powershell .\AddAsiatoGapps.ps1 -Email 'Brian.mcelhinney@allegisgroup' -SamAccountNAme 'Brian.mcelhinney'
#>
[CmdletBinding (DefaultParameterSetName="Set 2")]
 param (
    [Parameter(Parametersetname = "Set 1")][String]$Inputfile , 
    [Parameter(Mandatory=$True,HelpMessage="Please enter Email Address",Parametersetname = "Set 2" )][string] $Email,
    [Parameter(Mandatory=$True,HelpMessage="Please enter SamAccountName",Parametersetname = "Set 2" )][string] $SamAccountName
 )
Function CloseGracefully()
{
    $Stream.writeline( $Date +  " PSSession and log file closed.")
    Write-Host $Date  +  " PSSession and log file closed."
    $Stream.Close()
    $fs.Close()
    $error.clear()
    Exit
}
function WriteLine ($LineTxt) 
{
    $Date = get-date -Format G
    $Date = $Date + "    : "  
    $LineTxt = $date + $LineTxt  
    $Stream.writeline( $LineTxt )
}
Function AddGroup($TargetUser,$TargetGroup)
{
    Try
    {
            Add-ADGroupMember -Identity $TargetGroup -Members $TargetUser
            $Line = "$TargetUser has been added to $TargetGroup"
            Writeline $Line
    }
    Catch
    {
        $line = "Error could not add $TargetUser into $TargetGroup"
        Writeline $line
    }
}

function CheckandImportModule ($ModuleName)
{
    $Modules = Get-Module -ListAvailable
    foreach ($Module in $Modules)
    {
        if ($Module.name -eq $ModuleName)
        {
            $ModuleFlag = $true
        }
    }
    If ($ModuleFlag -ne $true)
    {
        Import-Module $ModuleName
    }
}
Function ProcessUser($MigratedUser, $SamAccountName)

{
    #get-aduser  find any gaaps groups, removegapps them,   add the right gaap group
    Try 
        {
            $TargetUser = get-aduser -identity $SamAccountName #-Filter {Emailaddress -eq $UserEmail} -Properties *
        }       
    Catch 
    {
        $Error = "$MigratedUser Does not exist, please check SamAccountName $SamAccountName"
        WriteLine $Error
    }
    AddGroup $TargetUser $Gaap 
    
}
# Create Log file stream
$mode = [System.IO.FileMode]::Append
$access = [System.IO.FileAccess]::Write
$sharing = [IO.FileShare]::Read
$LogPath = [System.IO.Path]::Combine("C:\temp\OMigratedT2Tasks.txt")


# create the FileStream and StreamWriter objects
$fs = New-Object IO.FileStream($LogPath, $mode, $access, $sharing)
$Stream = New-Object System.IO.StreamWriter($fs)

#Set Target Group
$Gaap = "APAC-Migrated-Gapps"

CheckandImportModule "ActiveDirectory"

if ($inputfile.Length -ne 0)
{
    $SMigratedUsersT2 = import-csv -Path $Inputfile
    foreach ($MigratedUser in $SMigratedUsersT2)
    {
        $SamAccountName = $MigratedUser.SamAccountName 
        $O365User = $MigratedUser.Emailaddress
        ProcessUser $O365User $SamAccountName
    }
}
else
{
    ProcessUser $email $SamAccountName
}

#Close and write stream to file
closegracefully 