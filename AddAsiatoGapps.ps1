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
    [Parameter(Parametersetname = "Set 1")][String]$Inputfile #, 
    #[Parameter(Mandatory=$True,HelpMessage="Please enter Email Address",Parametersetname = "Set 2" )][string] $Email,
    #[Parameter(Mandatory=$True,HelpMessage="Please enter SamAccountName",Parametersetname = "Set 2" )][string] $SamAccountName
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
    #$Date = get-date -Format G
    #$Date = $Date + "    : "  
    #$LineTxt = $date + $LineTxt  
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
        $Message = ($_.Exception.Message).ToString()
        $line = $TargetUser.name  + "  " + $Message 
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
Function Check_Object ( $TargetUser)
{
    $TargetUser = get-adobject -Filter {mail -eq $MigratedUser} -Properties *
    If ($TargetUser) 
    {
        AddGroup $TargetUser $Gaap
    }
    else
    {
        $MigratedUser = $MigratedUser.split("@").item(0)
        $MigratedUser = $MigratedUser + "@talent2.com"
        $TargetUser = get-adobject -Filter {mail -eq $MigratedUser} -Properties *
        If ($TargetUser) 
        {
            AddGroup $TargetUser $Gaap
        }
        else
        {
            $SAM = $MigratedUser.split("@").item(0)
            $TargetUser = get-adobject -Filter {SamAccountName -eq $SAM} -Properties *
            If ($TargetUser) 
            {
                AddGroup $TargetUser $Gaap
            }
            else
            {
                $Error = "$Sam Does not exist, please check SamAccountName $SamAccountName"
                WriteLine $Error
            }   
        }
    }
}
Function ProcessUser($MigratedUser, $CriteriaFlag)

{
    #get-aduser  find any gaaps groups, removegapps them,   add the right gaap group
    if ($CriteriaFlag -ne 2)
    {
        $TargetUser = get-adobject -Filter {mail -eq $MigratedUser} -Properties * -erroraction stop #get-aduser -Filter {Emailaddress -eq $UserEmail} -Properties *  
    }
    Else
    {
        $TargetUser = Get-aduser $MigratedUser
    }
    If ($TargetUser) 
    {
        $Cast = $TargetUser.gettype()
        if ($Cast.basetype.name -eq "ADentity" -or $Cast.basetype.name -eq "ADAccount")
        {
            # We have found a valid user 
            AddGroup $TargetUser $Gaap    
        } 
        else
        {
            foreach ($TargUser in $TargetUser) # ($i=0; $i -le $targetuser.count(); $i++)
            {
                ProcessUser $TargUser.name  
            }
        }
    }
    else
    {
        if ($CriteriaFlag -eq 0)
        {   $MigratedUser = $MigratedUser.split("@").item(0)
            $MigratedUser = $MigratedUser + "@talent2.com"
        }
        if ($CriteriaFlag -eq 1)
        {
            $MigratedUser = $MigratedUser.split("@").item(0)
        }
        $CriteriaFlag++
        if ($CriteriaFlag -lt 3)
        {
            ProcessUser $MigratedUser $CriteriaFlag
        }
        else
        {
            $Error = "$MigratedUser Does not exist, please check SamAccountName $SamAccountName"
            WriteLine $Error
        }    
    }   
}
# Create Log file stream
$mode = [System.IO.FileMode]::Append
$access = [System.IO.FileAccess]::Write
$sharing = [IO.FileShare]::Read
$LogPath = [System.IO.Path]::Combine("C:\temp\AddGroups.csv")


# create the FileStream and StreamWriter objects
$fs = New-Object IO.FileStream($LogPath, $mode, $access, $sharing)
$Stream = New-Object System.IO.StreamWriter($fs)

WriteLine "Status"
#Set Target Group
$Gaap = "APAC-Migrated-Gapps"

# CheckandImportModule "ActiveDirectory"
$Inputfile = "C:\temp\MigrationUsers.csv"
if ($inputfile.Length -ne 0)
{
    $SMigratedUsersT2 = import-csv -Path $Inputfile
    foreach ($MigratedUser in $SMigratedUsersT2)
    {
        #$SamAccountName = $MigratedUser.SamAccountName 
        $O365User = $MigratedUser.SourceEmail
        ProcessUser $O365User 0
        $Stream.flush()
    }
}
else
{
    ProcessUser $email 0
}

#Close and write stream to file
closegracefully 