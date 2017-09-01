 <# 
.synopsis 
This script will set Out Of Office for O365 mailboxes. In addition it sets the DeliverToMailboxAndForward to true
when Scheduled and false when Disabled.

.Description
The script will set the internal and external message for out of office, default behaviour is to reply externally and internally.
It can be run to update on mailbox, or with the -inputFile parameter to update multiple mailboxes.

.Parameter  EmailAddress
The email address of the mailbox to set OOO.

.Parameter AutoState
This set the OOO to enabled or disabled.  Values are "Scheduled" or "Disabled"

.Parameter StartDate
Scheduled start time of OOO, in short date and time format, in quotes "06/06/2017 8:00 PM"

.Parameter EndDate
Scheduled end date of OOO, in short date and time format, in quotes "06/06/2017 8:00 PM"

.Parameter IntMessage
Internal Message.

.Parameter Extmessage
External Message.

.Parameter InputFile
The path to the CSV file that contains the list of mailboxes to be actioned.

.Outputs
System.io.FileStream Appendes to the log file OOO.log in the %AppsData% directory.

.Example
To call this script in Windows 10 with an input file, in the same directory as the script type: powershell .\OutOfOffice -inputfile 'C:\temp\OOOUpdate.csv'

.Example    
To call this script in windows 10 to update a single user: in the same sirectory as the script: powershell .\OutOfOffice.ps1 it will prompt you for all the required fields.
#>

#Version number 1.2.2  20170604 14:35

 [CmdletBinding (DefaultParameterSetName="Set 2")]
 param (
    [Parameter(Parametersetname = "Set 1")][String]$Inputfile , #="C:\Temp\APAC-DistListBulkCreate.CSV" ,
    [Parameter(Mandatory=$True,HelpMessage="Please enter Email Address",Parametersetname = "Set 2" )][string] $Email,
    [Parameter(Mandatory=$True,HelpMessage="Please enter AutoState",Parametersetname = "Set 2" )][string] $AutoState,
    [Parameter(Mandatory=$True,HelpMessage="Please enter the Start date",Parametersetname = "Set 2")][system.datetime] $StartDate ,
    [Parameter(Mandatory=$True,HelpMessage="Please enter the End Date",Parametersetname = "Set 2")][system.datetime] $EndDate,
    [Parameter(Mandatory=$True,HelpMessage="Please enter the internal message",Parametersetname = "Set 2")][string] $IntMessage,
    [Parameter(Mandatory=$True,HelpMessage="Please enter the external message",Parametersetname = "Set 2")][string] $Extmessage
 )
# Declare Functions to call from main body.

function LogError ($KnownError)
{
    $Date = get-date -Format G
    $Date = $Date + "    : "   
    If ($error.count -eq 0)
    {
        $error.Add("No PS Error Detected")
    }
    If ($KnownError.length -eq 0)
    {
        $KnownError = "Unexpected this was.."
    }
    $Stream.writeline( $Date + $error + "   :" + $KnownError )
    Write-Host $Date  +  $error + $KnownError 
    $Error.clear()
    $KnownError = "" 
}

Function CloseGracefully()
{
    $Stream.writeline( $Date +  " PSSession and log file closed.")
    Write-Host $Date  +  " PSSession and log file closed."
    $Stream.Close()
    $fs.Close()
    # Close PS Sessions
    Get-PSSession | Remove-PSSession
    $error.clear()
    Exit
}

Function Set-Store ($Identity,$StoreFlag)
{
    if ($StoreFlag)
    {
        Set-Mailbox -Identity $Identity -DeliverToMailboxAndForward $True
    }
    Else
    {
        Set-Mailbox -Identity $Identity -DeliverToMailboxAndForward $False
    }
}

Function Set-OOO($Identity,$AutoState,$StartTime,$EndTime,$ExtMessage,$IntMessage)
{
    # Have to convert string date to US date short format grrrr
    #$StartDateParts = $StartTime.Split("/")
    #$StartTime = $StartDateParts[1] + "/" + $StartDateParts[0] + "/" + $StartDateParts[2]
    #$EndDateParts = $EndTime.Split("/")
    #$EndTime = $EndDateParts[1] + "/" + $EndDateParts[0] + "/" + $EndDateParts[2]
    $ExtMessage = "$ExtMessage <br>"
    $IntMessage = "$IntMessage <br>"
    Try
    {
        set-MailBoxAutoReplyConfiguration -identity $Identity -AutoReplyState $AutoState -StartTime $StartTime -EndTime $EndTime -ExternalMessage $ExtMessage -InternalMessage $IntMessage -ExternalAudience:all
        $ErrLog = " : OOO set for :" + $Identity
        LogError $Errlog
        Set-Store $identity $true
    }
    Catch
    {
        $Errlog =  " : Could not set OOO for :" + $Identity
        Logerror $ErrLog
    }
}

Function Disable-OOO ($identity)
{
    Try
    {
        Set-MailboxAutoReplyConfiguration -Identity $Identity -AutoReplyState Disabled -ExternalMessage $null -InternalMessage $null
        $Errlog = " : OOO Removed for :" + $Identity
        LogError $Errlog
        Set-Store $Identity $False
    }
    Catch
    {
        $ErrLog = " : Could not remove OOO for :" + $Identity
        LogError $Errlog
    }
}

Function ConnectToO365 ()
{
    $PSsessions = Get-PSSession
    foreach ($PsSession in $PSsessions)
    {
        If ($PsSession.computername -eq  "outlook.office365.com")
        {
            $O365SessionExists = $True
        }
    }
    If ( -not $O365SessionExists)
    {
        $usercredential = Get-Credential -Message "Please enter user name and password for allegisgroup.com"
        $ProxyAddress = [System.Net.WebProxy]::GetDefaultProxy() | select-object address
        if ($ProxyAddress.address)
        {
            $proxyOptions = New-PSSessionOption -ProxyAccessType IEConfig
            $session = new-pssession -configurationname Microsoft.exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection  -SessionOption $proxyOptions
        }
        Else
        {
            $session = new-pssession -configurationname Microsoft.exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
        }
        Import-PSSession $session  
    }
}

#Define variables and constants

$ErrorActionPreference = "Stop"
# Get date time for Log file entry
$Date = get-date -Format G
$Date = $Date + "    : "
# Create Log file stream
$mode = [System.IO.FileMode]::Append
$access = [System.IO.FileAccess]::Write
$sharing = [IO.FileShare]::Read
$LogPath = [System.IO.Path]::Combine($Env:AppData,"OutOfOffice.log")
# create the FileStream and StreamWriter objects
$fs = New-Object IO.FileStream($LogPath, $mode, $access, $sharing)
$Stream = New-Object System.IO.StreamWriter($fs)

# Create Remote Exchange Session

ConnectToO365

#$Inputfile = "C:\temp\ooo.csv" # Debug

if($Inputfile.Length -ne 0)  # Was a input file selected, are we processing one mailbox manualy or many from a file.
{
    $Mailboxes = import-csv -Path $Inputfile
    foreach ($mailbox in $mailboxes)
    {
        If ($AutoState -eq "Disabled")
        {
            Disable-OOO $mailbox.Email
        }
        Else
        {
            Set-OOO $Mailbox.Email $Mailbox.AutoState $Mailbox.StartDate $Mailbox.EndDate $Mailbox.ExtMessage $Mailbox.IntMessage
        }
    }
}
else
{
    If ($AutoState -eq "Disabled")
    {
        Disable-OOO $Email
    }
    Else
    {
        Set-OOO $Email $AutoState $StartDate $EndDate $ExtMessage $IntMessage
    }
}
CloseGracefully
