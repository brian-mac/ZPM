[CmdletBinding (DefaultParameterSetName="Set 1")]
param (
    [Parameter(Mandatory=$True,HelpMessage="Please enter name for Room input file",Parametersetname = "Set 1" )][string] $RoomFile
)

Function CreateFile ($ReqPath, $ReqMode)
{   
    # Create the FileStream and StreamWriter objects, returns Stream Object
    $date = get-date -Format d
    $date = $date.split("/")
    $date = $date.item(2) + $date.item(1) + $date.item(0)
    $FileParts = $ReqPath.split(".")
    $ReqPath = $FileParts.item(0) + $date +"." + $FileParts.Item(1)

    $mode       = [System.IO.FileMode]::$ReqMode
    $access     = [System.IO.FileAccess]::Write
    $sharing    = [IO.FileShare]::Read
    $LogPath    = [System.IO.Path]::Combine($ReqPath)
    $fs = New-Object IO.FileStream($LogPath, $mode, $access, $sharing)
    Return $fs
}
function WriteLine ($LineTxt,$Stream) 
{
    $Date = get-date -Format G
    $Date = $Date + "    : "  
    $LineTxt = $date + $LineTxt  
    $Stream.writeline( $LineTxt )
}
Function CloseGracefully($Stream,$FileSystem)
{
    # Close all file streams, files and sessions. Call this for each FileStream and FileSystem pair.
    $Line = $Date +  " PSSession and log file closed."
    WriteLine $Line $stream 
    Write-Host $Date  +  " PSSession and log file closed."
    $Stream.Close()
    $FileSystem.Close() 
    # Close PS Sessions
    Get-PSSession | Remove-PSSession
    $error.clear()
    Exit
}
Function ConnectToExch ($Username, $Target_Path)
{
    <# Connects to a remote Exchange enviroment.  If a UserName is passed it will look in $Target_Path
       for the Encrypted password and use those for credentials. Otherwise it will prompt interactivley
       If a vild session exists it will not create a new session 
    #>

    $PSsessions = Get-PSSession
    foreach ($PsSession in $PSsessions)
    {
        If ($PsSession.computername -eq  "outlook.allegisgroup.com" -and $PsSession.State -ne "Broken")
        {
            $ExSessionExists = $True
        }
    }
    If ( -not ($ExSessionExists))
    {
       if ($Username)
       {
        $SecPass = SecurePassword $Target_Path "Retreive"
        $UserCredential = New-Object System.Management.Automation.PSCredential($Username,$SecPass)
       } 
       Else
       {
        $UserCredential = Get-Credential  -Message "Please enter you credentials for Remote Exchange:" 
       }
       
        $ProxyAddress = [System.Net.WebProxy]::GetDefaultProxy() |select-object address
        if ($ProxyAddress.address)
        {
            $proxyOptions = New-PSSessionOption -ProxyAccessType IEConfig
            $Global:ExSession = new-pssession -configurationname Microsoft.exchange -ConnectionUri https://outlook.allegisgroup.com/powershell/ -Credential $UserCredential -Authentication Basic -AllowRedirection  -SessionOption $proxyOptions
        }
        Else
        {
            $Global:ExSession = new-pssession -configurationname Microsoft.exchange -ConnectionUri https://outlook.allegisgroup.com/powershell/ -Credential $UserCredential -Authentication Basic -AllowRedirection
        } 
        Import-PSSession $ExSession   -Prefix OnPrem 
        If (!$ExSession)
        {
            Exit 
        }
    }
}

Function ConnectToO365 ($Username, $Target_Path)
{
    <# Connects to O365 OnLine enviroment.  If a UserName is passed it will look in $Target_Path
       for the Encrypted password and use those for credentials. Otherwise it will prompt interactivley
       If a vild session exists it will not create a new session 
    #>

    $PSsessions = Get-PSSession
    foreach ($PsSession in $PSsessions)
    {
        If ($PsSession.computername -eq  "outlook.office365.com" -and $PsSession.State -ne "Broken")
        {
            $ExSessionExists = $True
        }
    }
    If ( -not ($ExSessionExists))
    {
       if ($Username)
       {
        $SecPass = SecurePassword $Target_Path "Retreive"
        $UserCredential = New-Object System.Management.Automation.PSCredential($Username,$SecPass)
       } 
       Else
       {
        $UserCredential = Get-Credential  -Message "Please enter you credentials for Remote Exchange:" 
       }
       
        $ProxyAddress = [System.Net.WebProxy]::GetDefaultProxy() |select-object address
        if ($ProxyAddress.address)
        {
            $proxyOptions = New-PSSessionOption -ProxyAccessType IEConfig
            $Global:O365Session = new-pssession -configurationname Microsoft.exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection  -SessionOption $proxyOptions
        }
        Else
        {
            $Global:O365Session = new-pssession -configurationname Microsoft.exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
        } 
        Import-PSSession $O365Session  # -Prefix OnPrem only requires if dual Exch O365 sessions
        If (!$O365Session)
        {
            Exit 
        }
    }
}

# Main Code
<#Connect to Exchange, pass the UserName and Path to the encrypted password.
  If no parrameters passed, you will be prompted for them
#>
# ConnectToExch # "Brian.mcelhinney@allegisgroup.com" "C:\temp\e.txt"  
ConnectToO365 # "Brian.mcelhinney@allegisgroup.com" "C:\temp\e.txt"
<# Create the log file specifed from the input parameter.  
   Append will append to an exisitng file or create a new one if it does not exisit
   Write will create a new file or overwrite an existing one
   The file will have todays date appened to it
#>
$LogFile = CreateFile "C:\temp\RoomLog.txt"  "Append"
$LogStream = New-Object System.IO.StreamWriter($LogFile)

$Rooms = import-csv $RoomFile

foreach ($Room in $Rooms)
{
    $RoomName = $Room.'Room Name'
    $Command = $RoomName + ":\calendar"
    try
    {
        Invoke-Command -Session $O365session -ScriptBlock {Set-MailboxFolderPermission -AccessRights LimitedDetails -Identity $Using:Command -User default} -ErrorAction Stop
        $Line = "Sucsess: Limited details addedd for $RoomName"
    }
    Catch 
    {
        $Line = "Error: Limited details not addedd for $RoomName"
    }
    WriteLine $line $LogStream
    try
    { Invoke-Command -Session $O365session -ScriptBlock {Set-CalendarProcessing -Identity $using:RoomName -AddOrganizerToSubject $true -DeleteComments $false -DeleteSubject $false}    -ErrorAction Stop
        $Line = "Sucsess: Organiser and Subject visiable for $RoomName"
    }
    Catch 
    {
        $Line = "Sucsess: Organiser and Subject visiable for $RoomName"
    }
    WriteLine $line $LogStream
}





CloseGracefully $LogStream $LogFile
