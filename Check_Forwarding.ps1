[CmdletBinding (DefaultParameterSetName="Set 2")]
param (
    [Parameter(Parametersetname = "Set 1")][String]$Inputfile
)
# Check_forwarding Johns Fork
# Version 2.0.0 20170914

Function CloseGracefully()
{
    # Close all file streams, files and sessions.
    $Stream.writeline( $Date +  " PSSession and log file closed.")
    Write-Host $Date  +  " PSSession and log file closed."
    $Stream.Close()
    $fs.Close()
    
    # Close PS Sessions
    Get-PSSession | Remove-PSSession
    $error.clear()
    Exit
}

Function ConnectToO365 ()
{
    $usercredential = Get-Credential -UserName "jkontoni.admin@allegisgroup.com" -Message "Please enter:" 
    $PSsessions = Get-PSSession
    foreach ($PsSession in $PSsessions)
    {
        If ($PsSession.computername -eq  "outlook.office365.com")
        {
            $O365SessionExists = $True
        }
    }
    If ( -not ($O365SessionExists))
    {
        $ProxyAddress = [System.Net.WebProxy]::GetDefaultProxy() |select-object address
        if ($ProxyAddress.address)
        {
            $proxyOptions = New-PSSessionOption -ProxyAccessType IEConfig
            $session = new-pssession -configurationname Microsoft.exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection  -SessionOption $proxyOptions
        }
        Else
        {
           
        } $session = new-pssession -configurationname Microsoft.exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
        Import-PSSession $session  
        If (!$session)
        {
            CloseGracefully
        }
    }
}

$mode       = [System.IO.FileMode]::Create
$access     = [System.IO.FileAccess]::Write
$sharing    = [IO.FileShare]::Read
$LogPath    = [System.IO.Path]::Combine("C:\temp\ForwardingResults.csv")

# create the FileStream and StreamWriter objects
$fs = New-Object IO.FileStream($LogPath, $mode, $access, $sharing)
$Stream = New-Object System.IO.StreamWriter($fs)
$Line = "TargetUser,ForwardingSmtpAddress,ForwardingAddress"
$Stream.writeline( $line )
ConnectToO365

#$Inputfile = "C:\temp\checkforwarding.csv"
$SMigratedUsersT2 = import-csv -Path $Inputfile
foreach ($MigratedUser in $SMigratedUsersT2)
{
    $TargetUser = $MigratedUser.Allegisemail 
    $TMailBox = get-mailbox $TargetUser -errorAction silentlycontinue
    if ( $TMailBox)
    {
        $Line = $TargetUser + "," + $TMailBox.ForwardingSmtpAddress + "," +  $TMailBox.ForwardingAddress
        $Stream.writeline( $line )
        Write-Host $Line
    }
    else
    {
        $Line = $TargetUser + " Mailbox could not be found"
        $Stream.writeline( $line )
        Write-Host $Line -ForegroundColor Red
    }
    $TMailBox = $null
}
CloseGracefully
