Function CloseGracefully()
{
    # Close all file streams, files and sessions.
    $Stream.writeline( $Date +  " PSSession and log file closed.")
    Write-Host $Date  +  " PSSession and log file closed."
    $Stream.Close()
    $fs.Close()

    
    # Close PS Sessions
   # Get-PSSession | Remove-PSSession
    $error.clear()
    Exit
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

$Inputfile = "C:\temp\checkforwarding.csv"
$SMigratedUsersT2 = import-csv -Path $Inputfile
foreach ($MigratedUser in $SMigratedUsersT2)
{
    $TargetUser = $MigratedUser.Allegisemail
    $TMailBox = get-mailbox $TargetUser 
    if ( $TMailBox)
    {
        $Line = $TargetUser + "," + $TMailBox.ForwardingSmtpAddress + "," +  $TmaiMailBoxlbox.ForwardingAddress
        $Stream.writeline( $line )
        Write-Host $Line
    }
    else
    {
        $Line = $TargetUser + " Mailbox colud not be found"
        $Stream.writeline( $line )
        Write-Host $Line -ForegroundColor Red
    }
    $TmaiMailBoxlbox = $null
}
CloseGracefully
