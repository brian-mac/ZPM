function SecurePassword ($Target_Path, $Action)
{
    <# Function will either encrypt and save a password to the file specified in the $Target_Path
    Or it will return an encrypted password from the file specified in $Target_Path.
    Returns password 
    #>

    If ($Action -eq "Store")
    {
        $Secure = Read-Host -AsSecureString
        $Encrypted = ConvertFrom-SecureString -SecureString $Secure -Key (244,102,80,104,223,19,65,130,183,11,132,245,74,147,46,142)
        $Encrypted | Set-Content $Path  
    }
    elseif ($Action -eq "Retreive")
    {
        $Secure = Get-Content $Target_Path | ConvertTo-SecureString -Key (244,102,80,104,223,19,65,130,183,11,132,245,74,147,46,142)
    }
    Return $Secure
}
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
            $Global:ExSession = new-pssession -configurationname Microsoft.exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection  -SessionOption $proxyOptions
        }
        Else
        {
            $Global:ExSession = new-pssession -configurationname Microsoft.exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
        } 
        Import-PSSession $ExSession  # -Prefix OnPrem only requires if dual Exch O365 sessions
        If (!$ExSession)
        {
            Exit 
        }
    }
}

