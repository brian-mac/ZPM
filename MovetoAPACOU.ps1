

Function CloseGracefully()
{
    # Close all file streams, files and sessions.
    $Stream.writeline( $Date +  " PSSession and log file closed.")
    Write-Host $Date  +  " PSSession and log file closed."
    $Stream.Close()
    $fs.Close()
    $StreamMove.close()
    $fsMove.close()
    
    # Close PS Sessions
    Get-PSSession | Remove-PSSession
    $error.clear()
    Exit
}

function ConnectToTalent2Asia ()
{
    $ADcred = Get-Credential  -Message "Please enter your Talent2Asia.com user name and password"
    Try
    {
        #Set-Item wsman:localhost\client\trustedhosts sg01svr01.talent2asia.com
        $ADSession = New-PSSession -ComputerName sg01svr01.talent2asia.com -Credential $ADcred 
        Import-PSSession $ADSession  -ErrorAction Continue
         #Import-Module activedirectory
    }
    Catch
    {
        WriteLine " : Could not create session to Allegisgroup.com.  Incorrect users details?"
    }
}
Function ConnectToO365 ()
{
    $usercredential = Get-Credential #-UserName "jkontoni.admin@allegisgroup.com" -Message "Please enter:" 
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
      #Create A session specific for the invoke commands
      $Global:Invsession = Get-PSSession  -InstanceId (Get-OrganizationConfig).RunspaceId.Guid
}

function WriteLine ($LineTxt) 
{
    $Date = get-date -Format G
    $Date = $Date + "    : "  
    $LineTxt = $date + $LineTxt  
    $Stream.writeline( $LineTxt )
}

function WriteMove ($Linetxt)
{
    $StreamMove.writeline( $LineTxt )
}

Function MoveADUser ($CheckMailbox)
{
    $ADUser = Get-ADUser -Identity $CheckMailbox.UserPrincipleName
    if ($ADUser.DistinguishedName -notlike "*OU=APAC,OU=Employees,OU=Users,OU=Enterprise,DC=allegisgroup,DC=com*")
    {
        If ($ADUser.PrimarySmtpAddress -like "*allegisgroup*")
        {
            $UserType = "AG"
        }
        if ($ADUser = "AGS" -like "*allegisglobalsolutions*")
        {
            $UserType = "AGS"
        }
        if ($CheckMailbox.IsShared)
        {
            $usertype = "Shared"
        }
        Switch ( $UserType)
        {
            "Shared"    {$TargDN = "OU=APAC,OU=Shared Accounts,OU=Special Accounts,OU=Users,OU=Enterprise,DC=allegisgroup,DC=com"}
            "AGS"       {$TargDN = "OU=AGS,OU=APAC,OU=Employees,OU=Users,OU=Enterprise,DC=allegisgroup,DC=com"}
            "AG"        {$TargDN = "OU=APAC,OU=Employees,OU=Users,OU=Enterprise,DC=allegisgroup,DC=com"}
        }
        try 
        {
            Move-ADObject -Identity $ADUser -TargetPath $TargDN 
            $line = "Sucsess: $ADUser moved to OU:$TargDN"
            WriteMove $line
        }
        catch
        {
            $line = "Error: could not move $ADUser to OU:TargDN "
            WriteMove $line
        }
    }
}
# Main Code Body

# Create Log file stream
$date = get-date -Format d
$date = $date.split("/")
$date = $date.item(2) + $date.item(1) + $date.item(0)
$MoveApacUsers = "C:\temp\MoveApacUsers$date.csv"
$mode       = [System.IO.FileMode]::Create
$ModeMove   = [System.IO.FileMode]::Create
$access     = [System.IO.FileAccess]::Write
$sharing    = [IO.FileShare]::Read
$LogPath    = [System.IO.Path]::Combine("C:\temp\GoogleSourceDestination.txt")
$MoveLog    = [System.IO.Path]::Combine($MoveApacUsers)

# create the FileStream and StreamWriter objects
$fs = New-Object IO.FileStream($LogPath, $mode, $access, $sharing)
$Stream = New-Object System.IO.StreamWriter($fs)

$fsMove = New-Object IO.FileStream($MoveLog, $ModeMove, $access, $sharing)
$StreamMove = New-Object System.IO.StreamWriter($fsMove)

#Write headers for AsiaAddToGapp.csv
$MoveLine = "EmailAddress,DestinationOU"
WriteMove $MoveLine
ConnectToO365

$SourceEmails = Import-CSV "C:\temp\MigrationStatusReport.csv"
foreach ($SourceMailbox in $SourceEmails)
{
    $SourceEmail = ($SourceMailbox.Source).trim()
    $sourcemail = $sourcemail.Replace(" ",".")
    $TargMailbox = Get-Mailbox $SourceEmail -ErrorAction  SilentlyContinue
    if ($TargMailbox)
    {
        # Great Google SMTP address matched an O365 mailbox 
        
        MoveADUser $TargMailbox
    }
    else
    {
        # Lets see if we can find an O365 account that has this source as a forwading address.
        $TargMailbox = get-mailbox -Filter "$_.Forwardingsmtpaddress -eq '$SourceEmail'" -ErrorAction SilentlyContinue
        if ($TargMailbox)
        {
            MoveADUser $TargMailbox
        }
        else
        {
            # Hmm OK, lets seee if we can any recipient that has a primary SMTP address that matches.
            $TargRecp = Get-Recipient -Filter "$_.PrimarySmtpAddress -eq '$SourceEmail'" -ErrorAction SilentlyContinue
            if ($TargRecp)
            {
                # Found Something.
                If ($TargRecp.RecipientType -eq "MailContact")
                {
                    # Yep it is a contact, now we have to find the account this is a forwarder for.
                    $ContactID = $TargRecp.DistinguishedName
                    $TargMailbox = get-mailbox -Filter "$_.ForwardingAddress -eq '$ContactId'" -ErrorAction SilentlyContinue
                    If ($TargMailBox)
                    {
                        MoveADUser $TargMailbox
                    }
                    else 
                    {
                        $Line = "Error:  No Matching O365 mailbox for: $SourceEmail"
                    }
                }
            } 
        }
    }
# Loop End
}
CloseGracefully

    
            
                
                 