

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
#$MoveLine = "EmailAddress,DestinationOU"
#WriteMove $MoveLine
ConnectToO365

$SourceEmails = Import-CSV "C:\temp\delta.csv"
foreach ($SourceMailbox in $SourceEmails)
{
    $SourceEmail = ($SourceMailbox.Source).trim()
    #$sourcemail = $sourcemail.Replace(" ",".")
    $TargMailbox = Get-Mailbox $SourceEmail -ErrorAction  SilentlyContinue
    if ($TargMailbox)
    {
        # Great Google SMTP address matched an O365 mailbox 
        $Line =  "Sucsess:   Google: " + $SourceEmail + " : Matches O365: UPN:" +  $targMailbox.UserPrincipalName + " : PrimarySmtpAddress:" + $TargMailbox.PrimarySmtpAddress
        WriteLine   $Line # need a cant find error some where 
    }
    else
    {
        # Lets see if we can find an O365 account that has this source as a forwading address.
        $TargMailbox = get-mailbox -Filter "$_.Forwardingsmtpaddress -eq '$SourceEmail'" -ErrorAction SilentlyContinue
        if ($TargMailbox)
        {
            $Line =  "Sucsess:   Google: " + $SourceEmail + " : Matches O365: UPN:" +  $targMailbox.UserPrincipalName + " : PrimarySmtpAddress:" + $TargMailbox.PrimarySmtpAddress
            WriteLine   $Line #need a cant find error some where
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
                        # Ok we found a mailbox that forwards to the Google SMTP address.
                        $Line =  "Sucsess:   Google: " + $SourceEmail + " : Matches O365: UPN:" +  $targMailbox.UserPrincipalName + " : PrimarySmtpAddress:" + $TargMailbox.PrimarySmtpAddress
                        WriteLine $Line
                    }
                    else 
                    {
                        $Line = "Error:  No Matching O365 mailbox for: $SourceEmail"
                        WriteLine $Line
                    }
                }
            }
            else
            {
                $Line = "Error:  No Matching O365 mailbox for: $SourceEmail"
                WriteLine $Line
            } 
        }
    }
# Loop End
}
CloseGracefully

    
            
                
                 