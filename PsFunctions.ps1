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
    $StartDateParts = $StartTime.Split("/")
    $StartTime = $StartDateParts[1] + "/" + $StartDateParts[0] + "/" + $StartDateParts[2]
    $EndDateParts = $EndTime.Split("/")
    $EndTime = $EndDateParts[1] + "/" + $EndDateParts[0] + "/" + $EndDateParts[2]
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
        $ProxyAddress = [System.Net.WebProxy]::GetDefaultProxy() |select-object address
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