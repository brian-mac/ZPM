Function AddDeligations ($GoogleUPN,$O365Specific)
{
    #Check to see if current mailbox has a dependcey.
    
    If (!$O365Specific)
    {
        # This means there will be a discrepencey between O365 account and GoogleUPN.
        # Check for dependecies using Google UPN.
        If ($hash[$GoogleUPN])
        {
            # Mailbox delegation found, however we need to use O365 specific value to bind to O365 mailbox
            $Target = get-mailbox -identity $O365Specific
            $HashValue = $Hash[$GoogleUPN]
            
        }
    }
   
}

Function ConnectToO365 ()
{
    $usercredential = Get-Credential # -UserName "jkontoni.admin@allegisgroup.com" -Message "Please enter:" 
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

#ConnectToO365

$Hash=@{}
$DependFile ="C:\temp\dependacyreport.csv"
$hash =import-csv -Path $DependFile
AddDeligations "brian.mcelhinney@allegisgroup.com" ""


