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
connecttoo365
