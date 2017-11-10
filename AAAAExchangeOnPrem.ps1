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
        Import-PSSession $ExSession  # -Prefix OnPrem 
        If (!$ExSession)
        {
            Exit 
        }
    }
}
ConnectToExch



