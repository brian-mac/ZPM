$proxyOptions = New-PSSessionOption -ProxyAccessType IEConfig
$usercredential = Get-Credential
$session = new-pssession -configurationname Microsoft.exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection  -SessionOption $proxyOptions
Import-PSSession $Session

