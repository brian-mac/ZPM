$cred = Get-Credential
#$proxyOptions = New-PSSessionOption -SkipRevocationCheck
$sessionOption = New-PSSessionOption -ProxyAccessType IEConfig 
$ExSession   = New-PSSession -Authentication basic -Credential $cred -ConnectionUri https://outlook.allegisgroup.com/PowerShell/ -ConfigurationName Microsoft.Exchange -AllowRedirection  -SessionOption $sessionOption 
Import-PSSession $ExSession 
 