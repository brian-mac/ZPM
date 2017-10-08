

$ADcred = Get-Credential  -Message "Please enter your Talent2Asia.com user name and password"
Try
{
    #Set-Item wsman:localhost\client\trustedhosts sg01svr01.talent2asia.com
    $ADSession = New-PSSession -ComputerName APAURHDCPRDV01.allegisgroup.com -Credential $ADcred 
    Import-PSSession $ADSession  -ErrorAction Continue
     #Import-Module activedirectory
}
Catch
{
    WriteLine " : Could not create session to Allegisgroup.com.  Incorrect users details?"
}