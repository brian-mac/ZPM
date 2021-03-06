######################################################################################################
#provisionusers.ps1                                                                                  #
#version 1.0 Created by Sean McGovern 3/18/2014                                                      #
#version 1.1 corrections and testing done -sm 3/19/2014                                              #
#                                                                                                    #
#Description: for talent 2 merge will create accounts, mailbox, contact, and forward based on CSV    #
######################################################################################################
#Start-Transcript -path C:\scripts\provisionusers.log -Append
$userfile = Import-Csv C:\scripts\talent2_new.csv
$newusers = @{}
$targetserver="rp-int-exdc1.allegisgroup.com"
$outputfile = "c:\scripts\talent2-3.31.14.csv"
[array]$outputcsv = '"First name","Last Name","UID","Password","primary smtp"'
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin
function create-contact($userid,$smtpaddress,$firstname,$lastname) #will create a contact to forward internal email to talent 2 employees
{
    $contactou="OU=Talent2,OU=Forwarding,OU=Contacts,OU=Accounts,DC=allegisgroup,DC=com"
    $displayname=$lastname + ", " + $firstname
    $userprincipalname=$userid
    $userid=$userid.substring(0,$userid.indexof('@'))
    $contactname=$userid + ".T2"
    $time=get-date
    $time.ToString() + ":new-mailcontact called for " + $smtpaddress
    New-MailContact -ExternalEmailAddress $smtpaddress -Name $contactname -DisplayName $displayname -FirstName $firstname -LastName ($lastname + ".T2") -OrganizationalUnit $contactOU -DomainController $targetserver -Alias $contactname
    $time=get-date
    $time.ToString() + ":set-mailbox called for " + $userid + " forward to " + $smtpaddress
    Set-Mailbox -Identity $userprincipalname -DeliverToMailboxAndForward $false -ForwardingAddress $smtpaddress -DomainController $targetserver
    $aduser=get-aduser $userid -Server $targetserver
    $time=get-date
    $time.ToString() + ":set-mailcontact called for " + $contactname
    Set-MailContact $contactname -HiddenFromAddressListsEnabled $true -CustomAttribute7 $aduser.ObjectGUID -DomainController $targetserver
    #$time=get-date
    #$time.ToString() + ":set-contact called for " + $contactname
    #Set-Contact  $contactname -LastName $lastname -DomainController $targetserver
}
function create-mailbox($userid) #will mail enable the userid it is given
{
$xmlfile="c:\scripts\mbxconfig.xml"
[xml]$xmlsettings = gc $xmlfile
$i=[int]$xmlsettings.Exchange.LastServerIndex
$mbtotal=$xmlsettings.Exchange.Servers.ExchangeServer.Count
$i++
if ($i -ge $mbtotal){$i=0}
$cnstring=$xmlsettings.exchange.servers.exchangeserver[$i].TrimStart("CN=")
$firstcomma=$cnstring.indexof(",")
$databasename=$cnstring.substring(0,$firstcomma)
$storagegroup=$cnstring.substring($databasename.Length + 4)
$firstcomma=$storagegroup.indexof(",")
$storagegroup=$storagegroup.Substring(0,$firstcomma)
$lastcomma=$cnstring.LastIndexOf(",")
$servername=$cnstring.Substring($lastcomma+1).trimstart("CN=")
$dbstring=$servername + ".allegisgroup.com\" + $storagegroup + "\" + $databasename
$time=get-date
$time.ToString() + ":enable-mailbox called for " + $userid + " on server entry " + $i
Enable-Mailbox -Database $dbstring -Identity $userid -DomainController $targetserver
$xmlsettings.Exchange.LastServerIndex=[string]$i
$xmlsettings.Exchange.MBXTotal=[string]$mbtotal
$xmlsettings.save($xmlfile)
}
function get-password() #will return a password based on settings indide function
{
$nonambiguous = "a","b","c","d","e","f","g","h","k","m","n","p","r","s","t","w","x","z","A","B","C","D","E","F","G","H","J","K","L","M","N","P","Q","R","T","W","X","Y"
$lowercase = "a","b","c","d","e","f","g","h","k","m","n","p","r","s","t","w","x","z"
$uppercase = "A","B","C","D","E","F","G","H","J","K","L","M","N","P","Q","R","T","W","X","Y"
$chartotal=6
$numtotal=1
$pwd=$null
do
{
    $pwd=$null
    $i=0
    do
    {
        $pwd=$pwd + [string](Get-Random -InputObject $nonambiguous)
        $i++
    }
    while ($i -le $chartotal)
    $ucasefound=$false
    $lcasefound=$false
    $chararray=$pwd.ToCharArray()
    foreach ($pwdchar in $chararray)
    {
       if($lowercase.Contains([string]$pwdchar))
       {
        $lcasefound=$true
       }
       if($uppercase.Contains([string]$pwdchar))
       {
        $ucasefound=$uppercase.Contains([string]$pwdchar)
       }
    }
    $i=0
    do
    {
        $pwd=$pwd + [string](Get-Random -Minimum 0 -Maximum 9)
        $i++
    }
    while ($i -le $numtotal)
    #if(-not ($lcasefound -band $ucasefound)){"it works"}
}
while (-not ($lcasefound -band $ucasefound))
return $pwd
}
foreach ($user in $userfile) #cycles through each user in the file
{
    $givenname=$user.'User First Name' #this section initially populates useful fields
    $surname=$user.'User Last Name'
    $fullname =$givenname + " " + $surname
    $usersearch=$null
    $i=1
    do #this loop finds a username not in use for us to use
    {
        $uid=$givenname.replace(" ","").Substring(0,$i) + $surname.replace(" ","")
        $uid.rep
        if ($uid.Length -ge 9)
        {
           $uid=$uid.Substring(0,8)
        }
        $usersearch=Get-ADUser -Filter {sAMAccountName -eq $uid} -Server $targetserver
        if ($usersearch -eq $null -band -bnot $newusers.ContainsKey($uid))
        {
            $i=100
            #$uid + " available"
        }
        else
        {
            $i++
            #$uid + " exists"
        }
    }
    while ($i -le $givenname.Length)
        if ($i -ne 100) #previous loop should return a valid userid to use which we then process
        {
            "Unable to find an available account for " + $user
            $uid = "unable to find account for " + $givenname + " " + $surname
        }
        else
        {
        $emailaddress=$uid + "@allegisglobalsolutions.com"
        $UserPrincipalName = $uid + "@allegisglobalsolutions.com"
        $OUpath = "OU=User Accounts,OU=Users,OU=Accounts,DC=allegisgroup,DC=com"
        #$newuserPW = "Ch@ngeMe2Day" 
        $newuserPW = get-password
        if ($uid.length -le 15) 
        {
            $time=get-date
            $time.ToString() + ":new-aduser called for " + $UserPrincipalName
            #New-ADUser -SamAccountName $uid -UserPrincipalName $UserPrincipalName -GivenName $givenname -Surname $surname -DisplayName ($surname + ", " + $givenname) -Name $fullname -Enabled $true -EmailAddress $emailaddress -Path $OUpath -AccountPassword (ConvertTo-SecureString $newuserPW -AsPlainText -Force) -ChangePasswordAtLogon $false -OtherAttributes @{Extensionattribute11="1";Extensionattribute3="AGST2";Extensionattribute2=$user.Division;physicalDeliveryOfficeName=$user.Location} -server $targetserver
            New-ADUser -SamAccountName $uid -UserPrincipalName $UserPrincipalName -GivenName $givenname -Surname $surname -DisplayName ($surname + ", " + $givenname) -Name $fullname -Enabled $true -EmailAddress $emailaddress -Path $OUpath -AccountPassword (ConvertTo-SecureString $newuserPW -AsPlainText -Force) -ChangePasswordAtLogon $false -OtherAttributes @{Extensionattribute11="1";Extensionattribute3="AGST2";Extensionattribute2=$user.Division;physicalDeliveryOfficeName=$user.Location;"msds-cloudextensionattribute17"=$user.'User E-mail'} -server $targetserver
            if ($user.'User E-mail'.Contains("@"))
            {
                create-mailbox -userid $UserPrincipalName
                create-contact -userid $UserPrincipalName -smtpaddress $user.'User E-mail' -firstname $givenname -lastname $surname
            }
            $time=get-date
            $time.ToString() + ":set-aduser called to disable user " + $UserPrincipalName
            if ($user.enable -ne "Y")
            {
                #Set-ADUser -Identity $uid -Enabled $false -Server $targetserver
            }
        }
    $newusers.Add($uid,$newuserPW)
    $outputcsv = $outputcsv + ('"' + $givenname + '","' + $surname + '","' + $uid + '","' + $newuserPW + '","' + $emailaddress + '"')
    }
}
$outputcsv | sc c:\scripts\t2_newoutput.csv
#Stop-Transcript