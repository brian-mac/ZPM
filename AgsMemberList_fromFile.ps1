Function FindAgExchMailbox ($TargetMailbox)
{
   $mailbox = get-mailbox $TargetMailbox.Trim()
   return $mailbox
}

Function FindAgContact ([String]$targetContact)
{
    $targetContact = $targetContact.trim()    
    $FindAgContact = get-mailcontact -filter "WindowsEmailAddress -eq '$targetContact'"
    return $FindAgContact

}

Function FindMailboxWithContactForward ([String]$TargetForward)
{
    if ($Targetforward.IndexOf("'") -gt 0)
    {
        $TargetForward = $Targetforward.Replace("'","''")
    }
    $FindMailboxWithContactForward = get-mailbox -filter "forwardingaddress -eq '$TargetForward'" 
    return  $FindMailboxWithContactForward
}

$AllOHP = " "
$AllOHP =""
$AgsOutput = New-Object -TypeName PsObject
  Add-Member -InputObject $AGSOutput   -MemberType NoteProperty -Name ObjectClass      -Value "Test"
  Add-Member -InputObject $AGSOutput   -MemberType NoteProperty -Name Samaccountname   -Value "Test"
  Add-Member -InputObject $AGSOutput   -MemberType NoteProperty -Name Displayname      -Value "Test"
  Add-Member -InputObject $AGSOutput   -MemberType NoteProperty -Name givename         -Value "Test"
  Add-Member -InputObject $AGSOutput   -MemberType NoteProperty -Name Sn               -Value "Test"
  Add-Member -InputObject $AGSOutput   -MemberType NoteProperty -Name mail             -Value "Test"
  Add-Member -InputObject $AGSOutput   -MemberType NoteProperty -Name AGSmail          -Value "Test"
  Add-Member -InputObject $AGSOutput   -MemberType NoteProperty -Name otherHomePhone   -Value "Test"
  Add-Member -InputObject $AGSOutput   -MemberType NoteProperty -Name Streetaddress    -Value "Test"
  Add-Member -InputObject $AGSOutput   -MemberType NoteProperty -Name Company          -Value "Test"
  Add-Member -InputObject $AGSOutput   -MemberType NoteProperty -Name department       -Value "Test"
  Add-Member -InputObject $AGSOutput   -MemberType NoteProperty -Name description      -Value "Test"

$OutputFile = "C:\Temp\AgsMemberList.csv"

#$AGS = Get-ADobject -LDAPFilter "NAME=*" -SearchBase "OU=AGS,OU=EXTERNAL,DC=talent2,DC=corp" -SearchScope Subtree -properties *
#$AGS = Get-ADobject -Filter * -SearchBase "OU=AGS,OU=EXTERNAL,DC=talent2,DC=corp" -SearchScope Subtree -properties *
$InputFile = Import-Csv -Path "C:\temp\AGS_Inputfull.csv"
Import-Module activedirectory
$cred = Get-Credential -Message "Please enter user name and password for allegisgroup.com"
$sessionOption = New-PSSessionOption -ProxyAccessType IEConfig 
$ExSession   = New-PSSession -Authentication basic -Credential $cred -ConnectionUri https://outlook.allegisgroup.com/PowerShell/ -ConfigurationName Microsoft.Exchange -AllowRedirection  -SessionOption $sessionOption 
Import-PSSession $ExSession
foreach($Email_rec in $InputFile) 
{
    $TargetEmail = $Email_rec.Mail
    $SamAccNAme = $TargetEmail.split("@").item(0)
    $AgsObject = Get-ADObject -Filter {samaccountname -eq $SamAccNAme} -properties *
    if ( $AgsObject.ObjectClass -eq "contact" -or$AgsObject.ObjectClass -eq "user" )
    {
        foreach($IndHPO in  $AgsObject.otherHomePhone)
        {
            $AllOHP = $AllOHP +  $IndHPO.Substring(5) + ";" 
            if ($IndHPO.Substring(5) -like "*@allegisglobalsolutions.com")
            {
                $AGSmail = $IndHPO.Substring(5) 
            }
        }
        if ($AGSmail -ne $null)
        {
            $AGExch = FindAgExchMailbox $AGSmail
            if ($AGExch -ne $null)
            {
                $AGSmail = $AGExch.PrimarySmtpAddress
            }
            else
            {
                $AGSmail = "Has alias but no mailbox"
            }
        }
        else
        {
            $AGScontact = FindAgContact  $agsobject.mail 
            if ($AGScontact -ne $Null)
            {
                $AGSForward = FindMailboxWithContactForward $AGScontact.DistinguishedName 
                if ($AGSForward -ne $Null)
                {
                    $AGSmail = $AGSForward.PrimarySmtpAddress
                }
                Else
                {
                    $AGSmail = "Contact No MailBox"
                }
            }
            else
            {
                $AGSmail = "No Contact"
            }
        }
    }
        #$AGSOutput.ObjectClass    = $AgsObject.ObjectClass
        #$AGSOutput.Samaccountname = $AgsObject.Samaccountname
        $AGSOutput.Displayname    = $AgsObject.Displayname
        $AGSOutput.givename       = $AgsObject.givenname
        $AGSOutput. Sn            = $AgsObject.sn
        $AGSOutput.mail           = $AgsObject.mail
        $AGSOutput.AGSmail        = $AGSmail
        $AGSOutput.otherHomePhone = $AllOHP
        #$AGSOutput.Streetaddress  = $AgsObject.streetaddress
        #$AGSOutput. Company       = $AgsObject.company
       # $AGSOutput.department     = $AgsObject.department
        #$AGSOutput.description    = $AgsObject.description
        $AgsOutput | export-csv -Path $OutputFile -Append -Force -NoTypeInformation 
        $AllOHP  = $Null
        $AGSmail = $Null
        $AGScontact = $null
        $AGSForward = $null

}
    

