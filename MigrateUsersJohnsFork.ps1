<# 
.synopsis 
 This script will: 
*Add the specified user/s to the Google MigratedGapp group.
*Set ExtenstionAttribute1&2 as directed to alter or stop forwarding.  These values synchronise to AG.com then O365.
*Sets the O365 mailbox ForwardingAddress to $Null and the ForwardingSmtpAddress to the value supplied.
*Adds the user to the correct Exclaimer signiture group.
*Check to see if the "user" is actualy a contact in T2.corp and creates a User object with the same settings.
*Check for a AG.com/O365 contact and deltes this contact from the GAL.
*Creates an output of all Asia users C:\temp\AsiaAddToGApp.csv, this  file is used as an input for a script that runs on the Talent2Asia.com domain.

.Description
Version 4.5.3 20170912 (Johns Fork)
The script can either create act on a single user at a time or multiple users using a CSV file.
The script has two input parameters: Target User email address and forwarding email address value.

.Parameter Email
The email address of the user to be migrated.

.Parameter O365ForwardingAddress
This is the email address the mailbox needs to forward to.  It can be set as None to stop forwarding.

.Parameter InputFile
If this parameter is present the script will read Email, O365ForwardingAddress; from a suitably headed CSV file.  A template is provided with the script.

.Outputs
System.io.FileStream Appendes to the log file C:\temp\OMigratedT2Tasks.txt.
System.io.FileStream Creates a new file       C:\temp\AsiaAddToGApp.csv.

.Example
To call this script in Windows 10 with an input file  : powershell .\MigrateUsers.ps1 -inputfile 'C:\temp\MigratedUsersT2.csv'

.Example    
To call this script in windows 10 to migrate a user DL: powershell .\MigrateUsers.ps1 -Email 'Brian.mcelhinney@allegisgroup' -O365ForwardingAddress 'Brian.mcelhinney@csfb.com'
#>
[CmdletBinding (DefaultParameterSetName="Set 2")]
param (
    [Parameter(Parametersetname = "Set 1")][String]$Inputfile , 
    [Parameter(ParameterSetName = "Set 1")][string]$DependFile,
    [Parameter(Mandatory=$True,HelpMessage="Please enter Email Address",Parametersetname = "Set 2" )][string] $Email,
    [Parameter(Mandatory=$True,HelpMessage="Please enter Forwarding Email Address",Parametersetname = "Set 2" )][string] $O365ForwardingAddress,
    [Parameter(Mandatory=$True,HelpMessage="Please enter AGS Email Address",Parametersetname = "Set 2" )][string] $AGSEmail
)
Function CloseGracefully()
{
    # Close all file streams, files and sessions.
    $Stream.writeline( $Date +  " PSSession and log file closed.")
    Write-Host $Date  +  " PSSession and log file closed."
    $Stream.Close()
    $fs.Close()
    $StreamAsia.close()
    $fsAsia.close()
    
    # Close PS Sessions
    Get-PSSession | Remove-PSSession
    $error.clear()
    Exit
}
Function ConnectToO365 ()
{
    $usercredential = Get-Credential -UserName "jkontoni.admin@allegisgroup.com" -Message "Please enter:" 
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
      $Global:Invsession = Get-PSSession -Credential $usercredential -InstanceId (Get-OrganizationConfig).RunspaceId.Guid
}

Function ConnectToExch ()
{
    $usercredential = Get-Credential -UserName "john.kontonis@allegisgroup.com" -Message "Please enter:" 
    $PSsessions = Get-PSSession
    foreach ($PsSession in $PSsessions)
    {
        If ($PsSession.computername -eq  "outlook.allegisgroup.com")
        {
            $ExSessionExists = $True
        }
    }
    If ( -not ($ExSessionExists))
    {
        $ProxyAddress = [System.Net.WebProxy]::GetDefaultProxy() |select-object address
        if ($ProxyAddress.address)
        {
            $proxyOptions = New-PSSessionOption -ProxyAccessType IEConfig
            $ExSession = new-pssession -configurationname Microsoft.exchange -ConnectionUri https://outlook.allegisgroup.com/powershell/ -Credential $UserCredential -Authentication Basic -AllowRedirection  -SessionOption $proxyOptions
        }
        Else
        {
           
        } $ExSession = new-pssession -configurationname Microsoft.exchange -ConnectionUri https://outlook.allegisgroup.com/powershell/ -Credential $UserCredential -Authentication Basic -AllowRedirection
        Import-PSSession $ExSession  -Prefix OnPrem 
        If (!$ExSession)
        {
            CloseGracefully
        }
    }
}
function ConnectToTalent2Asia ()
{
    If ( $ADsession -eq $Null -or $ADSession.State -eq "Closed")
    {
        $ADcred = Get-Credential  -Message "Please enter your Talent2Asia.com user name and password"
        Try
        {
            Set-Item wsman:localhost\client\trustedhosts sg01svr01.talent2asia.com
            $ADSession = New-PSSession -ComputerName sg01svr01.talent2asia.com -Credential $ADcred 

            Import-PSSession $ADSession  -ErrorAction Continue
            #Import-Module activedirectory
        }
        Catch
        {
            WriteLine "Error: Could not create session to Talent2.corp.  Incorrect users details?"
        }
    }
}
function WriteLine ($LineTxt) 
{
    $Date = get-date -Format G
    $Date = $Date + "    : "  
    $LineTxt = $date + $LineTxt  
    $Stream.writeline( $LineTxt )
}

function WriteAsia ($Linetxt)
{
    $StreamAsia.writeline( $LineTxt )
}

Function AddGroup($TargetUser,$TargetGroup)
{
    # Adds the specified user to the specifed group.
    Try
    {
        Add-ADGroupMember -Identity "$TargetGroup" -Members $TargetUser
        $Line = "Sucsess: $TargetUser has been added to $TargetGroup"
        Writeline $Line
    }
    Catch
    {
        $line = "Error: could not add $TargetUser into $TargetGroup"
        Writeline $line
    }
}

Function AddToDlist ($TargetUser, $TargetDlist)
{
    Add-OnpremDistributionGroupMember -Identity $TargetDlist -Member $TargetUser  -ErrorAction SilentlyContinue 
    If ($DlError)
    {
        $line = "Error: could not add $TargetUser into $TargetDlist"
        Writeline $line
    }
    else
    {
        $Line = "Sucsess: $TargetUser has been added to $TargetDlist"
        Writeline $Line
    }  
}

function AddSigniture ($targetUser,$AGSEmail)
{
    # Set Internal Signiture
    $TargetOffice = $targetuser.PhysicalDeliveryOfficeName
    Switch -wildcard ($TargetOffice) 
    {
        "sg01*"  {$SignitureGroup = "Exclaimer-APAC-Internal-SG01"
                  $GroupSuffix = "SG"}
        "sg02*"  {$SignitureGroup = "Exclaimer-APAC-Internal-SG02"
                  $GroupSuffix = "SG"}

        Default  {$SignitureGroup = "Exclaimer-APAC"
                  $GroupSuffix = "Std"}
    }
    IF ($AGSEmail)
    {
        $Targetemail = $AGSEmail
    }
    else
    {
        $Targetemail = $targetuser.mail    
    }   
    AddToDlist $Targetemail $SignitureGroup 
   
    #Set External Signiture
    
    Switch ($targetUser.company)
    {
        "Aston Carter"              {$SignitureGroup = "Exclaimer-APAC-AstonCarter-$GroupSuffix"}
        "AstonCarter"               {$SignitureGroup = "Exclaimer-APAC-AstonCarter-$GroupSuffix"}
        "Teksystems"                {$SignitureGroup = "Exclaimer-APAC-Teksystems-$GroupSuffix" }
        "Aerotek"                   {$SignitureGroup = "Exclaimer-APAC-Aerotek-$GroupSuffix"}
        "Allegis Global Solutions"  {$SignitureGroup = "Exclaimer-APAC-AGS-$GroupSuffix"}
        "AGS"                       {$SignitureGroup = "Exclaimer-APAC-AGS-$GroupSuffix"}
        "Allegis Partners"          {$SignitureGroup = "Exclaimer-APAC-AllegisPartners-$GroupSuffix"}
        "AllegisPartners"           {$SignitureGroup = "Exclaimer-APAC-AllegisPartners-$GroupSuffix"}
        "Allegis Group"             {$SignitureGroup = "Exclaimer-APAC-AG-AllBrands-$GroupSuffix"}
        "AllegisGroup"              {$SignitureGroup = "Exclaimer-APAC-AG-AllBrands-$GroupSuffix"}
        "Talent2"                   {$Error = "Company value is Talent2 for $TargetUser.DisplayName"
                                     WriteLine $Error}
    }
    IF ($AGSEmail)
    {
        $Targetemail = $AGSEmail
    }
    else
    {
        $Targetemail = $targetuser.mail    
    }
    AddToDlist $Targetemail $SignitureGroup 
    $GroupSuffix  = $null
}


function CheckandImportModule ($ModuleName)
{
    $Modules = Get-Module -ListAvailable
    foreach ($Module in $Modules)
    {
        if ($Module.name -eq $ModuleName)
        {
            $ModuleFlag = $true
        }
    }
    If ($ModuleFlag -ne $true)
    {
        Import-Module $ModuleName
    }
}
Function Remove-T2Contact ($TargetContact , $TargetRecipent)
{
    try
    {
        $T2Contact = get-mailcontact $TargetContact
        $T2CEx500 = $T2Contact.legacyExchangeDN
        $T2CEx500 = "X500:$T2CEx500"
    }
    catch
    {
        $Writeline = "Error: Could not find a contact $TContact  Unexpected this was"
        WriteLine $Writeline
    }
    try 
    {
        remove-OnPremMailContact -identity $T2Contact.alias -Confirm:$false
        $Line = "Sucsess: Removed the contact $T2Contact for $TargetUser"
        WriteLine $Line   
    }
    Catch 
    {
        $Line = "Error: Could not remove the contact $T2Contact for $TargetUser"
        WriteLine $Line     
    }
    try 
    {
        Set-OnPremRemoteMailbox -identity $TargetRecipent -emailAddresses  @{Add=$T2CEx500}  
        $Writeline = "Sucsess: $T2CEx500 proxy address added to $TargetRecipent"  
        WriteLine $Writeline
    }
    catch 
    {
        $Writeline = "Error: $T2CEx500 proxy address could not be added to $TargetRecipent"  
        WriteLine $Writeline
    }
}
Function ChangeForwarding ($TargetUser, $ForwardingAddress, $AGSEmail)
{
    # Changes the forwarding address in O365.
    # Either to $Null or to the specifed value.  In all cases contact forwarding is removed.
    IF ($AGSEmail)
    {
        $TargetIdentity = $AGSEmail
    }
    else 
    {
      $TargetIdentity = $TargetUser.EmailAddress  
    }
    $TestUser = get-mailbox -identity $TargetIdentity -ErrorAction 'SilentlyContinue' 
    If ($TestUser)
    {
        # Check to see if we are forwarding to a T2 contact
        if($TestUser.ForwardingAddress)
        {
            if (($TestUser.ForwardingAddress).contains(".T2"))
            {
                $T2Conact = $TestUser.ForwardingAddress
            }
        }
        Set-Mailbox -identity $TargetIdentity -ForwardingAddress  $Null
        if($ForwardingAddress -ne "" -and $ForwardingAddress -ne "None")
        {
            Set-Mailbox -identity $TargetIdentity -ForwardingSmtpAddress  $ForwardingAddress
            $Line = "Sucsess: $TargetIdentity  O365 forwarding has been set to $ForwardingAddress."
            WriteLine $Line
        }
        Else
        {
            Set-Mailbox -identity $TargetIdentity -ForwardingSmtpAddress  $Null
        }
        $Line = "Sucsess: $TargetIdentity  O365 forwarding has been set to Null."
        WriteLine $Line
        If ($T2Conact)
        {
            Remove-T2Contact $T2Conact $TargetIdentity
        }
    }
    Else 
    {
        $Line = "Error: $TargetIdentity  Error O365  forwarding has Not been set to $Forwardingaddress."
        WriteLine $Line
    }
    $T2Conact = $Null
    $TargetIdentity = $Null
    $TestUser = $Null   
}

Function ConvertUser($TargetUser)
{
    # Converts a contact to a user.
    # Only adds User properties that have a value in the original contact.
    $TempSAM = ($TargetUser.mail).split("@")
    $NewSAM = $TempSAM.item(0)
    # Old AGS users may have a company of Talent2.
    $Company = $TargetUser.company 
    if ($Company.tolower().contains("talent".tolower()))
    {
        if ($TargetUSer.DistinguishedName.ToLower().contains("AGS".ToLower()))
        {
            $Company = "Allegis Global Solutions"
        }
    }
    # Determine the path to create the User in.
    $TempPath = ($TargetUser.DistinguishedName).split(",")
    $CountPathParts = ($TempPath.Count) -1
    $NewPath = $Null 
    for ($I=1;$i -le $CountPathParts; $I++)
    {
        $NewPath = $Newpath + $TempPath.item($I) +","
    }
    $NewPath = $NewPath.TrimEnd(",")
    # Can not create a user with the same name as the contact.
    $UserName =  $TargetUser.name + "_User"
    $UPN = "$NewSAM@Talent2.corp"
    $Tmail = $TargetUser.mail
    if ($targetUser.ipphone)
    {
            $IPphone = $TargetUser.ipphone
    }
    else
    {
    $IPphone = "NA"    
    }
    if ($targetUser.extensionAttribute6)
    {
        $Ext6 = $targetUser.extensionAttribute6
    }
    else 
    {   
        $Ext6 = "NA"
    }
    if ($TargetUser.extensionAttribute11)
    {
        $Ext11 = $TargetUser.extensionAttribute11
    }
    else
    {
        $Ext11 = "NA"    
    }
    # Define the properties that we will set for the user.
    $Properties = @("-C",  $TargetUser.C), `
    @("-Country" , $TargetUser.C), `
    @("-Company" ,  $Company ), `
    @( "-DisplayName",  $TargetUser.DisplayName), `
    @( "-givenname" ,  $TargetUser.givenname), `
    @("-manager" , $TargetUser.manager), `
    @("-MobilePhone",  $TargetUser.mobile), `
    @( "-office" , $Office), `
    @("-officephone" ,  $TargetUser.telephonenumber), `
    @("-postalcode" ,  $TargetUser.postalcode), `
    @("-state" ,  $TargetUser.st), `
    @("-Streetaddress" ,  $TargetUser.streetaddress), `
    @( "-surname",  $TargetUser.sn), `
    @("-Title" ,  $TargetUser.Title ), `
    @("-UserPrincipalName", $UPN)
    Try
    {
       #Create the new user with basic properties.
        New-aduser -SamAccountName $NewSAM `
        -Path  $Newpath `
        -EmailAddress  $Tmail `
        -otherattributes @{'extensionAttribute1' = $UserEmail;'extensionattribute2' = "False"; `
        'extensionattribute3' = "4"; 'ipPhone' = $IPphone;'extensionAttribute6' = $Ext6; 'extensionAttribute11' = $Ext11} `
        -name $UserName
        $Line = "Sucsess:  created user $TargetUser with an SAM of $NewSam"
        WriteLine $Line
    }
    Catch
    {
        $Line = "Error: Could not create user $TargetUser with an SAM of $NewSam"
       WriteLine $Line
    } 
    try 
    {
        # Add any non null properties from the contact to the user. 
        foreach ($Property in $Properties)
        {
            if ($property.item(1))
            {
                $command = $Property.item(0)
                $Value = $Property.item(1)
                $INVCommand = "set-aduser -identity $NewSAM $command '$Value' "
                Invoke-Expression $INVCommand
            }
        }
        $Line = "Sucsess: Set defined properties of $NewSam"
        WriteLine $Line
   }
   catch
   {
        $Line = "Error: Could not set propertie $Command with a value of $Value for user $TargetUser with an SAM of $NewSam"
        WriteLine $Line
   }
   try 
   {
        # Add the group membership from the contact to the user. 
        $Tuser = get-aduser -identity $NewSAM -Properties *
        foreach ($GroupMemebr in $TargetUser.memberof)
        {
            AddGroup $Tuser $GroupMemebr
        }
        $Line = "Sucsess: Set all Group Memebrship of user $TargetUser with an SAM of $NewSam"
        WriteLine $Line    
    }
    catch
    {
        $Line = "Error: Could not set all group membership of user $TargetUser with an SAM of $NewSam"
        WriteLine $Line
    }
    Try
    {
        Set-ADAccountPassword -identity $NewSAM -reset -newpassword (ConvertTo-SecureString -AsPlainText "L3tm31n!" -Force)
        Set-ADUser -Identity $NewSAM -Enabled $True
        $Line = "Sucsess: Password Set for $NewSam and account enabled"
        WriteLine $Line
    }
    Catch
    {
        $Line = "Error: Could not set password for $NewSam and account not Enabled"
        WriteLine $Line
    }

    $ConvertedUser = Get-aduser $newSam -Properties *
    $Properties = $Null
    $UserName = $Null
    $Tuser = $Null
    $Tempsam = $Null
    $Newsam = $Null
    $CountPathParts = $Null
    $Temppath = $Null
    $NewPath = $Null
    $TargetUser = $Null
    $IPphone = $Null
    $Tmail = $Null
    Return $ConvertedUser
}
Function UnpackDelgates ($Delegates)   
{
    $ValidDelgates = New-Object System.Collections.ArrayList
    $delegates = $delegates.split("|")
    foreach ($Del in $delegates)
    {
        If ($del.Contains("talent2.com") )
        {
            # We need to find the coresponding O365 account
            $TargDel = get-mailbox -Filter "$_.Forwardingsmtpaddress -eq '$del'" -ErrorAction SilentlyContinue
            If ($TargDel)
            {
                #Found a mailbox with the T2 as a forwarding value
                $ValidDelgates.add($TargDel.PrimarySmtpAddress)
            }
            else
            {
                #Lets try and find a recipient (Ok a contact really) with an email address of Talent2 
                # Ok this firmly makes the assumption we are in T2 and AG not any mail enviroment
                $TargRecp = Get-Recipient -Filter "$_.PrimarySmtpAddress -eq '$Del'" -ErrorAction SilentlyContinue
                if ($TargRecp)
                {
                    # Found Something.
                    If ($TargRecp.RecipientType -eq "MailContact")
                    {
                        #Yep it is a contact, Now we have to find the account this is a forwarder for
                        $ContactID = $TargRecp.DistinguishedName
                        $TargDel = get-mailbox -Filter "$_.ForwardingAddress -eq '$ContactId'" -ErrorAction SilentlyContinue
                        If ($TargDel)
                        {
                           $ValidDelgates.add($TargDel.PrimarySmtpAddress)
                        }
                        else 
                        {
                            $Line = "Error: Unpack Delgate: Could not find a O365 mailbox for delegate of $Del " 
                            WriteLine $Line
                        }
                    } 
                }
                else
                {
                    $Line = "Error: Unpack Delgate: Could not find a O365 mailbox for delegate of $Del " 
                    WriteLine $Line    
                }
            }
        } 
        Else
        {
            # Check delgate exists with this (Branded) smtp address.
            $TargDel = get-mailbox -Identity $del -ErrorAction SilentlyContinue
            if ($TargDel)
            {
                 $ValidDelgates.add($del) 
            } 
            else
            {
                $Line = "Error: Unpack Delgate: Could not find a O365 mailbox for delegate of $Del " 
                WriteLine $Line
            }
        }
    }
    Return $ValidDelgates 
    $ValidDelgates =$null
}

Function AddDeligations ($GoogleUPN,$O365Specific)
{
    #Check to see if current mailbox has a dependcey.
    
    If ($O365Specific)
    {
        # This means there will be a discrepencey between O365 account and GoogleUPN.
        # Check for dependecies using Google UPN.
        If ($hash[$GoogleUPN])
        {
            # Mailbox delegation found, however we need to use O365 specific value to bind to O365 mailbox
            #Check mailbox existis 
            $Line = "Checking:  Google delgation found for $GoogleUPN"
            WriteLine $Line
            $Target = get-mailbox -identity $O365Specific -ErrorAction SilentlyContinue
            if ($Target)
            {
               $ValidatedDeliagtes = UnpackDelgates $Hash[$GoogleUPN]
               $DelgateFlag = $true
            }
            else
            {
                $Line = "Error: Could not find the Mailbox $Target, unexpected this was"
                WriteLine $Line    
            }
        }
    }
    else
    {
        If ($hash[$GoogleUPN])
        {
            # Mailbox delegation found
            #Check mailbox existis 
            $Line = "Checking:  Google Delgation Found for $GoogleUPN"
            WriteLine $Line
            $Target = get-mailbox -identity $GoogleUPN -ErrorAction SilentlyContinue    
            if ($Target)
            {
                $ValidatedDeliagtes = UnpackDelgates $Hash[$GoogleUPN]
                $DelgateFlag = $true
            }
            else
            {
                $TempName = $Target.PrimarySmtpAddress
                $Line = "Error: Could not find the Mailbox $TempName, unexpected this was"
                WriteLine $Line    
            }  
        }
    }
    # add permissions , check type of mailbox shared = send as  
    if ($DelgateFlag)
    {
        #Create A session specific for the invoke commands
        $Invsession = Get-PSSession -InstanceId (Get-OrganizationConfig).RunspaceId.Guid
        $TargName = $Target.name
        $Line = "Checking:  Valid delgates found for $TargName "
        WriteLine
        # Loop through each delegate
        foreach ($IndividualDel in $ValidatedDeliagtes)
        {
            $IndividualDel = ($IndividualDel).tostring()
            if ($IndividualDel.contains("@"))
            {
                $MailboxID = $Target.id
                try
                {
                    Invoke-Command -Session $Invsession -ScriptBlock {add-mailboxpermission -identity $Using:MailboxId  -User $Using:IndividualDel -AccessRight FullAccess} > $null
                    $Line = "Sucsess: $IndividualDel added to $Target"
                }
                Catch 
                {
                    $Line ="Error: $IndividualDel count NOT be added to $Target"
                }
                writeline $Line
                if ($Target.IsShared)
                {
                    try
                    {
                        Invoke-Command -Session $Invsession -ScriptBlock {Add-RecipientPermission -identity $Using:MailboxID  -AccessRights SendAs -Trustee $Using:IndividualDel -Confirm:$false} > $Null
                        $Line = "Sucsess: Sendas added for $IndividualDel to $Target"
                    }
                    Catch 
                    {
                        $Line ="Error: SendAs not added for $IndividualDel to $Target"
                    }
                    Writeline $Line
                    try
                    {
                        Invoke-Command -Session $Invsession -ScriptBlock {Set-Mailbox $Using:MailboxId  -MessageCopyForSentAsEnabled $True} > $Null
                        $Line = "Sucsess: MessageCopyForSentAsEnabled for $Target"
                    }
                    Catch 
                    {
                        $Line ="Error: MessageCopyForSentAsEnabled not set for $Target"
                    }
                    Writeline $Line                
                }
            }
        }
    }
 $DelgateFlag = $false
}
Function ProcessUser($MigratedUser, $ForwardingAddress, $AGSEmail)
{
    # Check and Modify Talent2.corp objects and properties
    # Check there are not multiple AD Objects with this smtp address.
    Try 
    {
        $UserEmail = $MigratedUser
        $TargetUser = get-adobject -Filter {mail -eq $UserEmail} -Properties *
        # If I have one item returned it will be an ADentity otherwise it will be an array/collection
        $Cast = $TargetUser.gettype()
        if ($Cast.basetype.name -eq "ADentity")
        {
            if ($targetuser.objectclass -eq "Contact")
            {
                $Error = "Warning: $targetuser is a Contact, attempting to convert to user"
                Writeline $Error
                $TargetUser = ConvertUser $targetUser
            }
            else
            {
                $TargetUser = get-aduser  -Filter {Emailaddress -eq $UserEmail} -Properties *
            }
        }
        else
        {
            $Error = "Error: More than one object with a smtp address of $UserEmail"
            WriteLine $Error
            Return 
        }
    }       
    Catch 
    {
        $Error = "Error: $MigratedUser does not exist, please check Email Address"
        WriteLine $Error
        Return 
    }
    AddGroup $TargetUser $Gaap
    # Check to see if user existis in an OU that equates to a talent2asia.com domain.
    $Disname = $TargetUser.DistinguishedName
    $Disname = $Disname.tolower()
    if ($Disname.contains("external") -and (!$Disname.contains("ausnz")))
    {
        $AsiaUser = $TargetUser.SamAccountName    
        $AsiaLog = "$MigratedUser,$AsiaUser,$Gaap"
        WriteAsia $AsiaLog
    }
    If ($ForwardingAddress)
    {
        try 
        {
            Set-ADUser $TargetUser -replace @{'extensionAttribute1'=$ForwardingAddress}
            $Line = "Sucsess: $TargetUser  extensionAttribute1 flag has been set $ForwardingAddress"
            WriteLine $Line
            Set-ADUser $TargetUser -replace @{'extensionAttribute2'=$True}
            $Line = "Sucsess: $TargetUser  extensionAttribute2 flag has been set $True"
            WriteLine $Line
        }
        catch 
        {
            $Line = "Error: $TargeUser extensionAttribute2 could not be set to True"
            WriteLine $Line
        }    
    }
    else
     {
        try 
        {
            Set-ADUser $TargetUser -replace @{'extensionAttribute2'=$False}
            $Line = "Sucsess: $TargetUser  extensionAttribute2 flag has been set to False"
            WriteLine $Line
        }   
        catch 
        {
            $Line = "Error: $TargeUser extensionAttribute2 could not be set to False"
            WriteLine $Line
        }
     }       
    # Call functions for O365 and Exchange Online Objects and properties.
    ChangeForwarding $TargetUser $ForwardingAddress $AGSEmail
    AddSigniture $TargetUser $AGSEmail
    AddDeligations $TargetUser $AGSEmail

    $AGSEmail = $null
    $ForwardingAddress = $null
}

# Main Body
# Create Log file stream
$date = get-date -Format d
$date = $date.split("/")
$date = $date.item(2) + $date.item(1) + $date.item(0)
$Temp_Asia_log = "C:\temp\AsiaAddToGApp$date.csv"
$mode       = [System.IO.FileMode]::Append
$ModeAsia   = [System.IO.FileMode]::Create
$access     = [System.IO.FileAccess]::Write
$sharing    = [IO.FileShare]::Read
$LogPath    = [System.IO.Path]::Combine("C:\temp\OMigratedT2Tasks.txt")
$AsiaLog    = [System.IO.Path]::Combine($Temp_Asia_log)

# create the FileStream and StreamWriter objects
$fs = New-Object IO.FileStream($LogPath, $mode, $access, $sharing)
$Stream = New-Object System.IO.StreamWriter($fs)

$fsAsia = New-Object IO.FileStream($AsiaLog, $ModeAsia, $access, $sharing)
$StreamAsia = New-Object System.IO.StreamWriter($fsAsia)
#Write headers for AsiaAddToGapp.csv
$AsiaLog = "EmailAddress,SamAccountName,GoogleGroup"
WriteAsia $AsiaLog

#Set Target Group
$Gaap = "APAC-Migrated-Gapps"

ConnectToExch
ConnectToO365
CheckandImportModule "ActiveDirectory" 
#Test the inputfile
#$Inputfile = "C:\temp\test.csv"
#$DependFile ="C:\temp\dependacyreport.csv"
# Load dependcey file as a has table for quick searching.
$Hash=@{}
$DependFile ="C:\Temp\dependacyreport.csv"
$Depends =import-csv -Path $DependFile
foreach ($dependecy in $Depends)
{
    $Hash.Add($dependecy.email, $dependecy.'Inbox Delegated To')
}

if ($inputfile.Length -ne 0)
{
    $SMigratedUsersT2 = import-csv -Path $Inputfile
    foreach ($MigratedUser in $SMigratedUsersT2)
    {
        $ForwardingAddress = ($MigratedUser.ForwardingAddress).Trim()
        $O365User = ($MigratedUser.email).Trim()
        $AGSEmail = ($MigratedUser.AGSEmail).trim()
        ProcessUser $O365User $ForwardingAddress $AGSEmail
        Writeline ""
        Writeline ""
    }
}
else
{
    ProcessUser $email $O365ForwardingAddress $AGSEmail
}

#Close and write stream to file
closegracefully 
