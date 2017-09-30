Function UnpackDelgates ($Delegates)   
{
    $ValidDelgates = New-Object System.Collections.ArrayList
    $delegates = $delegates.split("|")
    foreach ($Del in $delegates)
    {
        If ($del.Contains("talent2.com") )
        {
            # We need to find the coresponding O365 account
            $TargDel = get-mailbox | Where-Object {$_.ForwardingSmtpAddress -eq $Del} -ErrorAction SilentlyContinue
            If ($TargDel)
            {
                #Found a mailbox with the T2 as a forwarding value
                 $ValidDelgates.add($TargetDel.ForwardingSmtpAddress)
            }
            else
            {
                #Lets try and find a recipient (Ok a contact really) with an email address of Talent2 
                # Ok this firmly makes the assumption we are in T2 and AG not any mail enviroment
                $TargRecp = Get-Recipient | Where-Object {$_.PrimarySmtpAddress -eq $Del} -ErrorAction SilentlyContinue
                if ($TargRecp)
                {
                    # Found Something.
                    If ($TargRecp.RecipientType -eq "MailContact")
                    {
                        #Yep it is a contact, Now we have to find the account this is a forwarder for
                        $ContactID = $TargRecp.id 
                        $TargDel = get-mailbox | Where-Object {$_.ForwardinAddress -eq $ContactID} -ErrorAction SilentlyContinue
                        If ($TargDel)
                        {
                            $ValidDelgates.add($TargDel.PrimarySmtpAddress)
                        }
                        else 
                        {
                            $Line = "Error: Unpack Delgate: Could not find a O365 mailbox for delegate of $Del " 
                            # Write-Line $Line
                        }
                    } 
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
                # Write-Line $Line
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
            $Target = get-mailbox -identity $O365Specific -ErrorAction SilentlyContinue
            if ($Target)
            {
               $ValidatedDeliagtes = UnpackDelgates $Hash[$GoogleUPN]
               $DelgateFlag = $true
            }
            else
            {
                $Line = "Error: Could not find the Mailbox $Target, unexpected this was"
                #WriteLine $Line    
            }
        }
    }
    else
    {
        If ($hash[$GoogleUPN])
        {
            # Mailbox delegation found
            #Check mailbox existis 
            $Target = get-mailbox -identity $GoogleUPN -ErrorAction SilentlyContinue    
            if ($Target)
            {
                $ValidatedDeliagtes = UnpackDelgates $Hash[$GoogleUPN]
                $DelgateFlag = $true
            }
            else
            {
                $Line = "Error: Could not find the Mailbox $Target, unexpected this was"
                #WriteLine $Line    
            }  
        }
    }
    # add permissions , check type of mailbox shared = send as  
    if ($DelgateFlag)
    {
        #Create A session specific for the invoke commands
        $Invsession = Get-PSSession -InstanceId (Get-OrganizationConfig).RunspaceId.Guid
        # Loop through each delegate
        foreach ($IndividualDel in $ValidatedDeliagtes)
       {
            if ($IndividualDel)
            {
                $MailboxID = $Target.id
                $error = $null
                $IndividualDel = "urrgg"
                try
                {
                    Invoke-Command -Session $Invsession -ScriptBlock {add-mailboxpermission -identity $Using:MailboxId  -User $Using:IndividualDel -AccessRight FullAccess} -ErrorAction Stop > $null
                    $Line = "Sucsess: $IndividualDel added to $Target"
                }
                Catch 
                {
                    $Line ="Error: $IndividualDel count NOT be added to $Target"
                }
                #writeline $Line
                if ($Target.IsShared)
                {
                    try
                    {
                        Invoke-Command -Session $Invsession -ScriptBlock {Add-RecipientPermission -identity $Using:MailboxID  -AccessRights SendAs -Trustee $Using:IndividualDel} > $Null
                        $Line = "Sucsess: Sendas added for $IndividualDel to $Target"
                    }
                    Catch 
                    {
                        $Line ="Error: SendAs not added for $IndividualDel to $Target"
                    }
                    #Writeline $Line
                    try
                    {
                        Invoke-Command -Session $Invsession -ScriptBlock {Set-Mailbox $Using:MailboxId  -MessageCopyForSentAsEnabled $True} > $Null
                        $Line = "Sucsess: MessageCopyForSentAsEnabled for $Target"
                    }
                    Catch 
                    {
                        $Line ="Error: MessageCopyForSentAsEnabled not set for $Target"
                    }
                    #Writeline $Line                
                }
            }
        }
    }
$DelgateFlag = $false
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

# ConnectToO365

$Hash=@{}
$DependFile ="C:\Temp\dependancyreport.csv"
$Depends =import-csv -Path $DependFile
foreach ($dependecy in $Depends)
{
    $Hash.Add($dependecy.email, $dependecy.'Inbox Delegated To')

}
AddDeligations "aaron.clancy@talent2.com" "Aaron.Clancy@AllegisGlobalSolutions.com"
