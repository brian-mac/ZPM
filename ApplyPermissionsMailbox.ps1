Function CloseGracefully()
{
    # Close all file streams, files and sessions.
    $Stream.writeline( $Date +  " PSSession and log file closed.")
    Write-Host $Date  +  " PSSession and log file closed."
    $Stream.Close()
    $fs.Close()
    #$StreamAsia.close()
    #$fsAsia.close()
    
    # Close PS Sessions
    # Get-PSSession | Remove-PSSession
    $error.clear()
    Exit
}
function WriteLine ($LineTxt) 
{
    $Date = get-date -Format G
    $Date = $Date + "    : "  
    $LineTxt = $date + $LineTxt  
    $Stream.writeline( $LineTxt )
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
                            # Sooo if they are already migrated the forwarding will have been removed, here comes the hail Mary
                            $TrySamAcc = ($TargRecp.Name).split(".")
                            $TrySamAcc = $TrySamAcc.item(0)
                            $TargDel = get-mailbox -Identity $TrySamAcc -ErrorAction SilentlyContinue
                            if ($TargDel)
                            {
                                 $ValidDelgates.add($del) 
                            } 
                            Else
                            {
                                $Line = "Error: Unpack Delgate: Could not find a O365 mailbox for delegate of $Del or $TrySamAcc" 
                                WriteLine $Line
                            }
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
    
    If ($O365Specific.Length -gt 1)
    {
        # This means there will be a discrepencey between O365 account and GoogleUPN.
        # Check for dependecies using Google UPN.
        If ($hash[$GoogleUPN] )
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
    Else
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


# Main Body
ConnectToO365
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

#$fsAsia = New-Object IO.FileStream($AsiaLog, $ModeAsia, $access, $sharing)
#$StreamAsia = New-Object System.IO.StreamWriter($fsAsia)
#Write headers for AsiaAddToGapp.csv
#$AsiaLog = "EmailAddress,SamAccountName,GoogleGroup"
#WriteAsia $AsiaLog

ConnectToO365

$Hash=@{}
$DependFile ="C:\Temp\dependacyreport.csv"
$Depends =import-csv -Path $DependFile
foreach ($dependecy in $Depends)
{
    $DepEmail = ($dependecy.email).Trim()
    $DepDel = ($dependecy.'Inbox Delegated To'.Trim())
    $Hash.Add($DepEmail, $DepDel)

}
# AddDeligations "aaron.clancy@talent2.com" "Aaron.Clancy@AllegisGlobalSolutions.com"
#This Loop to use this as standalone from a file, it would usually be used in code that passes one user at a time.
foreach ($mailbox in $Depends)
{
    AddDeligations ($mailbox.email).trim() $null
}
#AddDeligations "ss.accounts@allegisglobalsolutions.com" $null
CloseGracefully