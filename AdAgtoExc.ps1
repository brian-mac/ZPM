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
Function ConnectToO365 ()
{
    $usercredential = Get-Credential #-UserName "jkontoni.admin@allegisgroup.com" -Message "Please enter:" 
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
      #$Global:Invsession = Get-PSSession  -InstanceId (Get-OrganizationConfig).RunspaceId.Guid
}
Function ConnectToExch ()
{
    $usercredential = Get-Credential #-UserName "john.kontonis@allegisgroup.com" -Message "Please enter:" 
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

# Main Code

# Init Files
$mode       = [System.IO.FileMode]::Append
$access     = [System.IO.FileAccess]::Write
$sharing    = [IO.FileShare]::Read
$LogPath    = [System.IO.Path]::Combine("C:\temp\AddToDlist.txt")


# create the FileStream and StreamWriter objects
$fs = New-Object IO.FileStream($LogPath, $mode, $access, $sharing)
$Stream = New-Object System.IO.StreamWriter($fs)

ConnectToO365
ConnectToExch

$DependFile ="C:\Temp\AgUsers.csv"
$Depends =import-csv -Path $DependFile

foreach ($mailbox in $Depends)
{
    $MailO365 = ($mailbox.AgsEmail).trim()
    $MailT2 = ($mailbox.email).trim()
    $TargetUser = get-adobject -Filter {mail -eq $MailT2} -Properties *
    AddSigniture $TargetUser $MailO365 
    $MailO365 = $null
}

CloseGracefully
