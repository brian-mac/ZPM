<# 
.synopsis 
 This script will create Distribution Lists in the AllegisGroup.com domain, Under the appropriate OU under allegisgroup.com/Enterprise/Groups/APAC/Distribution List

.Description
The script can either create a single DL at a time or multiple DLs using a CSV file.
The script has a number of input parameters: Display Name, Company, Email Address and a switch External, for an external facing distribution list.

.Parameter  DisplayName
The name that will appear internally in the GAL, the script will preface this with DLIST-Company_

.Parameter Company
This is the full company name, Allegis Group, Aston Carter, Aerotek, Teksystems.  It will be shortened to match AG DL standard. Used to build full display name and email address.

.Parameter EmailAddress
This is what will be seen externaly.  Only enter the email address before the @company.com.

.Parameter External
If this switch is present the DL will be accessable by external parties and will be created in the OU: allegisgroup.com/Enterprise/Groups/APAC/Distribution List/External access
If it is not present the DL will be only accessable internaly and it will be created in the OU: allegisgroup.com/Enterprise/Groups/APAC/Distribution List/Internal Only

.Parameter LegacyContact
If this parameter is set, a contact will be created in the Talent2.corp domain in OU=AllegisGroupDL-Contact,OU=T2 Distribution Groups,DC=talent2,DC=corp, with the same display
name and email address as created in the Allegisgroup.com

.Parameter InputFile
If this parameter is present the script will read Displayname, Company , EmailAddress, LegacyContact; from a suitably headed CSV file.  A template is provided with the script.


.Outputs
System.io.FileStream Appendes to the log file Create-APACDL.log in the %AppsData% directory.

.Example
To call this script in Windows 10 with an input file  : powershell .\Create-APACDistlistv02.ps1 -inputfile 'C:\temp\APAC-DistListBulkCreate.csv'

.Example    
To call this script in windows 10 to create a single DL: powershell .\Create-APACDistlistv02.ps1 -DisplayName 'Brian16 is' -Company 'Allegis Group' -EmailAddress 'Brian16.is' -External
#>

#Version number 3.2.0  20161130 22:00 Added AGS, MLA, Allegis Partners

 [CmdletBinding (DefaultParameterSetName="Set 2")]
 param (
    [Parameter(Parametersetname = "Set 1")][String]$Inputfile , #="C:\Temp\APAC-DistListBulkCreate.CSV" ,
    [Parameter(Mandatory=$True,HelpMessage="Please enter display name",Parametersetname = "Set 2" )][string] $DisplayName ,
    [Parameter(Mandatory=$True,HelpMessage="Please enter the company:Allegis Group Aston Cater Teksystems Aerotek",Parametersetname = "Set 2")][string] $Company ,
    [Parameter(Mandatory=$True,HelpMessage="Please enter the prefix of the external email address everything before the @",Parametersetname = "Set 2")][string] $Emailaddress,
    [Parameter(Mandatory=$False,Parametersetname = "Set 2")][Boolean] $LegacyContact,
    [Parameter(Mandatory=$False,Parametersetname = "Set 2")][Switch] [Boolean]$External 
 )

# Declare Functions to call from main body.

function LogError ($KnownError)
{
    $Date = get-date -Format G
    $Date = $Date + "    : "   
    If ($error.count -eq 0)
    {
        $error.Add("No PS Error Detected")
    }
    If ($KnownError.length -eq 0)
    {
        $KnownError = "Unexpected this was.."
    }
    $Stream.writeline( $Date + $error + $KnownError )
    Write-Host $Date  +  $error + $KnownError 
    $Error.clear()
    $KnownError = "" 
}

Function CloseGracefully()
{
    $Stream.writeline( $Date +  " PSSession and log file closed.")
    Write-Host $Date  +  " PSSession and log file closed."
    $Stream.Close()
    $fs.Close()
    # Close PS Sessions
    Get-PSSession | Remove-PSSession
    $error.clear()
    Exit
}

function CreatelegacyContact ([String]$AGDisplayname,[String]$AGEmail )
{
    
    $Date = get-date -Format G
    $Date = $Date + "    : "   
    $T2OU = "OU=AllegisGroupDL-Contact,OU=T2 Distribution Groups,DC=talent2,DC=corp"
    $GoogleSyncGroup = "CN=GAL-Gapps,OU=Talent2 Asia Security and Distribution Groups,OU=External Self Serve,DC=talent2,DC=corp"
    If ( $ADsession -eq $Null -or $ADSession.State -eq "Closed")
    {
        $ADcred = Get-Credential  -Message "Please enter your Talent2.corp user name and password"
        Try
        {
            $ADSession = New-PSSession -ComputerName T2EDC-DC03 -Credential $ADcred
            Import-PSSession $ADSession  -ErrorAction Continue
            #Import-Module activedirectory
        }
        Catch
        {
            LogError " : Could not create session to Talent2.corp.  Incorrect users details?"
        }
    }
    $CheckDL =  Get-ADObject -Filter {mail -eq $AGEmail} 
    If ($CheckDL -ne $Null)
    {
        LogError ": Contact exisits in Talent2.corp"  
    }
    Else
    {
        $GoogleMail = "smtp:" + $AGEmail
        Try
        {
           $T2Contact = Invoke-Command  -Session $ADSession  -ArgumentList $AGDisplayname,$T2OU,$AGEmail,$GoogleMail -ScriptBlock {
            param($AGDisplayname,$T2OU,$AGEmail,$GoogleMail)
                $T2Contact = New-ADObject -Name $AGDisplayname `
                -DisplayName $AGDisplayname `
                -Path $T2OU `
                -Type "Contact" `
                -OtherAttributes @{'mail'=$AGEmail;'otherHomePhone'=$GoogleMail} `
                -Passthru
            } 
        }
        Catch
        {
            LogError "Could not create Talent2.corp Contact"
        }

        $T2Contact =  Get-ADObject -Server T2EDC-DC03 -Filter {displayname -like $AGDisplayname}
        If ($T2Contact -ne $Null)
        {
            $Stream.writeline( $Date + "Created in Talent2.corp:  " + $AGDisplayname  +"  "+ $AGEmail)
        }
        Try
        {
            Invoke-Command -Session $ADSession -ArgumentList $GoogleSyncGroup,$T2Contact,$Date -ScriptBlock {
             Param ($GoogleSyncGroup,$T2Contact,$Date)
                $Ggroup = [adsi]"LDAP://$GoogleSyncGroup"      #Oh AD how I hate thee, do I realy have to remember this crap.
                $ContactName =  $T2Contact.DistinguishedName
                $Ggroup.Member.Add($ContactName) 
                $Ggroup.psbase.CommitChanges() 
            }
        }
        Catch
        {
            LogError "Could not add Contact to Google Sync Group"
        }
    }
}

function CreateDistList([String]$TargetDisplayName, [String]$TargetCompany,[String]$TargetEmail,[Boolean]$LegacyContact) 
{
    Switch  ($TargetCompany)
    {
        "Allegis Group"            {$ShortCompany = "Allegis"}
        "AllegisGroup"             
        {
            $ShortCompany = "Allegis" 
            $TargetCompany = "Allegis Group"
        }
        "Allegis"
        {
            $TargetCompany = "Allegis Group"
            $ShortCompany = "Allegis"
        }
        "Aston Carter"             {$ShortCompany = "AstonCarter"}
        "AstonCarter"              
        {
            $ShortCompany = "AstonCarter"
            $TargetCompany = "Aston Carter"
        }
        "AeroTek"                  {$ShortCompany = "Aero"}
        "Teksystems"               {$ShortCompany = "Tek"}
        "Allegis Global Solutions" {$ShortCompany = "AGS"}
        "AllegisGlobalSolutions"   {$ShortCompany = "AGS"}
        "AGS" 
        {
            $ShortCompany = "AGS"
            $TargetCompany = "Allegis Global Solutions"
        }
        "mlaglobal"                 {$ShortCompany ="MLA"}
        "Mla Global"
        {
            $ShortCompany = "MLA"
            $TargetCompany = "MLAGlobal"
        }
        "Allegis Partners"      {$ShortCompany = "AllegisPartners"}
        "AllegisPartners"
        {
            $ShortCompany = "AllegisPartners"
            $TargetCompany = "Allegis Partner"
        }
        Default {
                    LogError $Date + $Targetcompany + ":  Is incorrect"
                    $ShortCompany = "Error"
                }
    }
    If ($ShortCompany -ne "error")
    {
        $TargetDisplayName = "DLIST-" + $ShortCompany + "_" + $TargetDisplayName
        $TargetDL = Get-DistributionGroup -filter "name -eq  '$TargetDisplayName'"
        if ($TargetDL -ne $null)
        {  
            $ErrCurr = $TargetDL.name + ":  Display name exists, please make more unique, consider an APAC. prefix"
            logError   $ErrCurr
        }
        Else
        {
            $TargetEmail = $TargetEmail.Split("@").item(0)      #Remove @Company.com in case they added it by error.
            $CheckEmail = $TargetEmail + "@" + $TargetCompany.Replace(" ","") + ".com"
            $TargetDL = Get-Recipient -Filter "Primarysmtpaddress -eq '$CheckEmail'"
            if ($TargetDL -ne $null)
            {
                $ErrCurr = $TargetDL.name + ":  Email name exists, please make more unique, consider an APAC. sufix"
                LogError $ErrCurr
            } 
            else
            {
                If ($External)
                {
                    $Targetou = "OU=External access,OU=Distribution List,OU=APAC,OU=Groups,OU=Enterprise,DC=allegisgroup,DC=com"
                }
                Else
                {
                    $Targetou = "OU=Internal Only,OU=Distribution List,OU=APAC,OU=Groups,OU=Enterprise,DC=allegisgroup,DC=com"
                }
                Try
                {
                    New-DistributionGroup -name $TargetDisplayName  `
                     -Alias  $TargetEmail `
                     –OrganizationalUnit  $TargetOu
                }
                Catch
                {
                    LogError $TargetDisplaName + " Could not create Distribution Group, Unexpected error"
                }
                If ($External)           # This is an extrnal Distribution List
                {
                    Try
                    {
                        Set-DistributionGroup $TargetDisplayName -RequireSenderAuthenticationEnabled:$False `
                        -CustomAttribute14  $TargetCompany
                    }
                    Catch
                    {
                        LogError $TargetDisplayName + " Could not add Company Attribute or -RequireSenderAuthenticationEnabled:$False"
                    }
                }
                Else                     # This is a internal Distribution List
                {
                    Try
                    {
                        Set-DistributionGroup $TargetDisplayName -RequireSenderAuthenticationEnabled:$True `
                        -CustomAttribute14  $TargetCompany 
                    }
                    Catch
                    {
                        LogError $TargetDisplayName + " Could not add Company Attribute or -RequireSenderAuthenticationEnabled:$True"
                    }
                }
                $Stream.writeline( $Date + "Created in Allegisgroup.com:  " + $TargetDisplayName  +"  "+ $CheckEmail)
                If ($LegacyContact)
                {
                    CreatelegacyContact  $TargetDisplayName $CheckEmail
                }
        }
   }
}

}




#Define variables and constants

$ErrorActionPreference = "Stop"
# Get date time for Log file entry
$Date = get-date -Format G
$Date = $Date + "    : "
# Create Log file stream
$mode = [System.IO.FileMode]::Append
$access = [System.IO.FileAccess]::Write
$sharing = [IO.FileShare]::Read
$LogPath = [System.IO.Path]::Combine($Env:AppData,"Create-APACDL.log")
# create the FileStream and StreamWriter objects
$fs = New-Object IO.FileStream($LogPath, $mode, $access, $sharing)
$Stream = New-Object System.IO.StreamWriter($fs)

# Create Remote Exchange Session
$cred = Get-Credential -Message "Please enter user name and password for allegisgroup.com"
$sessionOption = New-PSSessionOption -ProxyAccessType IEConfig 
$ExSession   = New-PSSession -Authentication basic -Credential $cred -ConnectionUri https://outlook.allegisgroup.com/PowerShell/ -ConfigurationName Microsoft.Exchange -AllowRedirection  -SessionOption $sessionOption 
Try 
{
    Import-PSSession $ExSession 
}
Catch
{
    LogError "Remote Exchange Session did not initalise " 
    CloseGraecfully
}

# Main logic
$External = $True       # Debug remove!!
$LegacyContact = $True  # Debug Remove
if($Inputfile.Length -ne 0)  # Was a input file selected, are we processing one DL manualy or many from a file.
{
    $DistLists = import-csv -Path $Inputfile
    foreach ($Distlist in $Distlists)
    {
        if ($Distlist.External -eq "True")
        {
            $External = $True
        }
        Else
        {
            $External = $False
        }
        CreateDistList $Distlist.displayname $Distlist.company $Distlist.emailaddress $Dlist.legacycontact
    }
}
Else
{
    #Validate Parameters
    If ($DisplayName.Length -eq 0 -or $Company.Length -eq $0 -or $Emailaddress.Length -eq 0)
    {
        LogError  "Not all mandatory parameters where entered: Display Name, Company, Email Address"
    }
    Else
    {
        CreateDistList $DisplayName $Company $Emailaddress $LegacyContact
    }
}
CloseGracefully


