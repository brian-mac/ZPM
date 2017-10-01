<# 
.synopsis 
 This script will: 
* Set the LocaleID and TimeZone of a users OneDrive site based on their Country and Office location.

.Description
Version 1.0.0 20171001 RTM
The script has 1 parameter, the path to the log file.  The credentials for SharePoint consisit of a 
$Username variable, amd a path to a encrypted password.  The SecurePassword function can be called with the
"Store" value to prompt for a password and save it to the path specified.
The function ConnectToExchange cn be called with the parameters of a user name and path to an encrypted password
or with no parameters and it will prompt for credentials.

.Parameter LogFile
The path to the log file.

.Outputs
System.io.FileStream Appendes or writed to a file specified by the input parameter.
#>
[CmdletBinding (DefaultParameterSetName="Set 1")]
param (
    [Parameter(Mandatory=$True,HelpMessage="Please enter name for log file",Parametersetname = "Set 1" )][string] $LogFile
)


function SecurePassword ($Target_Path, $Action)
{
    <# Function will either encrypt and save a password to the file specified in the $Target_Path
    Or it will return an encrypted password from the file specified in $Target_Path.
    Returns password 
    #>

    If ($Action -eq "Store")
    {
        $Secure = Read-Host -AsSecureString
        $Encrypted = ConvertFrom-SecureString -SecureString $Secure -Key (244,102,80,104,223,19,65,130,183,11,132,245,74,147,46,142)
        $Encrypted | Set-Content $Path  
    }
    elseif ($Action -eq "Retreive")
    {
        $Secure = Get-Content $Target_Path | ConvertTo-SecureString -Key (244,102,80,104,223,19,65,130,183,11,132,245,74,147,46,142)
    }
    Return $Secure
}
Function CreateFile ($ReqPath, $ReqMode)
{   
    # Create the FileStream and StreamWriter objects, returns Stream Object
    $date = get-date -Format d
    $date = $date.split("/")
    $date = $date.item(2) + $date.item(1) + $date.item(0)
    $FileParts = $ReqPath.split(".")
    $ReqPath = $FileParts.item(0) + $date +"." + $FileParts.Item(1)

    $mode       = [System.IO.FileMode]::$ReqMode
    $access     = [System.IO.FileAccess]::Write
    $sharing    = [IO.FileShare]::Read
    $LogPath    = [System.IO.Path]::Combine($ReqPath)
    $fs = New-Object IO.FileStream($LogPath, $mode, $access, $sharing)
    Return $fs
}
function WriteLine ($LineTxt,$Stream) 
{
    $Date = get-date -Format G
    $Date = $Date + "    : "  
    $LineTxt = $date + $LineTxt  
    $Stream.writeline( $LineTxt )
}
Function CloseGracefully($Stream,$FileSystem)
{
    # Close all file streams, files and sessions. Call this for each FileStream and FileSystem pair.
    $Line = $Date +  " PSSession and log file closed."
    WriteLine $Line
    Write-Host $Date  +  " PSSession and log file closed."
    $Stream.Close()
    $FileSystem.Close() 
    # Close PS Sessions
    Get-PSSession | Remove-PSSession
    $error.clear()
    Exit
}
Function ConnectToExch ($Username, $Target_Path)
{
    <# Connects to a remote Exchange enviroment.  If a UserName is passed it will look in $Target_Path
       for the Encrypted password and use those for credentials. Otherwise it will prompt interactivley
       If a vild session exists it will not create a new session 
    #>

    $PSsessions = Get-PSSession
    foreach ($PsSession in $PSsessions)
    {
        If ($PsSession.computername -eq  "outlook.allegisgroup.com" -and $PsSession.State -ne "Broken")
        {
            $ExSessionExists = $True
        }
    }
    If ( -not ($ExSessionExists))
    {
       if ($Username)
       {
        $SecPass = SecurePassword $Target_Path "Retreive"
        $UserCredential = New-Object System.Management.Automation.PSCredential($Username,$SecPass)
       } 
       Else
       {
        $UserCredential = Get-Credential  -Message "Please enter you credentials for Remote Exchange:" 
       }
       
        $ProxyAddress = [System.Net.WebProxy]::GetDefaultProxy() |select-object address
        if ($ProxyAddress.address)
        {
            $proxyOptions = New-PSSessionOption -ProxyAccessType IEConfig
            $Global:ExSession = new-pssession -configurationname Microsoft.exchange -ConnectionUri https://outlook.allegisgroup.com/powershell/ -Credential $UserCredential -Authentication Basic -AllowRedirection  -SessionOption $proxyOptions
        }
        Else
        {
            $Global:ExSession = new-pssession -configurationname Microsoft.exchange -ConnectionUri https://outlook.allegisgroup.com/powershell/ -Credential $UserCredential -Authentication Basic -AllowRedirection
        } 
        Import-PSSession $ExSession  # -Prefix OnPrem only requires if dual Exch O365 sessions
        If (!$ExSession)
        {
            Exit 
        }
    }
}



function FindTimeZone ($targetOffice)
{
    # Find TimeZone based on Office code, defaults to Singapore as Asia has more staff.
    $TargetOffice = $TargetOffice.substring(0,4)
    Switch -wildcard ($TargetOffice) 
    {
        "AU*"  {$TimeZone = "(UTC+10:00) Canberra, Melbourne, Sydney"
                  $TZCode = 76}
        "AUQ*"  {$TimeZone = "(UTC+10:00) Brisbane"
                  $TZCode = 18}
        "HK*"   {$TimeZone = "(UTC+08:00) Beijing, Chongqing, Hong Kong, Urumqi"
                   $TZCode = 45  }
        "CN*"   {$TimeZone = "(UTC+08:00) Beijing, Chongqing, Hong Kong, Urumqi"
                   $TZCode = 45 }
        "SG*"   {$TimeZone = "(UTC+08:00) Kuala Lumpur, Singapore"
                   $TZCode = 21 }
        "MY*"   {$TimeZone = "(UTC+08:00) Kuala Lumpur, Singapore"
                   $TZCode = 21 }
        "JP*"   {$TimeZone = "(UTC+09:00) Osaka, Sapporo, Tokyo"
                   $TZCode = 20 }
        "NZ*"   {$TimeZone = "UTC+12:00) Auckland, Wellington"
                   $TZCode = 17 }
        Default  {$TimeZone = "(UTC+08:00) Kuala Lumpur, Singapore"
                  $TZCode = 21}
    }
    Return $TZCode
}


# Main Code
<#Connect to Exchange, pass the UserName and Path to the encrypted password.
  If no parrameters passed, you will be prompted for them
#>
ConnectToExch "Brian.mcelhinney@allegisgroup.com" "C:\temp\e.txt"  
<# Create the log file specifed from the input parameter.  
   Append will append to an exisitng file or create a new one if it does not exisit
   Write will create a new file or overwrite an existing one
   The file will have todays date appened to it
#>
$LogFile = CreateFile $LogFile  "Append"
$LogStream = New-Object System.IO.StreamWriter($LogFile)


# Build hash for Country LocaleID pairs.
$CountryHash = @{

"Albania"	=	"1052"	;
"Algeria"	=	"5121"	;
"Bahrain"	=	"15361"	;
"Egypt"	=	"3073"	;
"Iraq"	=	"2049"	;
"Jordan"	=	"11265"	;
"Kuwait"	=	"13313"	;
"Lebanon"	=	"12289"	;
"Libya"	=	"4097"	;
"Morocco"	=	"6145"	;
"Oman"	=	"8193"	;
"Qatar"	=	"16385"	;
"Saudi Arabia"	=	"1025"	;
"Syria"	=	"10241"	;
"Tunisia"	=	"7169"	;
"U.A.E."	=	"14337"	;
"Yemen"	=	"9217"	;
"Armenia"	=	"1067"	;
"Azerbaijan"	=	"2092"	;
"Belarus"	=	"1059"	;
"Bulgaria"	=	"1026"	;
"Hong Kong S.A.R."	=	"3076"	;
"Hong Kong SAR"	=	"3076"	;
"Macau S.A.R."	=	"5124"	;
"People's Republic of China"	=	"2052"	;
"China"	=	"2052"	;
"Singapore"	=	"4100"	;
"Taiwan"	=	"1028"	;
"Croatia"	=	"1050"	;
"Czech Republic"	=	"1029"	;
"Denmark"	=	"1030"	;
"Maldives"	=	"1125"	;
#"Belgium"	=	"2067"	;
"Netherlands"	=	"1043"	;
"Australia"	=	"3081"	;
"Belize"	=	"10249"	;
"Canada"	=	"4105"	;
"Caribbean"	=	"9225"	;
"Ireland"	=	"6153"	;
"Jamaica"	=	"8201"	;
"New Zealand"	=	"5129"	;
"Republic of the Philippines"	=	"13321"	;
"Philippines"	=	"13321"	;
"South Africa"	=	"7177"	;
"Trinidad and Tobago"	=	"11273"	;
"United Kingdom"	=	"2057"	;
"United States"	=	"1033"	;
"Zimbabwe"	=	"12297"	;
"Estonia"	=	"1061"	;
"Faeroe Islands"	=	"1080"	;
"Iran"	=	"1065"	;
"Finland"	=	"1035"	;
"Belgium"	=	"2060"	;
"France"	=	"1036"	;
"Luxembourg"	=	"5132"	;
"Principality of Monaco"	=	"6156"	;
"Switzerland"	=	"4108"	;
"Former Yugoslav Republic of Macedonia"	=	"1071"	;
"Georgia"	=	"1079"	;
"Austria"	=	"3079"	;
"Germany"	=	"1031"	;
"Liechtenstein"	=	"5127"	;
#"Luxembourg"	=	"4103"	;
"Greece"	=	"1032"	;
#"India"	=	"1095"	;
"Israel"	=	"1037"	;
"India"	=	"1081"	;
"Hungary"	=	"1038"	;
"Iceland"	=	"1039"	;
"Indonesia"	=	"1057"	;
"Italy"	=	"1040"	;
"Japan"	=	"1041"	;
#"India"	=	"1099"	;
"Kazakhstan"	=	"1087"	;
#"India"	=	"1111"	;
"Korea"	=	"1042"	;
"Kyrgyzstan"	=	"1088"	;
"Latvia"	=	"1062"	;
"Lithuania"	=	"1063"	;
"Brunei Darussalam"	=	"2110"	;
"Malaysia"	=	"1086"	;
#"India"	=	"1102"	;
"Mongolia"	=	"1104"	;
#"Norway"	=	"1044"	;
"Norway"	=	"2068"	;
"Poland"	=	"1045"	;
"Brazil"	=	"1046"	;
"Portugal"	=	"2070"	;
#"India"	=	"1094"	;
"Romania"	=	"1048"	;
"Russia"	=	"1049"	;
#"India"	=	"1103"	;
"Serbia and Montenegro"	=	"3098"	;
#"Serbia and Montenegro"	=	"2074"	;
"Slovakia"	=	"1051"	;
"Slovenia"	=	"1060"	;
"Argentina"	=	"11274"	;
"Bolivia"	=	"16394"	;
"Chile"	=	"13322"	;
"Colombia"	=	"9226"	;
"Costa Rica"	=	"5130"	;
"Dominican Republic"	=	"7178"	;
"Ecuador"	=	"12298"	;
"El Salvador"	=	"17418"	;
"Guatemala"	=	"4106"	;
"Honduras"	=	"18442"	;
"Mexico"	=	"2058"	;
"Nicaragua"	=	"19466"	;
"Panama"	=	"6154"	;
"Paraguay"	=	"15370"	;
"Peru"	=	"10250"	;
"Puerto Rico"	=	"20490"	;
"Spain"	=	"1034"	;
"Uruguay"	=	"14346"	;
"Venezuela"	=	"8202"	;
"Kenya"	=	"1089"	;
#"Finland"	=	"2077"	;
"Sweden"	=	"1053"	;
#"Syria"	=	"1114"	;
#"India"	=	"1097"	;
"Tatarstan"	=	"1092"	;
#"India"	=	"1098"	;
"Thailand"	=	"1054"	;
"Turkey"	=	"1055"	;
"Ukraine"	=	"1058"	;
"Islamic Republic of Pakistan"	=	"1056"	;
"Uzbekistan"	=	"2115"	;
#"Uzbekistan"	=	"1091"	;
"Viet Nam"	=	"1066"	;
#"United Kingdom"	=	"1106"	;


}


# Load .Net assemblies, use partial name as PC and Servers store the DL's in diffrent locations.
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")

$GroupToCheck = "Criteria_DLIST-Allegis_APAC-All_Criteria" #should we think about having this as a script parameter for EMEA etc...?

Try
{
    <# O365 runs in no language mode and does not report errors when using commandlets 
       bound to the remote session. You need to use Invoke-Command to get an error value. 
    Force of habit using it with Exchange online
    #>

    $MembersToCheck = Invoke-Command -Session $Exsession -ScriptBlock { get-distributiongroupmember -resultsize unlimited $using:GroupToCheck} -ErrorAction Stop    
    $Line = "Sucesess: Found Target Group $GroupToCheck"
    WriteLine $Line $LogStream
}
Catch 
{
    $ErrorLine = "Error: Could not find group $GroupToCheck"
    WriteLine $ErrorLine $LogStream
}

#Authenticate to Site
# Replace with System account etc...
$Username = "Brian.mcelhinney@allegisgroup.com" 
# Path to the file you saved the enctypted password
$Password = SecurePassword "C:\Temp\e.txt" "Retreive" 
$Site = "https://allegiscloud.sharepoint.com"
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($Site)
$Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username,$Password)
$Context.Credentials = $Creds

<#Retrieve the time zones that are available
$TZs = $Context.Web.RegionalSettings.TimeZones
$Context.Load($TZs)
$Context.ExecuteQuery()
$tzs > "C:\temp\Timez.csv"
#>


#Update the LocaleID

# Loop through each member of the DistList
foreach ($member in $MembersToCheck)
{
    if ($member.CountryOrRegion -eq $null) 
    {
        Write-Host "Country is null. Setting Australia as the LocaleID for $($member.PrimarySmtpAddress)"
        $line = "Country is null. Setting Australia as the LocaleID for $($member.PrimarySmtpAddress)"
        WriteLine $Line $LogStream
		$LocaleID = "3081"
    }
    elseif ($CountryHash.Item($member.CountryOrRegion))
	{ 
        Write-Host "Will set" $CountryHash.Item($member.CountryOrRegion) "as the LocaleID for $($member.PrimarySmtpAddress)"
        $line = "Will set $($CountryHash.Item($member.CountryOrRegion)) as the LocaleID for $($member.PrimarySmtpAddress)"
        Writeline $line $LogStream
        $LocaleID = $CountryHash.Item($member.CountryOrRegion)	
    }
    Else 
    {
        Write-Host "Warning: Cannot find $($member.CountryOrRegion) in LocaleID list. Setting Australia as the LocaleID for $($member.PrimarySmtpAddress)"
        $Line =  "Warning: Cannot find $($member.CountryOrRegion) in LocaleID list. Setting Australia as the LocaleID for $($member.PrimarySmtpAddress)"
        WriteLine $Line $LogStream
        $LocaleID = "3081"
    }	
    $Office = $Member.Office 
    # Call FindTimeZone to determine timezone based on the office
    $TimeZone = FindTimeZone $Office
    # Test Code $OneDriveSiteName = "https://allegiscloud-my.sharepoint.com/personal/brian_mcelhinney_allegisgroup_com1"
    try
    {
        # Hmm dont think this will throw an error.
        $OneDriveSiteName = "https://allegiscloud-my.sharepoint.com/personal/" + ($member.PrimarySmtpAddress.replace(".","_")) -replace "@","_"
        $Context2 = New-Object Microsoft.SharePoint.Client.ClientContext($OneDriveSiteName)
        $Line = "Sucsess: found $OneDriveSiteName"
        WriteLine $line $LogStream
    }
    Catch 
    {
        $Line = "Error: Could not find $OneDriveSiteName"
        WriteLine $Line $logStream
    }
    try 
    {
        $Context2.Credentials = $Creds
        $Context2.ExecuteQuery()
    }
    catch 
    {
        $Line = "Error: Credentials or other binding issue"
        WriteLine $Line $LogStream
    }
    try 
    {
        $Context2.Web.RegionalSettings.LocaleId = $LocaleID
        $Line = "Sucsess: LocaleID set to $LocaleID"
        WriteLine $line $LogStream
    }
    catch 
    {
        $Line = "Error: LocaleId could NOT be set to $LocaleID"
        WriteLine $line $LogStream
    }
    try 
    {
        $Context2.Web.RegionalSettings.TimeZone = $Context2.Web.RegionalSettings.TimeZones.GetbyID($TimeZone)
        $Line = "Sucsess: Time Zone has been set to TimeZoneID $TimeZone"
        WriteLine $Line $LogStream 
    }
    catch 
    {
        $line = "Error: TimeZone has NOT been set to TimeZoneID $TimeZone"
        WriteLine $Line $LogStream
    }
    try 
    {
        $Context2.Web.Update()
        $Context2.ExecuteQuery()
    }
    catch 
    {
        $Line = "Error:  Onedrive site not updated, unxepected this was"
        WriteLine $Line $LogStream
    }
    
}
CloseGracefully $LogStream $LogFile
