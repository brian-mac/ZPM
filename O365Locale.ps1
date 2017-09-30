[CmdletBinding (DefaultParameterSetName="Set 1")]
param (
    [Parameter(Mandatory=$True,HelpMessage="Please enter name for log file",Parametersetname = "Set 1" )][string] $LogFile
)


Function CreateFile ($ReqPath, $ReqMode)
{
   
    
    # create the FileStream and StreamWriter objects
    $date = get-date -Format d
    $date = $date.split("/")
    $date = $date.item(2) + $date.item(1) + $date.item(0)
    $ReqPath = "$ReqPath$date.csv"

    $mode       = [System.IO.FileMode]::$ReqMode
    $access     = [System.IO.FileAccess]::Write
    $sharing    = [IO.FileShare]::Read
    $LogPath    = [System.IO.Path]::Combine($ReqPath)
    $fs = New-Object IO.FileStream($LogPath, $mode, $access, $sharing)
    Return $fs
}
Function CloseGracefully($Stream,$FileSystem)
{
    # Close all file streams, files and sessions.
    $Stream.writeline( $Date +  " PSSession and log file closed.")
    Write-Host $Date  +  " PSSession and log file closed."
    $Stream.Close()
    $FileSystem.Close() 
    # Close PS Sessions
    Get-PSSession | Remove-PSSession
    $error.clear()
    Exit
}
Function ConnectToExch ()
{
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
        $usercredential = Get-Credential  -Message "Please enter you credentials for Remote Exchange:" 
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
        Import-PSSession $ExSession  -Prefix OnPrem 
        If (!$ExSession)
        {
            CloseGracefully #Hmm maybe call this function before opening the files, stop instead?
        }
    }
}

function WriteLine ($LineTxt,$Stream) 
{
    $Date = get-date -Format G
    $Date = $Date + "    : "  
    $LineTxt = $date + $LineTxt  
    $Stream.writeline( $LineTxt )
}

function FindTimeZone ($targeOffice)
{
    # Find TimeZone Based On City
    $TargetOffice = $TargetOffice.left(4)
    Switch -wildcard ($TargetOffice) 
    {
        "AU*"  {$TimeZone = "A.U.S. Eastern Standard Time"
                  $TZCode = 255}
        "AUQ*"  {$TimeZone = "E. Australia Standard Time"
                  $TZCode = 260}
        "HK*"   {$TimeZone = "China Standard Time"
                   $TZCode = 210  }
        "CN*"   {$TimeZone = "China Standard Time"
                   $TZCode = 210 }
        "SG*"   {$TimeZone = "Singapore Standard Time"
                   $TZCode = 215 }
        "MY*"   {$TimeZone = "Singapore Standard Time"
                   $TZCode = 215 }
        "JP*"   {$TimeZone = "Tokyo Standard Time"
                   $TZCode = 235 }
        "NZ*"   {$TimeZone = "New Zealand Standard Time"
                   $TZCode = 290 }
        Default  {$TimeZone = "Singapore Standard Time"
                  $TZCode = 215}
    }
    Return $TimeZone
}


# Main Code

$LogFile = CreateFile $LogFile  "Append"
$LogStream = New-Object System.IO.StreamWriter($LogFile)
ConnectToExch

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



[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")



Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"




$GroupToCheck = "Criteria_DLIST-Allegis_APAC-All_Criteria" #should we think about having this as a script parameter for EMEA etc...?

Try
{
    $MembersToCheck = Invoke-Command -Session $Exsession -ScriptBlock { get-distributiongroupmember -resultsize unlimited $using:GroupToCheck}
    $Line = "Sucesess: Found Target Group $GroupToCheck"
    WriteLine $Line $LogStream
}
Catch 
{
    $ErrorLine = "Error: Could not find group $GroupToCheck"
    WriteLine $ErrorLine    $LogStream
}



#Authenticate to Site
$Username = "change to admin username"
$Password = "Enter Password here" | ConvertTo-SecureString -Force -AsPlainText
$Site = "https://allegiscloud.sharepoint.com"
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($Site)
$Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username,$Password)
$Context.Credentials = $Creds

<#
#Retrieve the time zones that are available
$TZs = $Context.Web.RegionalSettings.TimeZones
$Context.Load($TZs)
$Context.ExecuteQuery()
#>


#Update the LocaleID


foreach ($member in $MembersToCheck)
{
    if ($member.CountryOrRegion -eq $null) 
    {
		Write-Host "Country is null. Setting Australia as the LocaleID for $($member.WindowsLiveID)"
        $line = "Country is null. Setting Australia as the LocaleID for $($member.WindowsLiveID)"
        WriteLine $Line $LogStream
		$LocaleID = "3081"
    }
    elseif ($CountryHash.Item($member.CountryOrRegion))
	{ 
		Write-Host "Will set" $CountryHash.Item($member.CountryOrRegion) "as the LocaleID for $($member.WindowsLiveID)"
        $line = "Will set $($CountryHash.Item($member.CountryOrRegion)) as the LocaleID for $($member.WindowsLiveID.toString())"
        Writeline $line $LogStream
        $LocaleID = $CountryHash.Item($member.CountryOrRegion)	
    }
    Else 
    {
        Write-Host "Cannont find $($member.CountryOrRegion) in LocaleID list. Setting Australia as the LocaleID for $($member.WindowsLiveID)"
		$Line =  "Cannont find $($member.CountryOrRegion) in LocaleID list. Setting Australia as the LocaleID for $($member.WindowsLiveID)"
        WriteLine $Line, $LogStream
        $LocaleID = "3081"
    }	
    $Office = $Member.Office 
    $TimeZone = FindTimeZone $Office
    $OneDriveSiteName = "https://allegiscloud-my.sharepoint.com/personal/" + ($member.WindowsLiveID.replace(".","_")) -replace "@","_"
    $Context2 = New-Object Microsoft.SharePoint.Client.ClientContext($OneDriveSiteName)
    $Context2.Credentials = $Creds
    $Context2.ExecuteQuery()
    $Context2.Web.RegionalSettings.LocaleId = $LocaleID
    $Context2.Web.RegionalSettings.TimeZone = $TimeZone
    $Context2.Web.Update()
    $Context2.ExecuteQuery()
}
CloseGracefully ($LogStream,$LogFile)
