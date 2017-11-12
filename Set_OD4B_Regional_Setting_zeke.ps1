Start-Transcript "C:\Scripts\O365APAC_Update_OD4B_RegionalSettings\Transcript.log"

import-module Microsoft.Online.SharePoint.PowerShell

$Logfile = "C:\Scripts\O365APAC_Update_OD4B_RegionalSettings\logfile.log"
if (Test-Path $Logfile) {Remove-Item $Logfile}
function LogWrite {
Param ([string]$Logstring)
Add-Content $Logfile -value $logstring -ErrorAction Stop
}

$ErrorLogfile = "C:\Scripts\O365APAC_Update_OD4B_RegionalSettings\Errorlogfile.log"
if (Test-Path $ErrorLogfile) {Remove-Item $ErrorLogfile}
function ErrorLogWrite {
Param ([string]$ErrorLogstring)
Add-Content $ErrorLogfile -value $Errorlogstring -ErrorAction Stop
}



function FindTimeZone ($targetOffice)
{
    # Find TimeZone based on Office code, defaults to Singapore as Asia has more staff.
    $TargetOffice = $TargetOffice.substring(0,3)
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







Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

$SecurestringCloud = "000000000000000000000000000000" | ConvertTo-SecureString 
$UserName = "EXO_Admin2@ALLEGISCLOUD.onmicrosoft.com"
$CloudCredentials = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurestringCloud

#Connect to Exchange Online
    if (!(Get-Command New-OnPremisesOrganization -ErrorAction SilentlyContinue))
   	{
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -Authentication Basic -ConnectionUri https://ps.outlook.com/powershell -AllowRedirection:$true -Credential $CloudCredentials 
        Import-PSSession $session
	}





#$GroupToCheck = "Criteria_DLIST-Allegis_APAC-All_Criteria"
$GroupToCheck = "Exchange_Auth_test@allegisgroup.com"
$MembersToCheck = get-distributiongroupmember -resultsize unlimited $GroupToCheck




$Error.Clear()
$SendEmail = "False"

foreach ($member in $MembersToCheck){

$ChangesReq = "False"
$TZChangeReq = "FAlse"
$LocaleIdChangeReq = "False"

#GetLocaleID value from Country

if ($member.CountryOrRegion -eq $null) {

				Write-Host "Country is null. Setting Australia as the LocaleID for $($member.WindowsLiveID)"
				LogWrite ("Country is null. Setting Australia as the LocaleID for $($member.WindowsLiveID)")
				$LocaleID = "3081"

$LocaleID = "3081" 



}
elseif ($CountryHash.Item($member.CountryOrRegion))
				{ 
				LogWrite ("")
				LogWrite ("-------------------------------------------------------")
				LogWrite ("LocaleID match based on Country for $($member.WindowsLiveID): $($CountryHash.Item($member.CountryOrRegion))")
				$LocaleID = $CountryHash.Item($member.CountryOrRegion)
				
				
				}
else {
				LogWrite ("")
				LogWrite ("-------------------------------------------------------")
				Write-Host "Cannont find $($member.CountryOrRegion) in LocaleID list. Will use Australia as the LocaleID for $($member.WindowsLiveID)"
				LogWrite ("Cannont find $($member.CountryOrRegion) in LocaleID list. Will use Australia as the LocaleID for $($member.WindowsLiveID)")
				$LocaleID = "3081"
				
				}


	
	
	try {

	$Error.Clear()
	$Username = "exo_Admin2@allegiscloud.onmicrosoft.com"
	$Password = "01000000d08c9ddf0115d1118c7a00c04fc297eb01000000bb96e119f41db84b93436d9d5fcd99e7000000000200000000001066000000010000200000006524cca08e60f956d286bf5d5ed56d84097a43500920b06f1421f7f329ce6112000000000e80000000020000200000003acb2d27888166a5ef41cf7b208d704a243aa9ba5de408cde1fe8b606f9b90df200000004707fd996b5968cf75da9cfa269246401372fc582e06fef568ba419a62449255400000009306e4a99096d2ad437c658e550fe5df3a228a0f6df3da4a5103fd6708e2fbaa259c199a8ccb419b8d5c581fa88c2d9c141ae0d136386b6b495e64bee0303b3e" | ConvertTo-SecureString 
	$Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username,$Password)
	$OneDriveSiteName = "https://allegiscloud-my.sharepoint.com/personal/" + ($member.WindowsLiveID.replace(".","_")) -replace "@","_"
	$Context = New-Object Microsoft.SharePoint.Client.ClientContext($OneDriveSiteName)
	$Context.Credentials = $Creds
	$RegionalSettings = $Context.Web.RegionalSettings
	$Context.Load($RegionalSettings)
	$Context.ExecuteQuery()
	$CurrentLocaleID = $RegionalSettings.LocaleID
	$Context.load($Context.web.RegionalSettings.timezone)
	$Context.ExecuteQuery()
	$CurrentTimeZone = $Context.Web.RegionalSettings.Timezone.Id
	Write-Host "Current LocaleID for $($member.WindowsLiveID) is $($RegionalSettings.LocaleID)" -ForegroundColor Green
	Write-Host "Current Time Zone for $($member.WindowsLiveID) is $($RegionalSettings.TimeZone.ID)" -ForegroundColor Green
	LogWrite ("Current LocaleID for $($member.WindowsLiveID) is $($RegionalSettings.LocaleID)")
	LogWrite ("Current Time Zone for $($member.WindowsLiveID) is $($RegionalSettings.TimeZone.ID)")

	}
	
catch {

	ErrorLogWrite ("")
   	$ErrorMessage = $_.Exception.Message
	ErrorLogWrite ("-------------------------------------------------------")
	ErrorLogWrite ("Error setting up PS context for:" + $member.WindowsLiveID)
	ErrorLogWrite ("Message: " + $ErrorMessage)
	LogWrite ("Error setting up PS context - See Errorlog") 
	write-host "Error setting PS context for $($member.WindowsLiveID)" -fore red
	write-host "Message: " $ErrorMessage -fore red



	}
	
	#Get Time Zone ID
	
	$TZ_ID = FindTimeZone $member.office
	
	if ($CurrentTimeZone -ne $TZ_ID) {$TZChangeReq = "True";$ChangesReq="True";$SendEmail="True"}
	if ($CurrentLocaleID -ne $LocaleID) {$LocaleIdChangeReq = "True";$ChangesReq="True";$SendEmail="True"}

	if ($ChangesReq -eq "True") {


	try {

	$Error.Clear()
	connect-SPOService -url https://allegiscloud-admin.sharepoint.com -credential $CloudCredentials
	Set-SPOUser -site $OneDriveSiteName -loginname exo_admin2@allegiscloud.onmicrosoft.com -isSiteCollectionAdmin $true -errorAction Stop
	

	}
	
catch {

	ErrorLogWrite ("")
   	$ErrorMessage = $_.Exception.Message
	ErrorLogWrite ("-------------------------------------------------------")
	ErrorLogWrite ("Error setting Site Collection Admin for OneDrive site for:" + $member.WindowsLiveID)
	ErrorLogWrite ("Message: " + $ErrorMessage)
	LogWrite ("Error setting Site Collection Admin to Service Account - See Errorlog") 
	write-host "Error setting Site Collection Admin to Service Account for $($member.WindowsLiveID)" -fore red
	write-host "Message: " $ErrorMessage -fore red
	Write-Host "Could not set Admin perms! Check user has OneDrive: $($member.WindowsLiveID)"
	LogWrite ("Skipping $($member.WindowsLiveID) - See Error Log") 
	continue

	}



}



	
#Update the LocaleID, if necessary

	
if ($LocaleIdChangeReq -eq "True") { 
	
	
	LogWrite ("Setting $($member.WindowsLiveID) LocaleID to $($LocaleID)")
	Write-Host "Setting $($member.WindowsLiveID) LocaleID to $($LocaleID)" -ForegroundColor Green
	
try {
	
	$Error.Clear()
	$Context.Web.RegionalSettings.LocaleId = $LocaleID
	$Context.Web.Update()
	$Context.ExecuteQuery()

	}
	
catch {

	ErrorLogWrite ("")
   	$ErrorMessage = $_.Exception.Message
	ErrorLogWrite ("-------------------------------------------------------")
	ErrorLogWrite ("Error setting LocaleID for:" + $member.WindowsLiveID)
	ErrorLogWrite ("Message: " + $ErrorMessage)
	LogWrite ("Error setting LocaleID - See Errorlog") 
	write-host "Error setting LocaleID for $($member.WindowsLiveID)" -fore red
	write-host "Message: " $ErrorMessage -fore red



	}

}

else {

	
	LogWrite ("$($member.WindowsLiveID) already has correct LocaleID") 
	write-host "$($member.WindowsLiveID) already has correct LocaleID" -fore green
	


	}

	
#Update the timezone, if necessary


if ($TZChangeReq -eq "True") {
	
		
try {

	$Error.Clear()
	$TZs = $Context.Web.RegionalSettings.TimeZones
	$Context.Load($TZs)
	$Context.ExecuteQuery()

	}

catch {

	ErrorLogWrite ("")
   	$ErrorMessage = $_.Exception.Message
	ErrorLogWrite ("-------------------------------------------------------")
	ErrorLogWrite ("Error getting Time Zone Values for $($member.WindowsLiveID)")
	ErrorLogWrite ("Message: " + $ErrorMessage)
	LogWrite ("Error getting Time Zones - See Errorlog") 
	write-host "Error getting Time Zones for $($member.WindowsLiveID)" -fore red
	write-host "Message: " $ErrorMessage -fore red


	}	



LogWrite ("Setting $($member.WindowsLiveID) TimeZone to $($TZ_ID)")
Write-Host "Setting $($member.WindowsLiveID) TimeZone to $($TZ_ID)" -ForegroundColor Green

try {

	$Error.Clear()
	$TZ = $TZs | ? {$_.id -eq $TZ_ID}

	$Context.Web.RegionalSettings.TimeZone = $TZ
	$Context.Web.Update()
	$Context.ExecuteQuery()
	
	}


catch {

	ErrorLogWrite ("")
   	$ErrorMessage = $_.Exception.Message
	ErrorLogWrite ("-------------------------------------------------------")
	ErrorLogWrite ("Error setting Time Zone for:" + $member.WindowsLiveID)
	ErrorLogWrite ("Message: " + $ErrorMessage)
	LogWrite ("Error setting Time Zone - See Errorlog") 
	write-host "Error setting Time Zone for $($member.WindowsLiveID)" -fore red
	write-host "Message: " $ErrorMessage -fore red



	}

}

else {

	LogWrite ("$($member.WindowsLiveID) already has correct Time Zone") 
	write-host "$($member.WindowsLiveID) already has correct Time Zone" -fore green
	


	}

#Cleanup Service Acount Permissions


if  ($ChangesReq -eq "True") {

try {
	
	$Error.Clear()
	Set-SPOUser -site $OneDriveSiteName -loginname exo_admin2@allegiscloud.onmicrosoft.com -isSiteCollectionAdmin $false

}
catch {

	ErrorLogWrite ("")
   	$ErrorMessage = $_.Exception.Message
	ErrorLogWrite ("-------------------------------------------------------")
	ErrorLogWrite ("Error setting removing Site Collection Admin entry :" + $member.WindowsLiveID)
	ErrorLogWrite ("Message: " + $ErrorMessage)
	LogWrite ("Error setting removing Site Collection Admin entry - See Errorlog") 
	write-host "Error setting removing Site Collection Admin entry for $($member.WindowsLiveID)" -fore red
	write-host "Message: " $ErrorMessage -fore red



	}

#Get udpated Time Zone and LocaledId values to verify change was successful



	$RegionalSettings = $Context.Web.RegionalSettings
	$Context.Load($RegionalSettings)
	$Context.ExecuteQuery()
	$NewLocaleID = $RegionalSettings.LocaleID
	$Context.load($Context.web.RegionalSettings.timezone)
	$Context.ExecuteQuery()
	$NewTZ = $Context.Web.RegionalSettings.Timezone.Id
	Write-Host "Post-change LocaleID for $($member.WindowsLiveID) is: $($NewLocaleID)" -ForegroundColor Green
	LogWrite ("Post-change LocaleID for $($member.WindowsLiveID) is: $($NewLocaleID)")	
	Write-Host "Post-change Time Zone for $($member.WindowsLiveID) is: $NewTZ" -ForegroundColor Green
	LogWrite ("Post-change Time Zone for $($member.WindowsLiveID) is: $NewTZ") 

	}

}

stop-transcript

	if ($SendEmail -eq "True") {
	#Send Email
	
	$FromAddress = "Messaging Reports <messaging@allegisgroup.com>"
	#$ToAddress = "zsmith@teksystems.com,cubennet@allegisgroup.com,schinnat@allegisgroup.com,Brian.McElhinney@allegisgroup.com"
	$ToAddress = "zsmith@teksystems.com"
	$MessageSubject = "APAC One Drive Regional Setting Change(s)."
	$MessageBody = "See attached for logs related to LocaleID and Time Zone assignments for APAC OneDrive users"
	$SendingServer = "iSMTP.allegisgroup.com" 
	$SMTPMessage = New-Object System.Net.Mail.MailMessage $FromAddress, $ToAddress, $MessageSubject, $MessageBody
	$Attachment1 = New-Object Net.Mail.Attachment("C:\Scripts\O365APAC_Update_OD4B_RegionalSettings\logfile.log")
	$Attachment2 = New-Object Net.Mail.Attachment("C:\Scripts\O365APAC_Update_OD4B_RegionalSettings\Errorlogfile.log")
	$Attachment3 = New-Object Net.Mail.Attachment("C:\Scripts\O365APAC_Update_OD4B_RegionalSettings\Transcript.log")



	$SMTPMessage.Attachments.Add($Attachment1)
	$SMTPMessage.Attachments.Add($Attachment2)
	$SMTPMessage.Attachments.Add($Attachment3)
	$SMTPClient = New-Object System.Net.Mail.SMTPClient $SendingServer 
	$SMTPClient.Send($SMTPMessage)
	$Attachment1.dispose()
	$Attachment2.dispose()
	$Attachment3.dispose()
 
	}




