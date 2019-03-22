#variables

$caperr=$null
$importfile="C:\temp\Chatter_users.csv" #change this back to $importfile="D:\Download\chatter_users.csv
function ValidFileDate ($TargetFile,$DaysOld,$Sender,$Recipient)
{
    # Checks to see if the target file is older than a certian number of days, if so it will send a mail notifying a target recipient.
    $FileDate = (get-itemproperty -path ($TargetFile)).CreationTime
    $TodaysDate = get-date 
    
    if ($FileDate -lt $TodaysDate.AddDays(-$daysOld))
    {
        SendMail $Sender $Recipient "File is older than $($daysOld) days" "File was last updated on $($FileDate)"
    }

}
Function SendMail ($Sender, $target, $subject, $Body)
{
    $UnpackTarget =  (($Target.split("@").item(0)).replace("."," ")) + " <$Target>"
    send-mailmessage -from $Sender -to $UnpackTarget -subject $subject -Body $Body
}

Function stringConstruct  
{
Param([string]$chatterRole, [string]$location, [string]$EBA, [string]$ChatGroup)



if ($EBA.ToLower() -eq "true" )
    {
        Switch ($location.ToUpper())
        {
            "SYD" {$EBAG="STAREBA2017"}
            "BRIS" {$EBAG="BRISEBA2017"}
            "GC" {$EBAG="GCEBA2017"}

        }
    }

else 
        {
        $EBAG=""
        }    
if ([string]::IsNullOrEmpty($EBAG))

{
$strReturn += "$($chatterRole.ToUpper())|$($location.ToUpper())"
}
else
{
$strReturn += "$($chatterRole.ToUpper())|$($location.ToUpper())|$($EBAG.ToUpper())"
}



if ([string]::IsNullOrEmpty($ChatGroup))
    {
        $ChatGroup=$null
    }

else
{ 
       
       Switch ($ChatGroup.ToUpper())
        {
            "GAMING" {$strChatGroup += "|$($location.ToUpper())GAMING"}
            "FOOD & BEVERAGE" {$strChatGroup += "|$($location.ToUpper())FOODBEVERAGE"}
            "PROPERTY OPERATIONS" {$strChatGroup += "|$($location.ToUpper())PROPERTY"}
            "HOTEL" {$strChatGroup += "|$($location.ToUpper())HOTEL"}
            "SECURITY & SURVEILLANCE" {$strChatGroup += "|$($location.ToUpper())SECURITY"}

        }
}


if ([string]::IsNullOrEmpty($ChatGroup))

{
$strReturn += ""
}
else
{
$strReturn += $strChatGroup
}



return $strReturn
}

$chattermoderatorgroup = "GG-U-Chatter-Moderators"

# Get Moderator Group Mambers
$ChatModerators = Get-ADGroupMember -identity $chattermoderatorgroup -Recursive | Select -ExpandProperty distinguishedName


#Check the age of the source file.  If it is older than xx days send a mail to xxx@star.com.au
ValidFileDate "D:\Download\chatter_users.csv" 7 "Chatter Server <SYDW@star.com.au>" "Craig.alchin@star.com.au"

#Update the below to point to the Oracle user extract.
$OUsers = Import-Csv $importfile

$n=0
Foreach ($Ousr in $OUsers)

{  

$chkEmpID = ($Ousr.EMPLOYEE_NUMBER).trim() # added trim function incase there are trailing spaces
$chkLocation = ($Ousr.SITE).trim()
$chkEBA = ($Ousr.EA).trim()
$chkChatGroup = ($Ousr."CHAT_GROUP ").trim() #The input file has a trailing space in the field name.

                        $ADUsr = Get-ADUser -Filter {employeeNumber -eq $chkEmpID} -Properties extensionAttribute6
                        # Iterate the users and format the Chatter configuration string
                        foreach($ADUser in $ADUsr)
                        {


                        If ($ChatModerators -contains $ADUser.distinguishedName) {
                         $chatterRole = "Moderator"
                         } Else {
                         $chatterRole = "Free"
                        }

                        $AextAttrib = stringConstruct -chatterRole $chatterRole -location $chkLocation -EBA $chkEBA -ChatGroup $chkChatGroup
                        
                        try 
                        {
                          
                        write-host $n "$($ADUser.Name)*$($AextAttrib)"
                        $n = $n + 1
                            # Update properties.
                            $ADUsr.extensionAttribute6 = $AextAttrib

 
                            # Update the user data in AD
                            Set-ADUser -Instance $ADUser 
                        }
                        catch
                        {
                        $dt=get-date
                        $caperr += "$($dt) | $($ADUser.Name) | $($ADUser.DistinguishedName) | $($_.Exception.Message)`n"
                        
                        }
                        }


  
}                            

#export error log
try
{
$caperr += "Total errors $(($caperr | Measure-Object -Line ).Lines)"
$caperr | Out-File -FilePath ChatterUserError.log -Force

}
catch
{
write-host "Could not write error log"
}


