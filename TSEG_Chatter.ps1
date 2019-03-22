#variables

$caperr=$null
$importfile="C:\temp\Chatter_users.csv" #change this back to $importfile="D:\Download\chatter_users.csv


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


