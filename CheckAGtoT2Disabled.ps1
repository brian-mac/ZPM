Function CreateFile ($ReqPath, $ReqMode)
{   
    <# Create the FileStream and StreamWriter objects, returns Stream Object
    Append will append to an exisitng file or create a new one if it does not exisit
    Create will create a new file or overwrite an existing one
    The file will have todays date appened to it #>

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
function WriteLog ($LineTxt,$Stream) 
{
    $Date = get-date -Format G
    $Date = $Date + "    : "  
    $LineTxt = $date + $LineTxt  
    $Stream.writeline( $LineTxt )
}

function WriteLine ($LineTxt,$Stream) 
{
    #$Date = get-date -Format G
    #$Date = $Date + "    : "  
    #$LineTxt = $date + $LineTxt  
    $Stream.writeline( $LineTxt )
}
Function CloseGracefully($Stream,$FileSystem)
{
    # Close all file streams, files and sessions. Call this for each FileStream and FileSystem pair.
    $Line = $Date +  " PSSession and log file closed."
    WriteLine $Line $stream 
    Write-Host $Date  +  " PSSession and log file closed."
    $Stream.Close()
    $FileSystem.Close() 
    # Close PS Sessions
    Get-PSSession | Remove-PSSession
    $error.clear()
    #Exit
}

Function ObjectToCSVData ($UnpackObject)
{
    $ObjectMembers = $UnpackObject |get-member -membertype Properties
    Foreach($Property in $ObjectMembers)
    {   
        $name = $Property.Name
        $Unpack = $Unpack + $UnpackObject.$name + ","
    }
    Return $Unpack
}
Function CheckDisabled ($CheckUser, $AGUser)
{
    
    if ($CheckUser.DistinguishedName -like "*OU=Ex Employees*" -or $CheckUser.DistinguishedName -like "*OU=Disabled Users*")
    {
        $line = $Checkuser.UserPrincipalName + "," + "Disabled" 
    }
    Elseif ($CheckUser.ObjectClass -eq "User")
    {
        $CheckADUser = get-aduser $CheckUser
        if ($CheckADUser.Enabled -ne $True)
        {
            $line = $CheckADuser.UserPrincipalName + "," + "Disabled" 
        }
    }
    Return $Line
}
function ProcessUser ($ProUser, $Criteria)
{
    $Email = ($ProUser.UPN).trim()
    $ADOSam = ($ProUser.sam).Trim()
    $AltSamd = $ADOSam.Replace(" ",".")
    $MailAlias =  ($ProUser.Ext1).trim()
    $OtherHome = "smtp:" + $MailAlias
    $AGUPN = ($ProUser.UPN).trim()
    $ADUPN = ($ProUser.UPN).trim()
    $ADUPN = ($ADUPN.Split("@")).item(0) + "@talent2.corp"
    $ALTSam = (($PRoUser.UPN).Split("@")).item(0)
    $ALTSaf = (($ALTSam).split(".")).item(0)
    $ALTSaS = " " + (($ALTSam).split(".")).item(1)
    $ALTadupn = $ALTSaf + $ALTSas + "@Talent2.corp"
    $tempADObject = get-ADObject -filter {(UserPrincipalName -eq $ADUPN) -or (UserprincipalName -eq $ALTadupn)  -or (samaccountname -eq $ADOSam) -or (samaccountname -eq $AltSamd) -or (OtherHomePhone -eq $OtherHome)-or (email -eq $email) } -properties *  # -Properties UserPrincipalName otherHomePhone
    if (!$tempADObject)
    {
        $tempADObject = get-ADObject -filter {(extensionAttribute5 -eq $AGUPN)} -properties *
        If (!$tempADObject)
        {
            $tempADObject = get-ADObject -filter {(extensionAttribute1 -eq $MailAlias)} -properties *
        }
    }
    If ($tempADObject)         
    {
        $Cast = $tempADObject.gettype()
        if ($Cast.basetype.name -eq "ADentity" -or $Cast.basetype.name -eq "ADAccount")
        {
            # We have found a valid user 
            $DataLine =  CheckDisabled $tempADObject $ProUser   
            Return $DataLine 
        } 
        else
        {
            foreach ($IProUser in $tempADObject) # ($i=0; $i -le $targetuser.count(); $i++)
            {
                $DataLine =  CheckDisabled $IProUser $ProUser #ProcessUser $IProUser 
            }
        }
    }
    else
    {
        $dataLine = ",Can not find"    
    }
    Return $DataLine
}


# Import valid migrated user list.
$AGDomain_Users = Import-Csv "C:\Temp\UserOutput.csv"

# Create output file -CSV.  Output: 
#$LogFile = CreateFile "C:\temp\CheckSynch.csv" "Create"
#$LogStream = New-Object System.IO.StreamWriter($LogFile)
#$Line = "User_ID,TargetUser,Ext5,Status,Ext1,Ext2,Ext3"
#WriteLog $line $LogStream
$DataFile = CreateFile "C:\temp\AGT2Disabled.csv" "Create"
$DataStream = new-object System.IO.StreamWriter($DataFile)

foreach ($AGuser in $AgDomain_Users) 
{
    #$ADOName = (($Aguser.UPN).split("@")).item(0)
   $StatusResult = ProcessUser $AGuser 0
   $ObjectData = ObjectToCSVData ($AgUser)
   $Objectdata = $ObjectData + $StatusResult
   Writeline $ObjectData $DataStream
   $Datastream.flush()
}
CloseGracefully $DataStream $DataFile

 