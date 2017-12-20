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
    $FilePath    = [System.IO.Path]::Combine($ReqPath)
    $fs = New-Object IO.FileStream($FilePath, $mode, $access, $sharing)
    Return $fs
}
function WriteLog ($LineTxt,$Stream) 
{
    $Date = get-date -Format G
    $Date = $Date + "    : "  
    $LineTxt = $date + $LineTxt  
    $Stream.writeline( $LineTxt )
}

function WriteData ($LineTxt,$Stream) 
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
    WriteLog $Line $stream 
    Write-Host $Date  +  " PSSession and log file closed."
    $Stream.Close()
    $FileSystem.Close() 
    # Close PS Sessions
    Get-PSSession | Remove-PSSession
    $error.clear()
    #Exit
}

function Convert-DateString ([String]$Date, [String[]]$Format)
{
   $result = New-Object DateTime
 
   $convertible = [DateTime]::TryParseExact(
      $Date,
      $Format,
      [System.Globalization.CultureInfo]::InvariantCulture,
      [System.Globalization.DateTimeStyles]::None,
      [ref]$result)
 
   if ($convertible) { $result }
}
Function AddGroup($TargetUser,$TargetGroup)
{
    Try
    {
        if ($targetuser.objectclass -eq "contact")
        {
            Set-ADGroup -Identity $targetgroup -Add @{'member'=$targetuser.DistinguishedName}
        }
        Else 
        {
            Add-ADGroupMember -Identity $TargetGroup -Members $TargetUser.DistinguishedName
        }
        $Line = "Sucsess:" + $TargetUser.Name + " has been added to " +  $TargetGroup
        WriteLog $Line $LogStream
        $LogStream.Flush()
    }
    Catch
    {
        $Message = ($_.Exception.Message).ToString()
        $Errorline = "Error:" + $TargetUser.name  + "  " + $Message 
        WriteLog $Errorline $LogStream
        $LogStream.Flush()
    }
    $ToDate = (get-date).tostring()
    if ($targetuser.objectclass -eq "user")
    {
        Set-aduser $TargetUser -replace @{'extensionAttribute10'=$ToDate}
    }
    elseif ($targetuser.objectclass -eq "contact") 
    {
       Set-adobject $TargetUser -replace @{'extensionAttribute10'=$ToDate}    
    }
    Return $line
}

Function CompletePreviousAccounts ($EnableGroup, $tProcessedGroup, $TargetDate)
{
    $CurrentMembers = Get-ADGroupMember $EnableGroup      
    foreach ($CurrentMember in $CurrentMembers)
    {
        $CurrentObj = Get-ADObject $CurrentMember -Properties extensionAttribute10 , Memberof
        $EX10 = $CurrentObj.extensionAttribute10
        $ObjectDate = $Ex10  | get-date  #Convert-DateString $EX10 "dd/MM/yyyy hh:mm:ss tt"
        # $ObjectDate = "20/11/2017" | get-date
        if ( $ObjectDate -le $TargetDate -And (!$CurrentObj.memberof.contains($tProcessedGroup.DistinguishedName)))
        {
            Addgroup $CurrentMember $tProcessedGroup 
            $ProcessedUsersCounter++ 
            #$LogLine = $CurrentMember.Name + " has been added to: " +  $EnableGroup
            #WriteLog $LogLine $LogStream
            #$LogStream.Flush()
        }
    }
    Return $ProcessedUsersCounter
}
# open a log file
$LogFile = CreateFile "C:\temp\BackUpIfy.log" "Create"
$LogStream = New-Object System.IO.StreamWriter($LogFile)

$TargetOu = "OU=Disabled Accounts,DC=talent2,DC=corp"
$EnableGroup = "BackupifyTemp-GApps"
$ProcessedGroup = "BackupifyComplete-GApps"
$ProGroupDN = get-adgroup "CN=BackupifyComplete-Gapps,OU=Google Sync Groups,OU=Talent2 Security Groups,DC=talent2,DC=corp"
$EnableGroupDN = get-adgroup "CN=BackupifyTemp-GApps,OU=Google Sync Groups,OU=Talent2 Security Groups,DC=talent2,DC=corp"
$addedToBackupifyCounter = 0
$IterationCounter = 0
$TargetDate = get-date 
$TargetDate = $TargetDate.addhours(-72)

# {} find users who are already in $EnableGroup
$AmountOfUsersToProcess = CompletePreviousAccounts $enableGroup $ProGroupDN $TargetDate
$AmountOfUsersToProcess = $AmountOfUsersToProcess.item(($AmountOfUsersToProcess.count)-1)
 # $AmountOfUsersToProcess = 530 # test code

# for each user  {AddGroup, $Processed-GoogleApps}  çheck object versus User
$DisableObjects = get-adobject -filter {objectclass -eq "user" -or objectclass -eq "contact"} -SearchBase $TargetOu -Properties extensionAttribute10 , Memberof | Where-object{($_.memberof -notcontains $ProGroupDN.DistinguishedName) -and ($_.memberof -notcontains $EnableGroupDN.DistinguishedName)} 
while ($addedToBackupifyCounter -le $AmountOfUsersToProcess)
{
    $IterationCounter++
    $DisabledObject = $DisableObjects.item($IterationCounter)
    #If ( (!$DisabledObject.memberof.contains($ProGroupDN)) -or (!$DisabledObject.memberof.contains($ProGroupDN)))
    #{
        $GroupReturnValue =   AddGroup $DisabledObject $EnableGroup
        If ($GroupReturnValue)
        {
            $addedToBackupifyCounter++
        }
    #}
}
# Find all users in target OU that are not in $processed-GoogleApps
# Loop until  $ProcessedCounter = 500
#  {AddGroup, $#EnableGroup}
# increment counter
CloseGracefully $LogStream $logfile 




