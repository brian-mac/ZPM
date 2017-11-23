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
Function AddGroup($TargetUser,$TargetGroup)
{
    Try
    {
        Add-ADGroupMember -Identity $TargetGroup -Members $TargetUser
        $Line = "$TargetUser has been added to $TargetGroup"
        WriteLog $Line $LogStream
    }
    Catch
    {
        $Message = ($_.Exception.Message).ToString()
        $Errorline = $TargetUser.name  + "  " + $Message 
        WriteLog $Errorline $LogStream
    }
    if ($targetuser.objecttype -eq "user")
    {
        Set-aduser $TargetUser -exstentionsattribute10 = Get-Date
    }
    elseif (condition) 
    {
       Set-adobject $TargetUser -exstentionsattribute10 Get-Date     
    }
    Return $Line
}

Function CompletePreviousAccounts ($EnableGroup, $TargetDate)
{
    $CurrentMembers = Get-ADGroupMember $EnableGroup 
    foreach ($CurrentMember in $CurrentMembers)
    {
        if ( $CurrentMember.ExtenstionAttrinbute10 -le $TargetDate)
        {
            Addgroup $CurrentMember $ProcessedGroup
            $ProcessedUsersCounter++ 
            $LogLine = "$CurrentMember.Name has been added to $EnableGroup"
            WriteLog $LogLine $LogStream
        }
    }
    Return $ProcessedUsersCounter
}
# open a log file
$LogFile = CreateFile "C:\temp\CheckSynch.csv" "Create"
$LogStream = New-Object System.IO.StreamWriter($LogFile)

$TargetOu = "OU=Disabled Accounts,DC=talent2,DC=corp"
$EnableGroup = "BackupifyTemp-GApps"
$ProcessedGroup = "BackupifyComplete-GApps"

$addedToBackupifyCounter = 0
$IterationCounter = 0
$TargetDate = get-date 
$TargetDate = $TargetDate.addhours(-72)

# {} find users who are already in $EnableGroup
$AmountOfUsersToProcess = CompletePreviousAccounts $enableGroup $TargetDate
$AmountOfUsersToProcess =1 # test code

# for each user  {AddGroup, $Processed-GoogleApps}  çheck object versus User
$DisableObjects = get-adobject -filter * -SearchBase $TargetOu -Properties *
while ($addedToBackupifyCounter -le $AmountOfUsersToProcess)
{
    $DisabledObject = $DisableObjects.item($IterationCounter)
    $IterationCounter++
    If ( !($DisabledObject.memberof).contains($ProcessedGroup))
    {
     $GroupReturnValue =   AddGroup $DisabledObject $EnableGroup
    If ($GroupReturnValue)
    {
        $addedToBackupifyCounter++
    }
    }
}
# Find all users in target OU that are not in $processed-GoogleApps
# Loop until  $ProcessedCounter = 500
#  {AddGroup, $#EnableGroup}
# increment counter
# #close




