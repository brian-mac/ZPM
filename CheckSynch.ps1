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
function WriteLine ($LineTxt,$Stream) 
{
    $Date = get-date -Format G
    $Date = $Date + "    : "  
    $LineTxt = $date + $LineTxt  
    $Stream.writeline( $LineTxt )
}

function WriteLog ($LineTxt,$Stream) 
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
# Import valid migrated user list.
$Migrated_Users = Import-Csv "C:\Temp\MigrationUsers.csv"

# Create output file -CSV.  Output: target user, Ext5 setting, valid status.
$LogFile = CreateFile "C:\temp\CheckSynch.csv" "Create"
$LogStream = New-Object System.IO.StreamWriter($LogFile)
$Line = "User_ID,TargetUser,Ext5,Status"
WriteLog $line $LogStream

# Create Error log - TXT. Output: Any errors.
$ErrorFile = CreateFile "C:\temp\CheckSynch.log" "Create"
$ErrorStream = New-Object System.IO.StreamWriter($ErrorFile)

foreach ($User_Email in $Migrated_USers)
{
    try
    {
        $TargetEmail = $User_Email.SourceEmail
        $Target_user = Get-ADUser -filter {mail -eq $TargetEmail} -Properties extensionAttribute5
        if ( $Target_user.extensionAttribute5 -ne $Null -or $Target_user.extensionAttribute5 -contains "@Talent2")
        {
            $Line = $target_user.Name + "," + $User_Email.SourceEmail + "," + $Target_user.extensionAttribute5 + "," + "Valid"
        }
        else
        {
            $Line = $target_user.Name + "," + $User_Email.SourceEmail + "," + $Target_user.extensionAttribute5 + "," + "Not Valid"
        }
        WriteLog $Line $LogStream
    }
    Catch 
    {
        $ErrorLine = "Could not find $User_Email.SourceEmail"
        WriteLine $ErrorLine $ErrorStream
    } 

}
CloseGracefully $ErrorStream $ErrorFile
CloseGracefully $LogStream $LogFile
