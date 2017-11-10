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
$Line = "User_ID,TargetUser,Ext5,Status,Ext1,Ext2,Ext3"
WriteLog $line $LogStream

# Create Error log - TXT. Output: Any errors.
$ErrorFile = CreateFile "C:\temp\CheckSynch.log" "Create"
$ErrorStream = New-Object System.IO.StreamWriter($ErrorFile)
$user_error = $False
foreach ($User_Email in $Migrated_USers)
{
    $TargetEmail = $User_Email.SourceEmail
    $Target_user = Get-ADUser -filter {mail -eq $TargetEmail} -Properties * 
    If ($target_user)
        {
            if ( $Target_user.enabled -ne $True)
            {
                $Line = $target_user.Name + "," + $User_Email.SourceEmail + "," + $Target_user.extensionAttribute5 + "," + "Account Disabled" + "," + $Target_user.extensionAttribute1 + "," +  $Target_user.extensionAttribute2 + "," +  $Target_user.extensionAttribute3 
            }
            Elseif ( $Target_user.extensionAttribute5 -eq $Null )
            {
                $Line = $target_user.Name + "," + $User_Email.SourceEmail + "," + $Target_user.extensionAttribute5 + "," + "Not Valid" + "," + $Target_user.extensionAttribute1 + "," +  $Target_user.extensionAttribute2 + "," +  $Target_user.extensionAttribute3 
            }
            else
            {
                $Line = $target_user.Name + "," + $User_Email.SourceEmail + "," + $Target_user.extensionAttribute5 + "," + "Valid" + "," + $Target_user.extensionAttribute1 + "," +  $Target_user.extensionAttribute2 + "," +  $Target_user.extensionAttribute3 
            }
            WriteLog $Line $LogStream
        }
        else 
        {
            $ErrorLine = "Could not find $User_Email.SourceEmail"
            WriteLine $ErrorLine $ErrorStream
        }   
        
    $User_Error = $False
}
CloseGracefully $ErrorStream $ErrorFile
CloseGracefully $LogStream $LogFile
