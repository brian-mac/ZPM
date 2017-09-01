function WriteLine ($LineTxt) 
{
     $Stream.writeline( $LineTxt )
}

# Create Log file stream
$mode = [System.IO.FileMode]::Append
$access = [System.IO.FileAccess]::Write
$sharing = [IO.FileShare]::Read
$LogPath = [System.IO.Path]::Combine("C:\temp\Delegates.csv")
# create the FileStream and StreamWriter objects
$fs = New-Object IO.FileStream($LogPath, $mode, $access, $sharing)
$Stream = New-Object System.IO.StreamWriter($fs)
#Open Data File
$Inputfile = "C:\temp\shared.csv"
$SharedMailBoxes = import-csv -Path $Inputfile
foreach ($SharedMailBox in $SharedMailBoxes)
{
    $Mailbox = $SharedMailBox.Email
    $Delgates =  $Sharedmailbox.InboxDelegatedTo.Split("|")
    foreach ($Delegate in $Delgates)
    {
           $Line = $MailBox +"," + $Delegate
            WriteLine $Line
    }
}

#Close and write stream to file
$Stream.Close()
$fs.Close()
