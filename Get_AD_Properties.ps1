function WriteLine ($LineTxt) 
{
     $Stream.writeline( $LineTxt )
}

# Create Log file stream
$mode = [System.IO.FileMode]::Append
$access = [System.IO.FileAccess]::Write
$sharing = [IO.FileShare]::Read
$LogPath = [System.IO.Path]::Combine("C:\temp\OutPut.csv")
# create the FileStream and StreamWriter objects
$fs = New-Object IO.FileStream($LogPath, $mode, $access, $sharing)
$Stream = New-Object System.IO.StreamWriter($fs)

#Open Input Data File
# $Inputfile = "C:\temp\shared.csv"
# $SharedMailBoxes = import-csv -Path $Inputfile

$UserSet = Get-ADUser -filter * -properties *  

foreach ($User in $UserSet)
{
    $line = $user.extensionAttribute3 + ";" +  $user.UserPrincipalName + ";"+ $user.DistinguishedName
    WriteLine $Line
}

#Close and write stream to file
$Stream.Close()
$fs.Close()
