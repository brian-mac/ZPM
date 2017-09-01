Function CloseGracefully()
{
    $Stream.writeline( $Date +  " PSSession and log file closed.")
    Write-Host $Date  +  " PSSession and log file closed."
    $Stream.Close()
    $fs.Close()
    # Close PS Sessions
    Get-PSSession | Remove-PSSession
    $error.clear()
    Exit
}

#Main
#$ADcred = Get-Credential  -Message "Please enter your Talent2.corp user name and password"
#$ADSession = New-PSSession -ComputerName T2EDC-DC03 -Credential $ADcred
#Import-PSSession $ADSession  -ErrorAction Continue
Import-Module -Name ActiveDirectory

#$users = Get-ADUser -filter "*" -Properties title
$TestUser = "Brian.mcelhinney"
$mode = [System.IO.FileMode]::Append
$access = [System.IO.FileAccess]::Write
$sharing = [IO.FileShare]::Read
$LogPath = [System.IO.Path]::Combine($Env:AppData,"C:\Temp\extensionattribute4.log")
# create the FileStream and StreamWriter objects
$fs = New-Object IO.FileStream($LogPath, $mode, $access, $sharing)
$Stream = New-Object System.IO.StreamWriter($fs)

$InputFile = "C:\temp\extensionattribute4.csv"
$users = import-csv -Path $Inputfile


#$user = Get-ADUser -identity $TestUser -Properties title
$ADattribute = "extensionAttribute4"
foreach ($user in $users)
{
       If ($User.Enabled -eq "true")
       {
           Set-ADUser -Identity $User.SamAccountName -Replace @{$ADattribute=$user.extensionattribute4}
       }
       
}
CloseGracefully