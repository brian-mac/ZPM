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
    # $Line = $Date +  " PSSession and log file closed."
    #WriteLog $Line $stream 
    Write-Host $Date  +  " PSSession and log file closed."
    $Stream.Close()
    $FileSystem.Close() 
    # Close PS Sessions
    #Get-PSSession | Remove-PSSession
    $error.clear()
    #Exit
}
# Start main
$OutPath = "C:\temp\MFA-Disabled-Users.csv" 
$OutputFile = CreateFile  $OutPath "Create"
$OutPutStream = New-Object System.IO.StreamWriter($OutputFile)
$DataLine = "GivenName,Surname,UserPrincipalname,Location1,Location2,Email,Phone,Depertment"
WriteData $DataLine $OutPutStream

$MFADisabledMembers = get-adgroupmember -identity "gg-aa-mfa disabled"
foreach ($User in $MFADisabledMembers )
{
    $TargetUser = Get-ADUser $user.samaccountname -properties EmailAddress, telephoneNumber, StreetAddress, physicalDeliveryOfficeName, Department
    $UserProperties = $TargetUser.PropertyNames
    foreach ($Field in $UserProperties)
    {
        if ($null -eq $TargetUser.$Field)
        {
            $TargetUser."$Field" = " "
        }
    }
    $degugcounter ++
    $TargetUser.physicalDeliveryOfficeName  = $TargetUser.physicalDeliveryOfficeName -replace ',',' ' 
    $TargetUser.physicalDeliveryOfficeName = $TargetUser.physicalDeliveryOfficeName -replace "\n",' '
    $TargetUser.physicalDeliveryOfficeName = $TargetUser.physicalDeliveryOfficeName -replace "\r",' '
    $TargetUser.StreetAddress = $TargetUser.StreetAddress -replace ',',' ' 
    $TargetUser.StreetAddress = $TargetUser.StreetAddress -replace "\n",' '
    $TargetUser.StreetAddress = $TargetUser.StreetAddress -replace "\r",' '
    $DataLine = "$($TargetUser.GivenName),$($TargetUser.Surname),$($TargetUser.UserPrincipalname),$($TargetUser.streetaddress),$($TargetUser.physicalDeliveryOfficeName),$($TargetUser.emailaddress),$($TargetUser.telephonenumber),$($TargetUser.department)"
    WriteData $DataLine $OutPutStream
}
$degugcounter = $null
CloseGracefully $OutPutStream $OutputFile
