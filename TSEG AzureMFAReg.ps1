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

# Main
$OutPath = "C:\temp\UserAuthMethod.csv" 
$OutputFile = CreateFile  $OutPath "Create"
$OutPutStream = New-Object System.IO.StreamWriter($OutputFile)
$DataLine = "UPN,StrongAuth,Status"
WriteData $DataLine $OutPutStream
$G = Get-MsolGroup -SearchString "DYN-MFA Enabled"
$GM = Get-MsolGroupMember -GroupObjectId  $G.ObjectId -all
foreach ( $member in $GM)
{
    $TargetUser = Get-MsolUser -UserPrincipalName $member.EmailAddress
    if ($TargetUser.StrongAuthenticationMethods -ne $null) 
    {
        $Status = $true
    }   
    else
    {
        $Status = $false    
    }
    $DataLine = $TargetUser.UserPrincipalName + "," + $TargetUser.StrongAuthenticationMethods + "," + $Status
    WriteData $DataLine $OutPutStream
}
CloseGracefully $OutputStream $OutputFile
