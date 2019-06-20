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
Function ConnectMSOL
{
    try
    {
        Get-MsolDomain -ErrorAction Stop > $null
    }
    catch 
    {
        #if ($cred -eq $null) {$cred = Get-Credential -Message "Pleas eenter Azure Credentials" $O365Adminuser}
        Write-Host "Connecting to Office Azure..."
        Connect-MsolService # -Credential $cred
    }
}

Function GetGroupDetails($targetGroups)
{
    foreach ($MFAGroup in $targetGroups)
    {
        $GM = Get-MsolGroupMember -GroupObjectId  $MFAGroup.ObjectId -all
        foreach ($member in $GM)
        {
            $TargetUser = Get-MsolUser -UserPrincipalName $member.EmailAddress
            $AuthMethod = $TargetUser.StrongAuthenticationMethods 
            if ($AuthMethod.count) 
            {
                $Status = $true
            }   
            else
            {
                $Status = $false    
            }
            $DataLine = $TargetUser.UserPrincipalName + "," + $TargetUser.StrongAuthenticationMethods + "," + $Status + "," + $MFAGroup.DisplayName + "," + $TargetUser.Department
            WriteData $DataLine $OutPutStream
        }
    }
}
# Main
ConnectMSOL
$OutPath = "C:\temp\UserAuthMethod.csv" 
$OutputFile = CreateFile  $OutPath "Create"
$OutPutStream = New-Object System.IO.StreamWriter($OutputFile)
$DataLine = "UPN,StrongAuth,Status,MFAGroup,Department"
WriteData $DataLine $OutPutStream
#$MFAGroups = @("DYN-MFA Enabled","DYN-MFA Enabled2", "DYN-MFA Enabled3", "DYN-MFA Enabled4")
$MFAGroups = Get-MsolGroup -SearchString "DYN-SSPR Office Staff"
GetGroupDetails $MFAGroups
$MFAGroups = Get-MsolGroup -SearchString "DYN-SSPR Rostered Staff"
GetGroupDetails $MFAGroups
$MFAGroups = Get-MsolGroup -SearchString "DYN-SSPR-MFA-Rostered Staffs"
GetGroupDetails $MFAGroups
$MFAGroups = Get-MsolGroup -SearchString "DYN-SSPR-MFA-Non-Rostered Staff"
GetGroupDetails $MFAGroups

CloseGracefully $OutputStream $OutputFile
