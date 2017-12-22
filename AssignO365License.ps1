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


Function ConnectMSOL ()
{
    $Sessions = Get-MsolDomain -erroraction SilentlyContinue
    if (!$sessions)
    {
        $UserCredential = Get-Credential
        Connect-MsolService -Credential $UserCredential
    }
}

ConnectMSOL #Connect to Azure
# Set up files
$LogFile = CreateFile "C:\temp\SharedAccountsE2.log" "Create"
$LogStream = New-Object System.IO.StreamWriter($LogFile)
$SharedAccounts = import-csv "C:\temp\shared.csv"

# Define variables and constants
$SharedOU = "OU=APAC,OU=Shared Accounts,OU=Special Accounts,OU=Users,OU=Enterprise,DC=allegisgroup,DC=com"
$E2= "ALLEGISCLOUD:EXCHANGEENTERPRISE"
$E3 = "ALLEGISCLOUD:ENTERPRISEPACK"


foreach ($Sharedaccount in $sharedAccounts)
{
    
    $CurrentNSOLU = get-msoluser -UserPrincipalName $Sharedaccount.UPN
    if (!$CurrentNSOLU)
    {
        $Line = "Can not find the account " + $Sharedaccount 
        writelog $line $LogStream
        $LogStream.Flush()
    }
    Else
    {
        $Licenses = $CurrentNSOLU.Licenses
        If ((!$Licenses) -or (!($CurrentNSOLU.Licenses.accountskuid.contains($E2))))
        {
            Set-MsolUserLicense -UserPrincipalName $Sharedaccount.UPN -AddLicenses $E2  
            $Line = $CurrentNSOLU.displayname + "added to " + $E2
            writelog $line $LogStream
            $LogStream.Flush()
        }
        foreach ($AccSKU in $CurrentNSOLU.licenses)
        {
            if (!($AccSKU.AccountSkuId -eq $e2) )
            {
                Set-MsolUserLicense -UserPrincipalName $Sharedaccount.UPN -RemoveLicenses $AccSKU.AccountSkuId 
                $Line = $CurrentNSOLU.displayname + "removed from " + $AccSKU.AccountSkuId
                writelog $line $LogStream
                $LogStream.Flush()
            }
        }
    }
}
CloseGracefully $LogStream $LogFile
