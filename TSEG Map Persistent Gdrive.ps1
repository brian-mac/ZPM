function SecurePassword ($Target_Path, $Action)
{
    <# Function will either encrypt and save a password to the file specified in the $Target_Path
    Or it will return an encrypted password from the file specified in $Target_Path.
    Returns password 
    #>

    If ($Action -eq "Store")
    {
        $Secure = Read-Host -AsSecureString
        $Encrypted = ConvertFrom-SecureString -SecureString $Secure -Key (244,102,80,104,223,19,65,130,183,11,132,245,74,147,46,142)
        $Encrypted | Set-Content $Target_Path  
    }
    elseif ($Action -eq "Retrieve")
    {
        $Secure = Get-Content $Target_Path | ConvertTo-SecureString -Key (244,102,80,104,223,19,65,130,183,11,132,245,74,147,46,142)
    }
    Return $Secure
}

$Secure_Path = ""
$Secure_password = securepassword $Secure_Path "Retrieve"
$User = ""
#$user = $user.Split("@").item(0)
$Cred = # Get-Credential -Message "Please enter your Talent2.corp password" -UserName $user
$Root = "\\ \"
New-PSDrive -Name "G" -Root $root -PSProvider FileSystem  -scope global -Persist  -Credential $cred
