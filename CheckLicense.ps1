#$UserCredential = Get-Credential
#Connect-MsolService -Credential $UserCredential
$users = import-csv -Path "C:\temp\officeUsers.csv"
foreach ($user in $users)
{
    $UsersLicense = Get-MsolUser -UserPrincipalName $user.Globalid 
    $Licenses = $UsersLicense.licenses
    $EntCloud = $Licenses |where {$_.accountskuid -eq "ALLEGISCLOUD:ENTERPRISEPACK"}
    if ($EntCloud.AccountSku  -ne $null)
    {
        $E3Services = $EntCloud.servicestatus
       # Add-Member -InputObject $E3services -MemberType NoteProperty -Name UserPrincipleName -Value $user.Globalid 
    }
    Export-csv -InputObject $E3Services -force -NoTypeInformation  -append -Path "C:\temp\OutLicenses.csv"
}

