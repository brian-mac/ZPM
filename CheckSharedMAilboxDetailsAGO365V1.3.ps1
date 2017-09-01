$RecruitmentUsers = Import-Csv -Path 'C:\temp\SHaredMailbox.csv'
$OutPutFile = "C:\Temp\ShareMailboxtStatus.csv"

foreach ($RecruitmentUser in $RecruitmentUsers)
{
    $Status = "All Correct"
    $CorpEmailAddress = $RecruitmentUser.FutureEmailAddress.Trim()
    $SamAccountName = Get-ADUser  -filter {EmailAddress -eq  $CorpEmailAddress} -Properties SamAccountName
    #$SamAccountName = ($CorpEmailAddress -split "@",2)[0]
    if ($SamAccountName -ne $Null)
    {
        $targetUser = get-O365mailbox -identity $SamAccountName.SamAccountName
    }
    if ($targetUser -ne $null) 
    {
        if ($targetUser.UserPrincipalName -like $RecruitmentUser.FutureEmailAddress)    
        {   
            if ($targetUser.PrimarySmtpAddress -notlike $RecruitmentUser.FutureEmailAddress)
            {
                  $Status = "Email incorrect"
            }
        }
        else
        {
            $Status = "UPN not correct"
        }
    }
    else
    {
        $Status = "Mailbox does not exist"
        $targetuser = Get-OnPremMailbox -Anr $RecruitmentUser.Displayname.trim()
        if ($targetUser -ne $null)
        { 
            $Status = "Exisitng Mail Box"   
        }
        else
        {
            $lastChance = $RecruitmentUser.Firstname.trim()
            if ($RecruitmentUser.LastName.Length -ne 0)
            {
                $lastChance= $lastChance.trim() + " " + $RecruitmentUser.LastName.Trim()
            }
            $targetuser = Get-OnPremMailbox -Anr $lastChance
            If ($targetUser -ne $null)
            {
                $Status = "Exisitng Mail Box" 
            }
        }
    }
   Add-Member -InputObject $RecruitmentUser -MemberType NoteProperty -Name UserPrincipleName -Value $targetUser.UserPrincipalName
   Add-Member -InputObject $RecruitmentUser -MemberType NoteProperty -Name PrimarySmtpAddress -Value $targetUser.PrimarySmtpAddress
   Add-Member -InputObject $RecruitmentUser -MemberType NoteProperty -Name ForwardingAddressO365 -Value $targetUser.ForwardingSmtpAddress
   Add-Member -InputObject $RecruitmentUser -MemberType NoteProperty -Name ForwardingAddressOnPremises -Value $targetUser.ForwardingAddress
   Add-Member -InputObject $RecruitmentUser -MemberType NoteProperty -Name ExtensionCustomAttribute14 -Value $targetUser.CustomAttribute14
   Add-Member -InputObject $RecruitmentUser -MemberType NoteProperty -Name Status -Value $Status
   $RecruitmentUser | Export-Csv -Path $OutPutFile  -force -NoTypeInformation -Append
   $targetUser= $Null
}  