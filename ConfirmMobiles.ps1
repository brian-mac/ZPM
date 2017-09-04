function WriteLine ($LineTxt) 
{
    $Date = get-date -Format G
    $Date = $Date + "    : "  
    $LineTxt = $date + $LineTxt  
    $Stream.writeline( $LineTxt )
}

$mode       = [System.IO.FileMode]::Create
$access     = [System.IO.FileAccess]::Write
$sharing    = [IO.FileShare]::Read
$LogPath    = [System.IO.Path]::Combine("C:\temp\Mobile_Information.csv")
$fs = New-Object IO.FileStream($LogPath, $mode, $access, $sharing)
$Stream = New-Object System.IO.StreamWriter($fs)

$Inputfile = "C:\temp\Mobile_Input.csv"
$InputFile = import-csv -Path $Inputfile
Writeline "StaffEmail,Location,Company,PortalRegistered,PortalMobile,T2Mobile,MobileMatch"
ForEach ($Entry in $Inputfile)
{
    $TargetAccount = $Entry.email
    $TargetMobile = $Entry.sms_number
    $TargetReg = $Entry.Registered_sms
    $TargetUser = get-aduser -filter 'extensionattribute5 -eq $targetAccount -and UserAccountControl -ne 514' -Properties *
    if ($TargetUser)
    {
        $userLoc = $TargetUser.physicalDeliveryOfficeName
        $UserMob = $TargetUser.Mobile
        $UserCompany = $TargetUser.Company
        $MobileFlag = $False
        if ($UserMob -eq $TargetMobile)
        {
            $MobileFlag = $True
        }
        $OutPut = "$TargetAccount,$UserLoc, $UserCompany,$TargetReg,$TargetMobile,$UserMob,$MobileFlag"
        Writeline $Output
    }
}
$Stream.Close()
$fs.Close()