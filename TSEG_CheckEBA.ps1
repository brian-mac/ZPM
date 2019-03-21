

$SourceUsers = import-csv "C:\temp\Not_eba.csv"
foreach ($User in $SourceUsers)
{ 
    $AdUser = get-aduser -properties employeenumber, extensionAttribute6 -filter "employeenumber -eq '$($user.employeenumber)'"
    If ($aduser)
    {
       add-member -InputObject $user -NotePropertyName "Exisits" -NotePropertyValue "True"
       add-member -InputObject $user -NotePropertyName "Extens6" -NotePropertyValue $aduser.extensionAttribute6
    }
    else
    {
        add-member -InputObject $user -NotePropertyName "Exisits" -NotePropertyValue "False"
        add-member -InputObject $user -NotePropertyName "Extens6" -NotePropertyValue $aduser.extensionAttribute6
    }
}
$SourceUsers |Export-Csv -path "C:\temp\UserStatus.csv"   
