Import-Module -Name ActiveDirectory
$users = Get-ADUser -filter "*" -Properties title
$TestUser = "Brian.mcelhinney"
$user = Get-ADUser -identity $TestUser -Properties title
$ADattribute = "extensionAttribute4"
#foreach ($user in $users)
#{
       Set-ADUser $User -Replace @{$ADattribute=$user.title}
#}
