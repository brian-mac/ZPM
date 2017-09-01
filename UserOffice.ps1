import-module activedirectory



$Users = Get-ADUser -filter 'Enabled -eq $True'  -properties physicalDeliveryOfficeName , samaccountname
Add-Member -InputObject $users.Item(0)   -MemberType NoteProperty -Name physicalDeliveryOfficeName      -Value "Test" -force

$users |  Export-Csv -Path "C:\temp\UserOffice.csv"