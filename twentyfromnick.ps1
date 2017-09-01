
$User = $env:username
$user = $user.Split("@").item(0)
$Cred = Get-Credential -Message "Please enter your Talent2.corp password" -UserName $user
$Root = "\\talent2.corp\data"
New-PSDrive -Name "S" -Root $root -PSProvider FileSystem  -scope global -Persist  -Credential $cred
