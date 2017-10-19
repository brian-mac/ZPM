# Office 2016 Normal.dot hack I mean fix
#New-Item -Path HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\ -Name officeupdate –Force

# Set-Item -Path HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate\updatebranch -Value “Current”

#New-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate -Name updatebranch -PropertyType String -Value $Channel

New-Item         -Path HKCU:\Software\Microsoft\Office\Outlook\Addins\ -Name officeupdateMicrosoft.OutlookBackup.1 -Force    
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\Outlook\Addins\officeupdateMicrosoft.OutlookBackup.1 -Name RequireShutdownNotification -PropertyType DWord -Value 1
