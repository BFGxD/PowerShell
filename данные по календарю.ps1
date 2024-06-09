#Для делегирования прав на ящик руководителя помощнику.

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

#Два варианта имени по ФИО и по имени учетной записи

$user = ""  #Руководитель
$sec = ""   #Помощник

$userfio = ""
$secfio = ""

If ($user -eq ""){$user = (Get-ADUser -Filter {displayname -like $userfio}).SamAccountName }
If ($userfio -eq ""){$userfio = (Get-ADUser -Filter {SamAccountName -like $user}).name }

If ($sec -eq ""){$sec = (Get-ADUser -Filter {displayname -like $secfio}).SamAccountName }
If ($secfio -eq ""){$secfio = (Get-ADUser -Filter {SamAccountName -like $sec}).name }

#Вывлд нужных данных
$cal = $user  + ':\Календарь'
Get-MailboxFolderPermission -Identity $cal
Get-MailboxRegionalConfiguration -Identity $user
Get-MailboxRegionalConfiguration -Identity $sec
Get-MailboxPermission -Identity $user | ft -AutoSize
Get-MailboxFolderPermission -Identity $user | ft -AutoSize

#настройка тайм зоны и языка
#Set-MailboxRegionalConfiguration -Identity $user -TimeZone "Russian Standard Time" -Language ru-ru -LocalizeDefaultFolderName:$true -DateFormat dd.MM.yyyy -TimeFormat H:mm
#Set-MailboxRegionalConfiguration -Identity $sec -TimeZone "Russian Standard Time" -Language ru-ru -LocalizeDefaultFolderName:$true -DateFormat dd.MM.yyyy -TimeFormat H:mm

#Доступ к календарю
#remove-MailboxFolderPermission -identity $cal -user $sec
#set-MailboxFolderPermission -identity $cal -user $sec -AccessRights PublishingEditor
#add-MailboxFolderPermission -identity $cal -user $sec -AccessRights PublishingEditor

#Доступ к почте
Add-MailboxPermission -Identity $user -User $sec -AccessRights FullAccess -AutoMapping:$true -InheritanceType All