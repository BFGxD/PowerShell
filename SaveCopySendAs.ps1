#Добавление копии письма в ящик от чего имени идет отправка письма

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

$mailuser = "" #Почтовый ящик

Set-Mailbox $mailuser -MessageCopyForSentAsEnabled $True
set-mailbox $mailuser -MessageCopyForSendOnBehalfEnabled $True


#Set-Mailbox $mailuser -GrantSendOnBehalfTo @{Add="","","",""}  #Добавление "отправка от имени" в общий ящик (отсутствует в ECP)




