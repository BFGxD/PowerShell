Import-Module ActiveDirectory
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange01.domain.com/powershell/ -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking -AllowClobber

$path = "\\cluster.domain.com\common4$\1C_Common\Upload\mail\pst\" #сюда выгрузится почтовый архив

do {

$fio = ""
$sam = ""
$fio = Read-Host "Введите ФИО сотрудника"
if ($fio -eq "") { $sam = Read-Host "Введите SamAccountName" }  #Если ФИО пустое значение то попросит ввести имя учетной записи

If ($sam -eq ""){$sam = (Get-ADUser -Filter {displayname -like $fio}).SamAccountName }
If ($fio -eq ""){$fio = (Get-ADUser -Filter {SamAccountName -like $sam}).name }

$pathpst = $path + $sam + ".pst"  #Путь выгрузки + имя учетной записи для названия файла

if (Get-Mailbox -Anr $sam) {New-MailboxExportRequest -Mailbox $sam -FilePath $pathpst;} else {Write-host $sam "- нет ящика на Exch"} #Если ящик есть то выгружает, если нет то пишет "нет ящика"

$vse=Read-Host "Нажмите любую клавишу или 0 для выхода"} until ($vse -eq 0) #Цикл если надо пару ящиков выгрузить, иначе лучше воспользоваться скриптом со списом из файла.

Remove-PSSession $Session


