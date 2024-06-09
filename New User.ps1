#Выбор: запуск на сервере Exchange или будет подключаться к указанному серверу Exchange, аунтификация по текущей сессии (можно и с запросом данных сделать)
$Se = Read-Host "Удаленная сессия Y/N"

if (($se -eq 'Y') -or ($se -eq 'y')){
#Подключение удаленной сессии powershel на exchange
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange01.domain.com/powershell/ -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking -AllowClobber
}else{Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn}

Import-Module ActiveDirectory

$pwds = "Какой то пароль"
#Вывод данных по активным почтовым базам
Get-MailboxDatabase -Status | select-object Name,Server,DatabaseSize,Mounted | Sort-Object -Property databasesize | where {$_.name-like "mailbox0*"} | ft -AutoSize
Write-Host " "
#Выбираем базу для создания учетки
$NDB = Read-Host "Введите номер почтовой базы (последняя цифра)"
$mbdb = "mailbox0$NDB"
$pwd = ConvertTo-SecureString $pwds -AsPlainText -Force

$fio  = ""
$auto = "99"
while ($auto -ne "0") {
$sam  = ""
#Выбираем тип создания уч. записей ( через переменную auto, дальше на основе выбора скрипт будет брать либо из файла данные, либо просить ввести
Write-Host " "
Write-Host "===========================Меню создания уч. записи=============================="
Write-Host " "
Write-Host "1.  Ручной ввод
2.  Загрузка из файла
0.  Выход 
"
$auto = Read-Host "Введите значение"
while ($auto -notmatch "(0|1|2)" ) {
    $auto = Read-Host "Номера нет в списке, укажите с 0 по 2"
}
if ($auto -ne "0") {
while ($sam -eq "") {
#Путь к файлу с данными по учеткам 
Import-CSV "\\server.domain.com\common4$\1C_Common\Upload\mail\NewADusers.utf8" -Delimiter ";" -Encoding UTF8 | ForEach-Object {

$mail = ""
$sam  = ""
$fio  = ""

if ($auto -eq "1") {
Write-Host " "
Write-Host "=========================В какое OU поместить уч. запись======================="
Write-Host " "
Write-Host "1.  Главный офис
2.  Второй офис
3.  Виртуальная машина
4.  Сф
5.  Севастополь
6.  БГУ
7.  Геленджик
8.  Гор.
9.  Кемерово
10. Пермь
11. Владивосток
12. Магадан
13. S7
14. Калинград
15. Сторонние компании.
0. Проекты
"
$OUid = Read-Host "Введите значение"

while ($ouid -notmatch "(0|1|2|3|4|5|6|7|8|9|10|11|12|13|15)" ) {
    $OUid = Read-Host "Номера нет в списке, укажите с 0 по 15"
}} else {$OUid = $_.ouid}

#Список общих пк для входа
$NewUsWork= "dc01,dc02,dc03,main,ADFS01,CAS,EXCHANGE01,EXCHANGE02,EXCHANGE03"
#Ввод ФИО сотрудника, с проверкой пустого поля 
if ($auto -eq "1") {$fio = Read-Host "Введите ФИО сотрудника"} else {$fio = $_.fio}
$i=0;while ($fio -eq "")  {If ($fio -eq ""){if ($i -lt 3) {$fio = Read-Host "ФИО не должно быть пустым, введите ФИО";$i++}else{break}}}
#Ввод имени учетки, с провркой на пустое поле
if ($auto -eq "1") {$sam = Read-Host "Введите SamAccountName"} else {$sam = $_.sam}
$i=0;while ($sam -eq "")  {If ($sam -eq ""){if ($i -lt 3) {$sam = Read-Host "Имя не может быть пустым, введите имя";$i++}else{break}}}
if ($sam -eq "") {Write-Host "Имя учетной записи не может быть пустым, создание записи отменено";break}
#Ввод имени пк для входа
if ($auto -eq "1") {$pc = Read-Host "Введите имя ПК"} else {$pc = $_.pc}
#По выбору выше выбирает нужную OU
If ($OUid -eq "1" ) { $ou = "domain.com/Domain Users and User's Computers/Domain Users/Internet/Normal" }
If ($OUid -eq "2" ) { $ou = "domain.com/Projects/S Plaza/Users" }
If ($OUid -eq "3" ) { $ou = "domain.com/Domain Users and User's Computers/Domain Users/Internet/Horizon" }
If ($OUid -eq "4" ) { $ou = "domain.com/Projects/Sf/Users" }
If ($OUid -eq "5" ) { $ou = "domain.com/Projects/Sevastopol/Users" }
If ($OUid -eq "6" ) { $ou = "domain.com/Projects/Rostov Gidrouzel (BGU)/Users"}
If ($OUid -eq "7" ) { $ou = "domain.com/Projects/Gelendzhik/Users"}
If ($OUid -eq "8" ) { $ou = "domain.com/Projects/Gor/Users"}
If ($OUid -eq "9" ) { $ou = "domain.com/Projects/Kemerovo/Users" }
If ($OUid -eq "10" ) {$ou = "domain.com/Projects/Perm/Users"}
If ($OUid -eq "11" ) {$ou = "domain.com/Projects/Vladivostok/Users" }
If ($OUid -eq "12" ) {$ou = "domain.com/Projects/Magadan/Users"}
If ($OUid -eq "13" ) {$ou = "domain.com/Projects/S7/Users" }
If ($OUid -eq "14" ) {$ou = "domain.com/Projects/Kaliningrad/Users" }
If ($OUid -eq "15" ) {$ou = "domain.com/Domain Users and User's Computers/Domain Users/Internet/Horizon" }
If ($OUid -eq "0" ) { $ou = "domain.com/Projects" }

#Поиск совпадений в AD

    $userF = (Get-ADUser -Filter {Name -like $fio}).Name
    $userS = (Get-ADUser -Filter {SamAccountName -like $sam}).SamAccountName
     
    If (($userF -eq $fio) -or ($userS -eq $sam)) {
    if ($userF -eq $fio) {Write-Host "ФИО: $fio - уже используется"}
    if ($userS -eq $sam) {Write-Host "Имя: $sam - уже используется"} 
    $n = Read-Host "Продолжить Y/N";if (($n -eq "N") -or ($n -eq "n")) {break} } else {Write-Host "ФИО: $fio - Свободно";Write-Host "Имя: $sam - Свободно"}

#Проверка ФИО   
    while ($userF -eq $fio)  {
        $fio = Read-Host "'$fio' - используется, введите другое ФИО сотрудника"
        If ($fio -eq ""){break}else{$userF = (Get-ADUser -Filter {Name -like $fio}).Name}}

        #If ($fio -eq ""){Write-Host "ФИО не может быть пустым, завершение цикла";break}else{Write-Host "ФИО: $fio - свободно"}

#Проверка логина
   

    while ($userS -eq $sam) {
        $name = (Get-ADUser -Filter {SamAccountName -like $sam}).name  
        $sam = Read-Host "'$sam' используется в '$name', введите другое имя"
        If ($sam -eq ""){Write-Host "Имя не может быть пустым, завершение цикла";break}else{$userS = (Get-ADUser -Filter {SamAccountName -like $sam}).SamAccountName}}        

        #If ($sam -eq ""){break}else{Write-Host "Logon: $sam - свободно"}

$alias = $sam
 
#Разбивка ФИО на состовляющие (фамилия и имя, для записи в соответствующие поля)
    $fio2 = $fio.split("")
    $fname = $fio2[1]
    $lname = $fio2[0]

$upn = $sam + "@domain.com"

#Для сторонних компаний есть выбор создать учетку с почтой или без
if ($OUid -eq "15" ) {if ($auto -eq "1") {$mail = Read-Host "Создать почту Y/N"}else{$mail = $_.mail}}

#Создание учетки
if (($OUid -in 0..14 ) -or ($mail -like "Y") -or ($mail -like "y")) {New-Mailbox -name $fio -userprincipalname $upn -Alias $alias -OrganizationalUnit $ou -SamAccountName $sam -FirstName $fname -LastName $lname -Password $pwd –Database $mbdb}
else {$ou = "OU=Horizon,OU=Internet,OU=Domain Users,OU=Domain Users and User's Computers,DC=domain,DC=com";New-ADUser -Name $fio -DisplayName $fio -GivenName $fname -Surname $lname -UserPrincipalName $upn -SamAccountName $SAM -Path $OU -AccountPassword (ConvertTo-SecureString $pwds -AsPlainText -force) -Enabled $true}

#Ждем обновление сервера, после создания учетки, если не выждать то ситема выдаст, что учетка не найдена при добавлении параметров 
Start-Sleep -Seconds 16

#добавление к списку стандартных пк то, что ввели
$NewUsWork += "," + $pc

#Если создаем учетную запись VM то добавляется группа и серверы в список для входа
if (($ouid -eq "3") -or ($ouid -eq "15"))
    {
        $NewUsWork += "," + "hcon01" + "," + "hcon02" 
        Add-AdGroupMember -Identity "HORIZON Block policy for Users and Computers" -members $sam
        Write-Host " "
        Write-Host "Добавлена группа:"
        write-host "HORIZON Block policy for Users and Computers"
        #Если пункт сторонняя компания, то нужно еще ввести название компании и можно через сколько дней отключить учетку (задасть дату отключения)
	if ($ouid -eq "15") {
        if ($auto -eq "1") {$company = Read-Host "Введите название компании";$discrip = "";$discrip = Read-Host "Введите описание"; $dayYN = Read-Host "Установить дату отключения Y/N";if (($dayYN -eq 'Y')-or($dayYN -eq 'y')){$dayof=Read-Host "Укажите через сколько дней отключить"}}else{$company = $_.company;$discrip = $_.discrip;$dayYN = $_.dayyn;$dayof = $_.dayoff}
        if(($dayYN -eq "Y") -or ($dayYN -eq "y")){$dayoff = (Get-Date).AddDays(1+($dayof)).ToString('dd-MM-yyyy')}}
    }

#Если создаем учетную в Пермь то добавляем RODC10
if ($ou -eq "domain.com/Projects/Perm/Users")
    {
        Add-AdGroupMember -Identity "Allowed RODC Password Replication RODC10" -members $sam
        Write-Host " "
        Write-Host "Добавлена группа:"
        write-host "Allowed RODC Password Replication RODC10"
    }

#Если создаем учетную в Геленджик то добавляем RODC11
if ($ou -eq "domain.com/Projects/Gelendzhik/User")
    {
        Add-AdGroupMember -Identity "Allowed RODC Password Replication RODC11" -members $sam
        Write-Host " "
        Write-Host "Добавлена группа:"
        write-host "Allowed RODC Password Replication RODC11"
    }

#Если создаем учетную в Калинград то добавляем RODC14
if ($ou -eq "domain.com/Projects/Kaliningrad/Users")
    {
        Add-AdGroupMember -Identity "Allowed RODC Password Replication RODC14" -members $sam
        Write-Host " "
        Write-Host "Добавлена группа:"
        write-host "Allowed RODC Password Replication RODC14"
    }

#Если создаем учетную в Владивосток то добавляем RODC15
if ($ou -eq "domain.com/Projects/Vladivostok/Users")
    {
        Add-AdGroupMember -Identity "Allowed RODC Password Replication RODC15" -members $sam
        Write-Host " "
        Write-Host "Добавлена группа:"
        write-host "Allowed RODC Password Replication RODC15"
    }

#Если создаем учетную в Кемерово то добавляем RODC17
if ($ou -eq "domain.com/Projects/Kemerovo/Users")
    {
        Add-AdGroupMember -Identity "Allowed RODC Password Replication RODC17" -members $sam
        Write-Host " "
        Write-Host "Добавлена группа:"
        write-host "Allowed RODC Password Replication RODC17"
    }

#Если создаем учетную в Севастополь то добавляем RODC18
if ($ou -eq "domain.com/Projects/Sevastopol/Users")
    {
        Add-AdGroupMember -Identity "Allowed RODC Password Replication RODC18" -members $sam
        Write-Host " "
        Write-Host "Добавлена группа:"
        write-host "Allowed RODC Password Replication RODC18"
    }

#Если создаем учетную в Сфера то добавляем RODC19
if ($ou -eq "domain.com/Projects/Sf/Users")
    {
        Add-AdGroupMember -Identity "Allowed RODC Password Replication RODC19" -members $sam
        Write-Host " "
        Write-Host "Добавлена группа:"
        write-host "Allowed RODC Password Replication RODC19"
    }

#Если создаем учетную в Симонов то добавляем RODC23
if ($ou -eq "domain.com/Projects/S Plaza/Users")
    {
        Add-AdGroupMember -Identity "Allowed RODC Password Replication RODC23" -members $sam
        Write-Host " "
        Write-Host "Добавлена группа:"
        write-host "Allowed RODC Password Replication RODC23"
    }
#Записываем пк для входа учетке
Set-ADUser -Identity $sam -Replace @{userWorkstations=$NewUsWork} 

#Установка региональных настроек

if (($OUid -in 0..14 ) -or (($mail -like "Y") -or ($mail -like "y"))) {Set-Mailbox $sam -EmailAddresses @{Add="$sam@domain.com"};Set-Mailbox $sam -PrimarySmtpAddress "$sam@domain.com" -EmailAddressPolicyEnabled $false
if((($ouid -in 0..8)-or($ouid -in 13..15))-and(-not($ouid -eq 14))){Set-MailboxRegionalConfiguration -Identity $sam -TimeZone "Russian Standard Time" -Language ru-ru -LocalizeDefaultFolderName:$true -DateFormat dd.MM.yyyy -TimeFormat H:mm}else{
if($OUid -eq 9){Set-MailboxRegionalConfiguration -Identity $sam -TimeZone "North Asia Standard Time" -Language ru-ru -LocalizeDefaultFolderName:$true -DateFormat dd.MM.yyyy -TimeFormat H:mm}else{
if($OUid -eq 10){Set-MailboxRegionalConfiguration -Identity $sam -TimeZone "Ekaterinburg Standard Time" -Language ru-ru -LocalizeDefaultFolderName:$true -DateFormat dd.MM.yyyy -TimeFormat H:mm}else{
if($OUid -eq 11){Set-MailboxRegionalConfiguration -Identity $sam -TimeZone "Vladivostok Standard Time" -Language ru-ru -LocalizeDefaultFolderName:$true -DateFormat dd.MM.yyyy -TimeFormat H:mm}else{
if($OUid -eq 14){Set-MailboxRegionalConfiguration -Identity $sam -TimeZone "Kaliningrad Standard Time" -Language ru-ru -LocalizeDefaultFolderName:$true -DateFormat dd.MM.yyyy -TimeFormat H:mm}else{
if($OUid -eq 12){Set-MailboxRegionalConfiguration -Identity $sam -TimeZone "Magadan Standard Time" -Language ru-ru -LocalizeDefaultFolderName:$true -DateFormat dd.MM.yyyy -TimeFormat H:mm}}}}}}}

#Запись в учетку для сторонней компании: название компании и дату отключения

if (($OUid -eq 15 ) -and ($discrip -eq "")) {$discrip = $company}
if (($OUid -eq 15 ) -and (($dayYN -eq "Y") -or ($dayYN -eq "y"))) {set-aduser -Identity $sam -Company $company -Description $discrip -AccountExpirationDate $dayoff} else {if ($OUid -eq 15 ) {set-aduser -Identity $sam -Company $company -Description $discrip}}

#Создание папки пользователя на FS
$fs = '\\fs\users$\' + $sam
New-Item -Path $fs -ItemType Directory
#Назначение прав на папку для пользователя 
$UserPath = "domain\" + $sam
$Rights = "Read, ReadAndExecute, ListDirectory, Write, Modify"
$InheritSettings = "Containerinherit, ObjectInherit"
$PropogationSettings = "None"
$RuleType = "Allow"
$acl = Get-Acl $fs
$perm = $UserPath, $Rights, $InheritSettings, $PropogationSettings, $RuleType
$rule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $perm
$acl.SetAccessRule($rule)
$acl | Set-Acl -Path $fs
#Снова ждем обновления записанных данных
Start-Sleep -Seconds 12
#Выводит на экран, что записал учеткам
if (($OUid -in 1..14 ) -or (($mail -like "Y") -or ($mail -like "y"))) {
Get-MailboxRegionalConfiguration -Identity $sam | fl timezone
Write-Host "Почтовые адреса:"
write-host "Все адреса:" (get-Mailbox -Identity $sam ).EmailAddresses
write-host "Гравный адрес:" (get-Mailbox -Identity $sam ).PrimarySmtpAddress}
Write-Host " "
Write-Host "Список ПК для входа:"
$account=Get-ADUser -Filter {SamAccountName -eq $sam} -properties *
write-host $account.userWorkstations
Write-Host " "
#Для центрального офиса проверяет, если указанный пк в начальной OU То переносит его в нужную
if ($ou -eq "domain.com/Domain Users and User's Computers/Domain Users/Internet/Normal")
{if ($pc){
$oupc = (Get-ADComputer -Filter {Name -like $pc}).DistinguishedName
$oupc1 = "CN="+$pc+",CN=Computers,DC=domain,DC=com" 
if ($oupc -eq $oupc1)
{Get-ADComputer -Identity $PC | Move-ADObject -TargetPath "OU=Internet,OU=Domain Computers,OU=Domain Users and User's Computers,DC=domain,DC=com"
Write-Host "'$PC' перенесен в domain.com/Domain Users and User's Computers/Domain Computers/Internet"
} Else {Write-Host "'$PC' находится в '$oupc'"}}}
if ($auto -ne 2){break}
#Write-Host " ";Write-Host "=================================================================================";
}
}}
#При выходе проверяет, если было удаленное подключение, то закрывает сессию
}if (($se -eq 'Y') -or ($se -eq 'y')){Remove-PSSession $Session};break
