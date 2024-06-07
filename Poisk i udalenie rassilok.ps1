#Скрипт автоматической очистки ящиков пользователей от массовых рассылок по компании. Есть эталонный ящик включеный во все крупные группы рассылок. На основе данных о письмах в этом ящике происходит поик и удаление писем в заданном диапозоне времени. Шаг день, 30 дней назад. Все динамиское: кол-во баз, пользователей, писем. После выполнения скрипт пришлет отчет о проделанной работе.

Import-Module ActiveDirectory

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange01.domain.com/powershell/ -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking -AllowClobber

$FilePath = '\\cluster.domain.com\common4$\1C_Common\Upload\mail\logs\spam' 
New-PSDrive -Name file -PSProvider FileSystem -Root $FilePath

$namelog = "Letters"
$date_start = 31 
$date_to = 30
$date_stop = 30 # на каком дне остановиться
$kdate = 1 # шаг в днях

$log = "file:\$namelog.csv"

If (Test-Path $log )  {
$copy = 'file:\backup\'+$namelog+' '+(get-date).ToString('dd MM yyyy')+'.csv'
cpi $log $copy
}

$logSuccessResults = "file:\$namelog"+"Results.csv"

If (Test-Path $logSuccessResults )  {
$copy = 'file:\backup\'+$namelog+'Results '+(get-date).ToString('dd MM yyyy')+'.csv'
cpi $logSuccessResults $copy
}

set-Content -Path $filepath'\'$namelog".csv" -Value 'Отправитель;тема;Найдено;Размер(MB);Поиск;Удаление'
set-Content -Path $filepath'\'$namelog"Results.csv" -Value 'Отправитель;тема;Получатель;Найдено;Размер(МБ)'


New-ComplianceSearch -Name $namelog -ExchangeLocation newsletters@domain.com -ContentMatchQuery "(Received:$((get-date).AddDays(-$date_start).ToString('yyyy-MM-dd'))..$((get-date).AddDays(-$date_to).ToString('yyyy-MM-dd')))"
Start-ComplianceSearch -Identity $namelog
$seach = Get-ComplianceSearch -Identity $namelog
while ($seach.status -ne "Completed") {Start-Sleep -Seconds 15;$seach = Get-ComplianceSearch -Identity $namelog}

if ((Get-ComplianceSearch $namelog).items -gt 0) {

New-ComplianceSearchAction -SearchName $namelog -Preview -Confirm:$false

$find_letter = $namelog + "_Preview"
$seach = Get-ComplianceSearchAction -Identity $find_letter
while ($seach.status -ne "Completed") {Start-Sleep -Seconds 15;$seach = Get-ComplianceSearchAction -Identity $find_letter}
(Get-ComplianceSearchAction $find_letter -Details).Results -replace '{', "Location,Sender,Subject,Type,Size,ReceivedTime,DataLink`r`n" -replace '}' -replace 'Location: ' -replace '; Sender: ', ',' -replace '; Subject: ', ',' -replace '; Type: ', ',' -replace '; Size: ', ',' -replace '; Received Time: ', ',' -replace '; Data Link: ', ',' -replace ",`r`n", "`r`n" | Out-File c:\temp\letters.csv
remove-ComplianceSearch -Identity $namelog -Confirm:$false
[Array]$Letters_list = Import-CSV "c:\temp\letters.csv" -Delimiter "," -Encoding UTF8  

$i = 1
$s = ""
$mb = 0
$items = 0
$ii = 0

$ii = 0
$baseuserall = ""

$base = Get-MailboxDatabase -Status | select-object Name | where {$_.name-like "mailbox0*"}
$groupnames = ""

$l = "["
$lx = ""
$shkala = 30 #ширина шкалы
for ($i=0;$i -lt $shkala;$i++){$lx += " "}
$ln ="]"
$i = 0
$sb = ""

$base.ForEach({if($sb -eq ""){$sb+=$_.name}else{$sb+=","+$_.name}})


$base.foreach({

$r=0
if ($i -eq 0){$k = 0}else{$k = $i / $base.Count * 100;$lx="";for ($o =1; $o -le $shkala; $o++ ){ if ($o -le $k/10*($shkala/10)){$lx +="*"}else{$lx +=" "}}}
$i++
#clear
$NewGroupName = $_.name+'users'

write-host "Список баз:$sb"
write-host ""
write-host "Общий прогресс:$l$lx$ln $k%  Заполнение группы:$NewGroupName"
write-host ""
    
    if($groupnames -eq ""){$groupnames="$NewGroupName@stg.local"}else{$groupnames+=",$NewGroupName@stg.local"}
    if (((Get-DistributionGroup $NewGroupName -ErrorAction 'SilentlyContinue').IsValid) -eq $true) {} else {
        New-DistributionGroup -Name $NewGroupName -SamAccountName $NewGroupName -OrganizationalUnit “domain.com/Domain Groups mail” -DisplayName $NewGroupName -Alias $NewGroupName  -Type "Security" -ManagedBy secofr07
        start-sleep -Seconds 15
        Set-DistributionGroup -Identity $NewGroupName -HiddenFromAddressListsEnabled $true 
    }
    $bc = (Get-Mailbox -Database $_.name -ResultSize:Unlimited).count
    $gc = (Get-DistributionGroupMember -id $NewGroupName -resultsize:unlimited).count
    if ($bc -eq $gc) {} else {
    $exloc = Get-DistributionGroupMember -id exdelmailspam@stg.local
    $baseuser = Get-Mailbox -Database $_.name -resultsize unlimited
    Update-DistributionGroupMember –Identity $NewGroupName –Member (Import-Csv "\\cluster.domain.com\common4$\1C_Common\Upload\mail\0.csv").username -confirm:$false
    
    foreach ($member in $baseuser){$iq = 0;foreach ($member1 in $exloc){If ($member1.SamAccountName -eq $member.alias) {$iq = 1}}if ($iq -eq 0) {$memberid=$member.alias;Add-DistributionGroupMember $NewGroupName -member $memberid}}}

})

$base.foreach({
$NewGroupName = $_.name+'users'
if($groupnames -eq ""){$groupnames="$NewGroupName@stg.local"}else{$groupnames+=",$NewGroupName@stg.local"}
})


$baseuseralllist = $groupnames.split(',')

write-host "Список баз:$sb"
if ($i -eq 0){$k = 0}else{$k = $i / $base.Count * 100;$lx="";for ($o =1; $o -le $shkala; $o++ ){ if ($o -le $k/10*($shkala/10)){$lx +="*"}else{$lx +=" "}}}
write-host "Общий прогресс:$l$lx$ln $k%"
$s = ""
$i = 1

$Letters_list.foreach({
    $ii = 0
    $senderk = $_.sender                 
    if ($senderk -eq "Lotus01") {$from = "from:Lotus01/domain@domain.com"} else {$from = "from:"+(Get-ADuser -Filter {displayName -like $senderk}).userPrincipalName}
    $subj = "Subject:"+$_.Subject
    $base.foreach({
        $basename = $_.name
        New-ComplianceSearch -Name "$Namelog$basename$i" -ExchangeLocation $baseuseralllist[$ii] -ContentMatchQuery "(Received:$((get-date).AddDays(-$date_start).ToString('yyyy-MM-dd'))..$((get-date).AddDays(-$date_to).ToString('yyyy-MM-dd')) AND ($from) AND ($subj))"
        if($s -eq ""){$s+="$Namelog$basename$i"}else{$s+=",$Namelog$basename$i"}
        $i++
        $ii++
    })
})


$list = $s.Split(',')

$l = "["
$lx = ""
$shkala = 30 #ширина шкалы
for ($i=0;$i -lt $shkala;$i++){$lx += " "}
$ln ="]"
$i = 0
$iii = 0
$ll = 0
$jh = 0

$list.foreach({

$r=0

if ($jh -eq 4) {$ll++;$jh=0}


if ($i -eq 0){$k = 0}else{$k = $i / $list.Count * 100;$lx="";for ($o =1; $o -le $shkala; $o++ ){ if ($o -le $k/10*($shkala/10)){$lx +="*"}else{$lx +=" "}}}

#clear

write-host "Список задач:$s"
write-host ""
write-host "Общий прогресс:$l$lx$ln $k%  Текущая задача:$_"
write-host ""

Start-ComplianceSearch -Identity $_
$seach = Get-ComplianceSearch -Identity $_
while ($seach.status -ne "Completed") {Start-Sleep -Seconds 15;$seach = Get-ComplianceSearch -Identity $_}
$DataSearch = Get-ComplianceSearch -Identity $_
(Get-ComplianceSearch $_).SuccessResults -replace '{', "Location,Item,Size`r`n" -replace '}' -replace 'Location: ' -replace ', Item count: ', ',' -replace ', Total size: ', ',' -replace ",`r`n", "`r`n" | Out-File c:\temp\Findlist.csv

if ((Get-ComplianceSearch $_).items -eq 0) {$dels = "Completed"} else {

New-ComplianceSearchAction -SearchName $_ -Purge -PurgeType SoftDelete -Confirm:$false
#New-ComplianceSearchAction -SearchName $_ -Preview -Confirm:$false
$del_letter = $_ + "_Purge"
#$del_letter = $_ + "_Preview"
$seach = Get-ComplianceSearchAction -Identity $del_letter
while ($seach.status -ne "Completed") {Start-Sleep -Seconds 15;$seach = Get-ComplianceSearchAction -Identity $del_letter}
$del = Get-ComplianceSearchAction -Identity $del_letter #| FL name,status  
$dels = $Del.Status
}
remove-ComplianceSearch -Identity $_ -Confirm:$false

$dsi = $DataSearch.Items
$dsm = [math]::Round($DataSearch.size/1MB,2)
$dss = $DataSearch.Status
$namereq = $_
[array]$res = Import-CSV "C:\Temp\Findlist.csv" -Delimiter "," -Encoding UTF8

$sendn = $Letters_list[$ll].sender
$subn = $Letters_list[$ll].Subject

$res.ForEach({
    if ($_.item -gt 0){
        $us = $_.location
        $item = $_.item
        $smb = [math]::Round($_.size/1MB,2)
        Add-Content -Path $filepath'\'$namelog'Results.csv' -Value "$sendn;$subn;$us;$item;$smb"}
})

Add-Content -Path $filepath'\'$namelog".csv" -Value "$sendn;$subn;$dsi;$dsm;$dss;$dels"

$mb += $DataSearch.size/1MB
$items += $DataSearch.Items
if ($list[$list.Count-1] -eq $_) {$mb = [math]::Round($MB); Add-Content -Path $filepath'\'$namelog".csv" -Value "$namelog;$Items;$MB;-;-"}

$i++
$jh++
})

#clear
write-host "Задача: $namelog | Удалено: $items писем | Размер: $mb MB |общий прогресс: 100%"

} else {remove-ComplianceSearch -Identity $namelog -Confirm:$false;Add-Content -Path $filepath'\'$namelog".csv" -Value "Нет писем;Нет писем;0;0;Complite;Complite";write-host "Задача: $namelog | Удалено: 0 | Размер: 0 |общий прогресс: 100%";$Letters_list = "";$mb = 0;$items = 0} 


$sr = "Дата проверки: "
$sr += (get-date).AddDays(-$date_start).ToString('yyyy-MM-dd')
$srk = "<h1>Результаты очистки:</h1>
<p>&nbsp;</p>
<table width='100%'>
<tbody>
<tr>
<td style='width: 1%;'>
<p>&nbsp;</p>
</td>
<td style='width: 6.9149%;'>
<p><strong>Дата проверки:</strong></p>
</td>
<td style='width: 15.0851%;'>
<p>$sr</p>
</td>
</tr>
<tr>
<td style='width: 1%;'>
<p>&nbsp;</p>
</td>
<td style='width: 6.9149%;'>
<p><strong>Найдено писем:</strong></p>
</td>
<td style='width: 15.0851%;'>
<p>$items</p>
</td>
</tr>
<tr>
<td style='width: 1%;'>
<p>&nbsp;</p>
</td>
<td style='width: 6.9149%;'>
<p><strong>Общий размер:</strong></p>
</td>
<td style='width: 15.0851%;'>
<p>$mb МБ</p>
</td>
</tr>
<tr>
<td style='width: 27%;' colspan='3'>
<p>&nbsp;</p>
</td>
</tr>"

$iii = 1
$Letters_list.ForEach({
$srk += "<tr>
<td style='width: 1%;'>
<p>&nbsp;</p>
</td>
<td style='width: 6.9149%;'>
<p><strong>Запрос №"+$iii+":</strong></p>
</td>
<td style='width: 15.0851%;'>
<p>Отправитель:&nbsp;"+$_.Sender+"</p>
</td>
<td style='width: 30.0851%;'>
<p>Тема:&nbsp;"+$_.Subject+"</p>
</td>
</tr>"
$iii++

})


$srk += "</tbody>
</table>"


$result_recipient = 'it@domain.com' # Ящик получателя отчета
$result_sender = 'Delete.letters@domain.com' # Ящик отправителя отчета
$smtp_server = 'cas.domain.com' # SMTP-серввер для отправки отчета
$email_subject = 'Удаление рассылок' # Тема отправляемого письма

$email_body = $srk

$filepathl = $filepath+'\'+$namelog+".csv"
$filepathr = $filepath+'\'+$namelog+'Results.csv'

Send-MailMessage -SmtpServer $smtp_server -To $result_recipient -From $result_sender -Subject $email_subject -Body $email_body -BodyAsHtml -Encoding 'UTF8' -Attachments $filepathl,$filepathr

remove-PSDrive -Name file 

Remove-PSSession $Session 
