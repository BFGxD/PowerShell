#Помск писем в транспортных логах

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

$servers = @('exchange01','exchange02','exchange03') 

$kto = "" #кто писал
$komu = "n.shetinkin@stroytransgaz.com" #кому писал

$date = 1 #сколько дней отступить для поиска, 1 - сегодня, макс число зависит от настроек сервера.

$from_date = (Get-Date).AddDays(-$date+1).ToString('yyyy-MM-dd')
$to_date = (Get-Date).Adddays(+1).ToString('yyyy-MM-dd')

if (($komu -like "") -and ($kto -notlike "")) {$servers.foreach({ Get-MessageTrackingLog -server $_ -Start $from_date -End $to_date -Sender $kto}) | ft -AutoSize} else {
if (($komu -notlike "") -and ($kto -like "")) {$servers.foreach({ Get-MessageTrackingLog -server $_ -EventID "SEND" -Start $from_date -End $to_date -Recipients $komu}) | ft -AutoSize} else {
if (($komu -notlike "") -and ($kto -notlike "")) {$servers.foreach({ Get-MessageTrackingLog -server $_ -EventID "SEND" -Start $from_date -End $to_date  -Sender $kto -Recipients $komu}) | ft -AutoSize}}}

