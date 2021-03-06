#set-location "\\10.51.0.12\Tool Shop\recovery\Toolshop\Root\Script"

$iterations = 10000
#$Average = get-content -path "Runtimetest.txt"
$OldAverage=0

#$Runtime = get-content -path "runtimetest.txt"
$time = "{0:N1}" -f (($Runtime/1000)*$iterations)
write-host "Estimated completion time: $time seconds"
$Averagetime1 = (get-date)

for($i=0; $i -le $iterations; $i++){
Write-progress -activity "timing" -status "None" -percentcomplete (($i/$iterations)*100) -secondsremaining ($time-(($i/$iterations)*$time))
Write-host $i

start-sleep -m 10
$Averagetime = new-timespan $averagetime1 (Get-date)
$Averagetime = ($Averagetime).totalmilliseconds
Write-host "New time: $averagetime"
$Average = (($averagetime)+(($OldAverage)*$i))/($i+1)
Write-host "Average: $Average"


$OldAverage = $Average
$Averagetime1 = (get-date)
}
Write-progress -activity "timing" -status "None" -complete
<#
Remove-Item -Path "Runtimetest.txt" -Force
New-Item -Name "Runtimetest.txt" -ItemType File 
Add-Content -Path "Runtimetest.txt" -Value $Average
#>

#$Runtime = get-content -path "runtime.txt"
#write-host ($Runtime/60000)