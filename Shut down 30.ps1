$Time = 30

$timeout = new-timespan -Minutes $time
$sw = [diagnostics.stopwatch]::StartNew()
while ($sw.elapsed -lt $timeout){
    Write-Progress -id 98 -Activity "Self-Shutdown in $Time Minutes" -PercentComplete (((($time*60) - $sw.Elapsed.TotalSeconds)/($time*60))*100) -SecondsRemaining (($time*60) - $sw.Elapsed.TotalSeconds)
    Start-sleep -m 100
}

Write-Progress -id 98 -Activity "Shutdown"
Stop-Computer