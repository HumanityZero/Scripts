$TestT = $null
$testF = $null

$ErrorActionPreference = 'SilentlyContinue'

$time = 2
start-sleep -s $time

for(1){
    Start-sleep -s $time
    try{$test = test-path '\\10.51.0.12\tool shop'}catch{$test = $false}
    if($test){
        $testT = $testT + 1
    }
    else{$testF = $testF + 1
    }

    $Per = $testT/($testT + $testF)*100
    clear-host
    Write-Host "P: $TestT  F: $TestF   $("{0:N0}" -f $Per)% Passed"
    if($TestT -ge 100 -and $TestF -lt 2){"Test Successful"; exit}
    if($test){
        $time = $time*0.99
    }
    else{
        $time++
    }
    if($time -lt 0.2){"Test Successful";exit}
}
pause