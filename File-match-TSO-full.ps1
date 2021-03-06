﻿$random = Get-random
Start-Transcript -path "\\10.51.0.12\Tool Shop\recovery\Toolshop\Root\Script\Logs\$random.log" -append | out-null
if($PSVersionTable.PSVersion.major -lt 4){
    write-host "This script requires powershell version 4.0+"
    write-host "Update your computer to .net framework 4.0 to continue"
    read-host "Press enter to exit"
    exit
}
add-type -an system.windows.forms
write-host -ForegroundColor Black "Start: Win32ShowWindowAsync = Add-Type –memberDefinition @`nXXX`n $([System.Windows.Forms.Clipboard]::GetText())`nXXX`n END -name Win32ShowWindowAsync -namespace Win32Functions –passThru"
clear-host

cd "\\10.51.0.12\Tool Shop\recovery\Toolshop\Root\Script\"
. "\\10.51.0.12\Tool Shop\recovery\Toolshop\Root\Script\FlashWindow.ps1"
$search = Read-host "Enter full part name"
$department = Read-host "Enter two letter department name" 
$accuracy = read-host "Enter accuracy (press enter for default value of 81%)"
$2013 = read-host "2013 files? (y/n)"
$sw = [diagnostics.stopwatch]::StartNew()

$1st = @(0,0,0,0,0,0,0,0,0,0,0)
$2nd = 0
$averagetime = @("no data","no data","no data","no data","no data","no data","no data","no data","no data","no data","no data")
$lastmatch1 = $lastmatch2 = $lastmatch3 = "none"
$x=0
$department = $department.toupper()
$timeSearches = "\\10.51.0.12\Tool Shop\05. Share\35. Time Studies\"

if($accuracy -gt 1){$accuracy = (100-$accuracy)/100}
elseif($accuracy -lt 1 -and $accuracy -gt 0){$accuracy = 1-$accuracy}
else{$accuracy = (0.19)}
$percentage = $((1-$accuracy)*100)

For($number = 30; $number -ge 1; $number--){
    if($search -match "$number"){break}
}

if($2013 -eq "y"){
    $2013 = "Original - All as of 05.25.15\"
}else{$2013 = $null}

if(!$file){$file = "All"}

if($department){$message = "in $department"}

clear-host
write-host "Running`n`n`n`n`n`n`n`n`n`n"

$type = @("Multiaxial_Milling","CNC_Turning","Surface_Grinding","Circular_grinding","External_Heat_Treatment","Manual_Workstations","Wire_Cut_EDM","Material","Sawing","Standard_Milling","Standard_Turning")

$ErrorActionPreference = "silentlycontinue"

$data = import-csv "\\10.51.0.12\Tool Shop\recovery\Toolshop\Root\Script\Quote\TimeQuoteDataBase\$($2013)full.csv"
$total = $data.GetUpperBound(0)
$host.ui.RawUI.WindowTitle = "Time Search $total"
write-host "no results"
foreach ($item in $data){
    
    $x++
    $match1 = $item.Criterion8
    $match2 = $item.Criterion14
    $match3 = $item.criterion4

    Write-Progress -activity “Searching for: '$Search' with accuracy of $percentage% $message" -id 43 -PercentComplete (($x/($total))*100) -CurrentOperation "Last 3 matches: $($lastmatch1), $($lastmatch2), $($lastmatch3)" -Status "$([math]::Round(($x/$total)*100))% Complete ($x/$($total+1)) / Elapsed Time: $([math]::Round($sw.elapsed.totalseconds)) Seconds / Total Matches: $($matchlist.getupperbound(0)+1)"
    
    if($match3 -match $department -or $department -eq $null){
    $LDpercent = .\Get-ld.ps1 $search $match1 -i
    if($LDpercent -lt $accuracy){if($number -lt 1 -or $match1 -match $number -and $match1 -notmatch "1$number" -and $match1 -notmatch "2$number"){
        for($j =0; $j -lt 11; $j++){
            $LDpercent2 = .\Get-ld.ps1 $type[$j] $match2 -i
            if($LDpercent2 -lt 0.1){
                [float]$newav = (($item.Criterion20)/($item.Criterion10))
                if($newav -gt 0 -and $newav -lt 200){
                    $matchlist += @("$match3 $match1 --> $($type[$j]) $newav")
                    if($1st[$j] -eq 0){
                        $averagetime[$j] = $newav
                    }
                    else{
                        $averagetime[$j]=[math]::round((($averagetime[$j]*$1st[$j]+$newav)/($1st[$j]+1)),2)
                    }
                    $1st[$j] = $1st[$j]+1
                    $lastmatch3 = $lastmatch2
                    $lastmatch2 = $lastmatch1
                    $lastmatch1 = $match1
                    clear-host
                    write-host "Running`n`n`n`n`n`n`n`n`n`n`n"
                    for($jp =0; $jp -lt 11; $jp++){
                        if($averagetime[$jp] -ne "no data"){
                            Write-host $type[$jp] "($($1st[$jp])): " $averagetime[$jp]
                        }
                    }
                    $tru=$true
                    break
                }
            }}
        }}
    }
}
Write-Progress -id 43 -complete -Activity "x"
$sw.stop

clear-host
write-host "Complete`n$search`nSearch arccuracy: $percentage%`n$($sw.elapsed.minutes) Minutes $($sw.elapsed.seconds) Seconds`n`n"
[System.Diagnostics.Process]::GetCurrentProcess() | Invoke-FlashWindow
for($j =0; $j -lt 11; $j++){
    if($averagetime[$j] -ne "no data" -and $tru){
        Write-host $type[$j] "($($1st[$j])): " $averagetime[$j]
    }
}if(-not $tru){write-host "no results"}
$reportpath = "$TimeSearches\TimeSearches\$total$($department) $search.txt"
if($2013){$reportpath = "$TimeSearches\TimeSearches\2013\$total$($department) $search.txt"}
write-host "`n"
pause

if($tru){
    If((Test-Path -Path "$reportpath")){
        Remove-Item -Path "$reportpath" -Force
    }

    Add-Content -Path $reportpath -Value "$search"
    Add-Content -Path $reportpath -Value "Search arccuracy: $percentage%"
    Add-Content -Path $reportpath -Value "Total time: $($sw.elapsed)"
    Add-Content -Path $reportpath -Value ""
    for($j =0; $j -lt 11; $j++){
        $toast = Write-Output $type[$j] "($($1st[$j])):  " $averagetime[$j]
        Add-Content -Path $reportpath -Value "$toast"
    }

    Add-Content -Path $reportpath -Value ""
    Add-Content -Path $reportpath -Value ""
    Add-Content -Path $reportpath -Value "Matches: $($matchlist.GetUpperBound(0)+1)"
    for($i = 0; $i -le $matchlist.GetUpperBound(0); $i++){
        Add-Content -Path $reportpath -Value $matchlist[$i]
    }
}

if(test-path -path $reportpath){
    ii $reportpath
}


stop-transcript