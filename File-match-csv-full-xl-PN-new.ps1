﻿if($PSVersionTable.PSVersion.major -lt 4){
    write-host "This script requires powershell version 4.0+"
    write-host "Update your computer to .net framework 4.0 to continue"
    read-host "Press enter to exit"
    exit
}

cd "\\10.51.0.12\Tool Shop\recovery\Toolshop\Root\Script\"
. "\\10.51.0.12\Tool Shop\recovery\Toolshop\Root\Script\FlashWindow.ps1"
."\\10.51.0.12\Tool Shop\recovery\Toolshop\Root\Script\Get-data.ps1"
$timeSearches = "\\10.51.0.12\Tool Shop\05. Share\35. Time Studies\"

$PN = "145003-0404" #Read-host "Enter Part Number"
$PNdept = "ts" #Read-host "Enter Department"
$accuracy = 85 #read-host "Enter accuracy (press enter for default value of 81%)"
$file=$pn

$data = import-csv "\\10.51.0.12\Tool Shop\recovery\Toolshop\Root\Script\Quote\TimeQuoteDataBase\full.csv"

foreach($item in $data){

    $PNmatch = $item.Criterion1

    if($PN -like $PNmatch){
        $Department = $item.criterion4
        if($Department -match $PNdept -or $PNdept -eq $null){
            $partname = $item.Criterion8
            if($partlist -notcontains $partname -and $item.Criterion14 -ne "AUFTRAG GESAMT"){
                    $Partlist += @($partname)
            }
        }
    }
}

$sw = [diagnostics.stopwatch]::StartNew()

$1st = @(0,0,0,0,0,0,0,0,0,0,0)
$2nd = 0
$averagetime = @("no data","no data","no data","no data","no data","no data","no data","no data","no data","no data","no data")
$lastmatch1 = $lastmatch2 = $lastmatch3 = "none"
$x=0
$department = $department.toupper()

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

$total = $data.GetUpperBound(0)
$host.ui.RawUI.WindowTitle = "Time Search $total"
write-host "no results"
foreach ($item in $data){
    
    $x++
    $match1 = $item.Criterion8
    $match2 = $item.Criterion14
    $match3 = $item.criterion4

    Write-Progress -activity “Searching for: '$PN' with accuracy of $percentage% $message" -id 43 -PercentComplete (($x/($total))*100) -CurrentOperation "Last 3 matches: $($lastmatch1), $($lastmatch2), $($lastmatch3)" -Status "$([math]::Round(($x/$total)*100))% Complete ($x/$($total+1)) / Elapsed Time: $([math]::Round($sw.elapsed.totalseconds)) Seconds / Total Matches: $($matchlist.getupperbound(0)+1)"
    
    if($match3 -match $department -or $department -eq $null){
    for($par = 0; $par -lt $search.getupperbound(0); $par++){
    $LDpercent = .\Get-ld.ps1 $search $match1 -i
    if($LDpercent -lt $accuracy){if($number -lt 1 -or $match1 -match $number -and $match1 -notmatch "1$number" -and $match1 -notmatch "2$number"){
        for($j =0; $j -lt 12; $j++){
            $LDpercent2 = .\Get-ld.ps1 $type[$j] $match2 -i
            if($LDpercent2 -lt 0.1){
                [float]$newav = (($item.Criterion20)/($item.Criterion10))
                if($newav -gt 0 -and $newav -lt 200){
                    $matchlist += @("$match3 $match1 --> $($type[$j]) $newav")
                    if($1st[$par][$j] -eq 0){
                        $averagetime[$par][$j] = $newav
                    }
                    else{
                        $averagetime[$par][$j]=[math]::round((($averagetime[$par][$j]*$1st[$par][$j]+$newav)/($1st[$par][$j]+1)),2)
                    }
                    $1st[$par][$j] = $1st[$par][$j]+1
                    $tru=$true
                    break
                }
            }}
        }}}
    }
}
Write-Progress -id 43 -complete -Activity "x"
$sw.stop

clear-host
write-host "Complete`n$search`nSearch arccuracy: $percentage%`n$($sw.elapsed.minutes) Minutes $($sw.elapsed.seconds) Seconds`n`n"
[System.Diagnostics.Process]::GetCurrentProcess() | Invoke-FlashWindow

if($tru){
    $reportpath = "\\10.51.0.12\Tool Shop\05. Share\$ToolShopReportPkg\TimeSearches\$total$($department) $search.txt"
    If((Test-Path -Path "$reportpath")){
        Remove-Item -Path "$reportpath" -Force
    }

    Add-Content -Path $reportpath -Value "$($search[$par])"
    Add-Content -Path $reportpath -Value "Search arccuracy: $percentage%"
    Add-Content -Path $reportpath -Value "Total time: $($sw.elapsed)"
    Add-Content -Path $reportpath -Value ""
    for($j =0; $j -lt 11; $j++){for($par = 0; $par -lt $search.getupperbound(0); $par++){
        $toast = Write-Output $type[$j] "($($1st[$par][$j])):  " $averagetime[$par][$j]
        Add-Content -Path $reportpath -Value "$toast"
    }}

    Add-Content -Path $reportpath -Value ""
    Add-Content -Path $reportpath -Value ""
    Add-Content -Path $reportpath -Value "Matches: $($matchlist.GetUpperBound(0)+1)"
    for($i = 0; $i -le $matchlist.GetUpperBound(0); $i++){
        Add-Content -Path $reportpath -Value $matchlist[$i]
    }
}


if($file){
    if(test-path "\\10.51.0.12\Tool Shop\05. Share\$ToolShopReportPkg\TimeSearches\$file.xlsx"){
        $excel = New-Object -ComObject excel.application
        $excel.Visible = $true
        $workbook = $excel.Workbooks.open("\\10.51.0.12\Tool Shop\05. Share\$ToolShopReportPkg\TimeSearches\$file.xlsx")
        $ws = $workbook.Worksheets.Item(1)
    }
    else{
        $excel = New-Object -ComObject excel.application
        $excel.Visible = $true
        $workbook = $excel.Workbooks.Add()
        $ws = $workbook.Worksheets.Item(1)
        $ws.Name = 'Machine Times'
        $ws.Activate() | Out-Null

        $ws.Cells.Item(1,1)= "Name"
        for($column = 2; $type[$($column-2)]; $column++){
            $name=$column-2
            $ws.Cells.Item(1,$column)= $type[$name]
        }
    }
    $range=$ws.UsedRange
    $range=($range.Rows.count)+1
    $ws.Cells.Item($range,1)= "$search"
    for($column = 2; $type[$($column-2)]; $column++){
        $name=$column-2
        if($averagetime[$name] -ne "no data"){$ws.Cells.Item($range,$column)= $averagetime[$name]}
    }
    if(test-path "\\10.51.0.12\Tool Shop\05. Share\$ToolShopReportPkg\TimeSearches\$file.xlsx"){$workbook.save()}
    else{$workbook.SaveAs("\\10.51.0.12\Tool Shop\05. Share\$ToolShopReportPkg\TimeSearches\$file.xlsx")}
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel) | Out-Null
}