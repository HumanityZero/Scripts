#cd "\\10.51.0.12\Tool Shop\recovery\Toolshop\Root\Script\Quote\"
cd "\\10.51.0.12\Tool Shop\recovery\Toolshop\Root\Script\"
. "\\10.51.0.12\Tool Shop\recovery\Toolshop\Root\Script\FlashWindow.ps1"
$search = Read-host "Enter Criteria"

write-host "Running`n`n`n`n`n`n`n`n`n"

$1st = @(0,0,0,0,0,0,0,0,0,0,0)
$averagetime = @("no data","no data","no data","no data","no data","no data","no data","no data","no data","no data","no data")

$type = @("Multiaxial_Milling","CNC_Turning","Surface_Grinding","Circular_grinding","External_Heat_Treatment","Manual_Workstations","Wire_Cut_EDM","Material","Sawing","Standard_Milling","Standard_Turning")

$ErrorActionPreference = "inquire"

$data = import-csv "\\10.51.0.12\Tool Shop\recovery\Toolshop\Root\Script\Quote\TimeQuoteDataBase\full.csv"
$total = $data.GetUpperBound(0)
$x=0
$host.ui.RawUI.WindowTitle = "Time Search $total"
write-host "no results"
foreach ($item in $data){
    
    $x++
    $match1 = $item.Criterion8
    $match2 = $item.Criterion14

    Write-Progress -activity “$Search" -id 43 -PercentComplete (($x/($total))*100) -CurrentOperation "($x/$total) $match1"
    
    
    $LDpercent = .\Get-ld.ps1 $search $match1 -i
    if($LDpercent -lt 0.2){
        for($j =0; $j -lt 12; $j++){
            $LDpercent2 = .\Get-ld.ps1 $type[$j] $match2 -i
            if($LDpercent2 -lt 0.1){
                $matchlist += @("$match1 --> $($type[$j])")
                [float]$newav = (($item.Criterion20)/($item.Criterion10))
                if($newav -gt 0 -and $newav -lt 200){
                    if($1st[$j] -eq 0){
                        $averagetime[$j] = $newav
                    }
                    else{
                        $averagetime[$j]=[math]::round((($averagetime[$j]*$1st[$j]+$newav)/($1st[$j]+1)),2)
                    }
                    $1st[$j] = $1st[$j]+1
                    clear-host
                    write-host "Running`n`n`n`n`n`n`n`n`n`n"
                    for($jp =0; $jp -lt 11; $jp++){
                        if($averagetime[$jp] -ne "no data"){
                            Write-host $type[$jp] "($($1st[$jp])): " $averagetime[$jp]
                        }
                    }
                    $tru=$true
                }
            }
        }
    }
}

clear-host
write-host "$search`nComplete`n`n"
[System.Diagnostics.Process]::GetCurrentProcess() | Invoke-FlashWindow
for($j =0; $j -lt 11; $j++){
    if($averagetime[$j] -ne "no data" -and $tru){
        Write-host $type[$j] "($($1st[$j])): " $averagetime[$j]
    }
}if(-not $tru){write-host "no results"}
$ToolShopReportPkg = '$ToolShopReportPkg'
write-host "`n"
pause
if($tru){
    If((Test-Path -Path "\\10.51.0.12\Tool Shop\05. Share\$ToolShopReportPkg\TimeSearches\$total $search.txt")){
        Remove-Item -Path "\\10.51.0.12\Tool Shop\05. Share\$ToolShopReportPkg\TimeSearches\$total $search.txt" -Force
    }

    for($j =0; $j -lt 11; $j++){
        $toast = Write-Output $type[$j] "($($1st[$j])):  " $averagetime[$j]
        Add-Content -Path "\\10.51.0.12\Tool Shop\05. Share\$ToolShopReportPkg\TimeSearches\$total $search.txt" -Value "$toast"
    }

    Add-Content -Path "\\10.51.0.12\Tool Shop\05. Share\$ToolShopReportPkg\TimeSearches\$total $search.txt" -Value ""
    Add-Content -Path "\\10.51.0.12\Tool Shop\05. Share\$ToolShopReportPkg\TimeSearches\$total $search.txt" -Value ""
    for($i = 0; $i -lt $matchlist.GetUpperBound(0); $i++){
        Add-Content -Path "\\10.51.0.12\Tool Shop\05. Share\$ToolShopReportPkg\TimeSearches\$total $search.txt" -Value $matchlist[$i]
    }
}

if(test-path -path "\\10.51.0.12\Tool Shop\05. Share\$ToolShopReportPkg\TimeSearches\$total $search.txt"){
    ii "\\10.51.0.12\Tool Shop\05. Share\$ToolShopReportPkg\TimeSearches\$total $search.txt"
}