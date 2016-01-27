if($PSVersionTable.PSVersion.major -lt 4){
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
$2013 = #read-host "2013 files? (y/n)"
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

$scriptblock = {Get-Data -search $search -department $PNdept -accuracy $accuracy}

foreach($search in $Partlist){
    start-job -InitializationScript ${function:get-data} -scriptblock {param($search, $PNdept, $accuracy) "Get-Data -search $search -department $PNdept -accuracy $accuracy"} -arg $search, $PNdept, $accuracy
}

$sw1 = [diagnostics.stopwatch]::StartNew()

for(!(get-job -State running)){
    clear-host
    write-host "Waiting for jobs to complete... `nTime: $([math]::round($sw1.Elapsed.TotalSeconds)) Seconds"
    get-job
    <#for($i = 2; $i -lt 29; $i++){
        (Get-Job -Id $i).ChildJobs[0].progress.percentcomplete
    }#>
    start-sleep -s 15
    #if(wait-job -id 2,4,6,8,10,12,14,16,18,20,22,24,26,28){break}
}

if(test-path "$TimeSearches\$file.xlsx"){ii "$TimeSearches\$file.xlsx"}
else{write-host "Error on File creation"}

pause