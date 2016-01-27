cd "\\10.51.0.12\Tool Shop\recovery\Toolshop\Root\Script\"

$host.ui.RawUI.WindowTitle = get-date



$accuracy = (0.19)
$Path = "\\10.51.0.12\Tool Shop\recovery\Toolshop\Root\Script\Quote\TimeQuoteDataBase\full.csv"
(Get-Content $Path) | Set-Content $Path -Encoding UTF8
$data = import-csv $path
$total = $data.GetUpperBound(0)
$sw = [diagnostics.stopwatch]::StartNew()
foreach ($item in $data){
    $match1 = $item.Criterion8


    foreach($thing in $data){
        $LDpercent = .\Get-ld.ps1 $match2 $match1 -i
        if($LDpercent -lt $accuracy){
            break
        }
        $LDpercent = .\Get-ld.ps1 $thing $match1 -i
        if(!$LDpercent -lt $accuracy){
            foreach($carrot in $finallist){
                $LDpercent = .\Get-ld.ps1 $carrot $match1 -i
                if($LDpercent -lt $accuracy){
                    break
                }
            }
            For($number = 30; $number -ge 1; $number--){if($match1 -match "$number"){break}}
            if(!($number -lt 1 -or $thing -match $number -and $thing -notmatch "1$number" -and $thing -notmatch "2$number")){
            $finallist += @($match1)}
            $match2=$match1
            break
        }

    }
    $x++
    $peteyboy3000= [math]::round(($x/$total*100),4)
    clear-host
    write-host "%$peteyboy3000`nUnique Babies: $($finallist.length)"
}

$excel = New-Object -ComObject excel.application
$excel.Visible = $true
$workbook = $excel.Workbooks.Add()
$ws = $workbook.Worksheets.Item(1)
$ws.Name = 'Fulllist'
$ws.Activate() | Out-Null

for($row = 1; $row -le $finallist.getupperbound(0); $row++){
    $x = $row-1
    $ws.Cells.Item($row,1) = $finallist[$x]
}

$workbook.SaveAs("\\10.51.0.12\Tool Shop\recovery\Toolshop\Root\Script\fulllist.xlsx")
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel) | Out-Null
get-date
pause