$ToolShop = "T:"
$ToolShop = '\\10.51.0.12\DFSRoots\Tool Shop\'

$host.UI.RawUI.ForegroundColor = "green"
$host.UI.RawUI.BackgroundColor = "black"

#keygen
for($i=1; $i -le 10; $i++){
    $safe=(26/36)*100
    $randomfyer=get-random -min 0 -max 99
    if($randomfyer -le $safe){
        $a=Get-Random -Count 1 -InputObject (65..90) | % -begin {$aa=$null} -process {$aa += [char]$_} -end {$aa}
    }
    else{$a=get-random -max 9 -min 0}
    if($i -eq 3 -or $i -eq 7){$a = "$a-"}
    #write-host "$a" -nonewline
    $serial="$serial$a"
}
try{
    $log = "$ToolShop\recovery\Toolshop\Root\Script\Logs\XS-$serial.log"
    Start-Transcript -path $log -append | out-null
}
catch{
    try{
        $log = "$ToolShop\05. Share\35. Time Studies\Logs\XS-$serial.log"
        Start-Transcript -path $log -append | out-null
    }
    catch{
        $log = "C:\Windows\Temp\XS-$serial.log"
        Start-Transcript -path $log -append | out-null
    }
}

if($PSVersionTable.PSVersion.major -lt 4){
    write-host "This script requires powershell version 4.0+ (you have v$($PSVersionTable.PSVersion.major))"
    write-host "Update your computer to .net framework 4.0 to continue"
    read-host "Press enter to exit"
    exit
}
add-type -an system.windows.forms

$ErrorActionPreference = "continue"#"inquire"#"silentlycontinue"

cd "T:\recovery\Toolshop\Root\Script\"
. "T:\recovery\Toolshop\Root\Script\FlashWindow.ps1"
."T:\recovery\Toolshop\Root\Script\Get-data.ps1"
$timeSearches = "T:\05. Share\35. Time Studies\XPPS\"


$PN = Read-host "Enter XPPS Number"
$PNdept = Read-host "Enter Department"
$accuracy = read-host "Enter accuracy (press enter for default value of 90%)"

if($accuracy -gt 1){
    $accuracy = (100-$accuracy)/100
}
elseif($accuracy -lt 1 -and $accuracy -gt 0){
    $accuracy = 1-$accuracy
}
else{
    $accuracy = (0.1)
}
$percentage = $((1-$accuracy)*100)


$Path = "T:\recovery\Toolshop\Root\Script\Quote\TimeQuoteDataBase\full.csv"
(Get-Content $Path) | Set-Content $Path -Encoding UTF8
$data = import-csv $path

foreach($item in $data){

    $PNmatch = $item.Criterion1

    if($PNmatch -match $PN){
        $Department = $item.criterion4
        if($Department -match $PNdept -or $PNdept -eq ''){
            $partname = $item.Criterion8
            if($partlist -notcontains $partname -and $item.Criterion14 -ne "AUFTRAG GESAMT"){
                    $Partlist += @($partname)
                    $partname
            }
        }
    }
}
try{
    $parcount = $partlist.getupperbound(0)+1
}
catch{
    write-host "`nNo Gühring Data" -foregroundcolor "red"
    pause
    exit
}
$sw = [diagnostics.stopwatch]::StartNew()
write-host "Number of parts: $parcount"
start-sleep -s 4

For($number = 30; $number -ge 1; $number--){
    if($search -match "$number"){break}
}

$type = @("Multiaxial Milling","CNC Turning","Surface Grinding","Circular grinding","External Heat Treatment","Manual Workstations","Material","Standard Turning","Wire Cut_EDM","Sawing","Standard Milling")

$1st = ,(0,0,0,0,0,0,0,0,0,0,0)
$averagetime = ,('','','','','','','','','','','')
for($x=0; $x -le $parcount; $x++){
    $1st += ,(0,0,0,0,0,0,0,0,0,0,0)
    $averagetime += ,('','','','','','','','','','','')
}
$PNdept = $PNdept.toupper()
if($PNdept){$message = "in $PNdept"}

$x = 0
$total = $data.GetUpperBound(0)
$host.ui.RawUI.WindowTitle = "XPPS Search  [$($total)CSVXPPSv3.2]  $PN  XS-$Serial"
foreach ($item in $data){
    
    $x++
    $match = $item.Criterion8
    $position = $item.Criterion14
    $dept = $item.criterion4
    $actime = $item.Criterion20
    $Quantity = $item.Criterion10
    if($ding -gt 0){$status = "$ding Match found!"; $ding=(($ding*0.99)-1)}else{$status = "Searching..."}

    Write-Progress -activity “Searching for: '$PN' with accuracy of $percentage% $message" -id 43 -currentoperation $status -PercentComplete (($x/($total))*100) -Status "$([math]::Round(($x/$total)*100))% Complete ($x/$($total+1)) / Elapsed Time: $([math]::Round($sw.elapsed.totalseconds)) Seconds / Total Matches: $count"
    if($dept -match $PNdept -or $PNdept -eq $null){
        try{[float]$newav = (($actime)/($Quantity))}catch{$newav = $null}
        if($newav -gt 0 -and $newav -lt 200){
            for($par = 0; $par -lt $parcount; $par++){
                $LDpercent = .\Get-ld.ps1 $partlist[$par] $match -i
                if($LDpercent -lt $accuracy){
                    For($number = 30; $number -ge 1; $number--){
                        if($partlist[$par] -match "$number"){break}
                    }
                    if($number -lt 1 -or $match -match $number -and $match -notmatch "1$number" -and $match -notmatch "2$number"){
                        $ding = 50
                        for($j =0; $j -lt 11; $j++){
                            if($position -match $type[$j]){
                                if($1st[$par][$j] -eq 0){
                                    $averagetime[$par][$j] = [math]::round($newav,2)
                                }
                                else{
                                    $averagetime[$par][$j]=[math]::round((($averagetime[$par][$j]*$1st[$par][$j]+$newav)/($1st[$par][$j]+1)),2)
                                }
                                $1st[$par][$j] = $1st[$par][$j]+1
                                $count++
                                break
                            }
                        }
                    }
                }
            }
        }
    }
}
Write-Progress -id 43 -complete -Activity "x"
#$sw.stop
[System.Diagnostics.Process]::GetCurrentProcess() | Invoke-FlashWindow

$reportpath = "$timeSearches\$PN.xlsx"
If((Test-Path -Path "$reportpath")){
    Remove-Item -Path "$reportpath" -Force
}
$excel = New-Object -ComObject excel.application
$excel.Visible = $false
$workbook = $excel.Workbooks.Add()
$ws = $workbook.Worksheets.Item(1)
$ws.Name = 'Machine Times'
$ws.Activate() | Out-Null
$ws.Cells.Item(1,1)= "Name"
for($column = 2; $type[$($column-2)]; $column++){
    $name=$column-2
    $ws.Cells.Item(1,$column)= $type[$name]
}

for($par = 0; $par -lt $parcount; $par++){
    $range=$ws.UsedRange
    $range=($range.Rows.count)+1
    $ws.Cells.Item($range,1)= "$($partlist[$par])"
    for($column = 2; $type[$($column-2)]; $column++){
        $name=$column-2
        if($($1st[$par][$name]) -ne 0){$ws.Cells.Item($range,$column) = "$($averagetime[$par][$name]) ($($1st[$par][$name]))"}
    }
}

$ws.columns.Item(1).EntireColumn.AutoFit()
$xlContinuous = 1
$xlInsideHorizontal = 12 
$xlInsideVertical = 11 
$xlNone = -4142 
$xlThin = 1
$selection = $ws.range("B1:M1") 
$selection.select()
$selection.Orientation = 45
$selection = $ws.range("A1:M$($parcount+1)") 
$selection.select()
$selection.IndentLevel = 0 
$selection.ShrinkToFit = $false
$selection.Borders.Item($xlInsideHorizontal).LineStyle = $xlContinuous 
$selection.Borders.Item($xlInsideVertical).LineStyle = $xlContinuous
$selection.Borders.Item($xlInsideHorizontal).Weight = $xlThin 
$selection.Borders.Item($xlInsideVertical).Weight = $xlThin 

$workbook.SaveAs("$TimeSearches\$PN.xlsx")

$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel) | Out-Null

ii "$TimeSearches\$PN.xlsx"

<#
$xlAutomatic = -4105 
$xlBottom = -4107 
$xlCenter = -4108 
$xlContext = -5002 
$xlContinuous = 1 
$xlDiagonalDown = 5 
$xlDiagonalUp = 6 
$xlEdgeBottom = 9 
$xlEdgeLeft = 7 
$xlEdgeRight = 10 
$xlEdgeTop = 8 
$xlInsideHorizontal = 12 
$xlInsideVertical = 11 
$xlNone = -4142 
$xlThin = 2
$xl = new-object -com excel.application 
$xl.visible=$true 
$wb = $xl.workbooks.open("C:ScriptsBook1.xlsx") 
$ws = $wb.worksheets | where {$_.name -eq "sheet1"} 
$selection = $ws.range("A1:F29") 
$selection.select()
$selection.HorizontalAlignment = $xlCenter 
$selection.VerticalAlignment = $xlBottom 
$selection.WrapText = $false 
$selection.Orientation = 0 
$selection.AddIndent = $false 
$selection.IndentLevel = 0 
$selection.ShrinkToFit = $false 
$selection.ReadingOrder = $xlContext 
$selection.MergeCells = $false 
$selection.Borders.Item($xlInsideHorizontal).Weight = $xlThin 
$selection.Borders.Item($xlInsideHorizontal).LineStyle = $xlContinuous 
$selection.Borders.Item($xlInsideHorizontal).ColorIndex = $xlAutomatic
#>


[System.Windows.Forms.Clipboard]::GetText()
clear-host
stop-transcript