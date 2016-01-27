$random = Get-random
Start-Transcript -path "\\10.51.0.12\Tool Shop\recovery\Toolshop\Root\Script\Logs\TM$random.log" -append | out-null
cd A:\

#version 1.6 for Quote sheet version 4

$QN = read-host "Enter Quote Number"
$z = $null

Write-host "Loading tracking sheet info"
##############Intial Quote Loadup##########################################
attrib +r '\\10.51.0.12\Tool Shop\05. Share\$ToolShopReportPkg\ToolShopOrders.xlsm'
$file = '\\10.51.0.12\Tool Shop\05. Share\$ToolShopReportPkg\ToolShopOrders.xlsm'
$sheetname = "active"
$before = @(Get-Process [e]xcel | %{$_.Id})
$xl = New-Object -ComObject Excel.Application
$ExcelId = Get-Process excel | %{$_.Id} | ?{$before -notcontains $_}
$wb = $xl.Workbooks.Open($file)
$ws = $wb.Worksheets.Item($sheetName)
$xl.Visible=$false
$mainRng = $ws.usedRange
$mainRng.Select() | out-null
$objSearch = $mainrng.Find($QN)
$First = $objSearch
Do
{
    $objSearch = $mainrng.FindNext($objSearch)
    $row = $objSearch.row
    $z += ,($ws.Cells.Item($row,1).text, $ws.Cells.Item($row,3).text, $ws.Cells.Item($row,4).text, $ws.Cells.Item($row,6).text, $ws.Cells.Item($row,10).text, $ws.Cells.Item($row,"BN").text, $ws.Cells.Item($row,"BO").text, $ws.Cells.Item($row,"BR").text, $ws.Cells.Item($row,"BS").text, $ws.Cells.Item($row,"BT").text)
}
While ($objSearch -ne $NULL -and $objSearch.AddressLocal() -ne $First.AddressLocal())
Stop-Process -Id $ExcelId -Force -ErrorAction SilentlyContinue

#############Launch and Load TM ############################################
write-host "Launch and Load TM"
#[void] [Tricks]::SetForegroundWindow($PSwindow)
#[void] [Tricks]::SetForegroundWindow($TMwindow)

get-process TM_schrank -ErrorAction SilentlyContinue | stop-process -ErrorAction SilentlyContinue
$PSwindow = (Get-Process -PID $pid).MainWindowHandle
Add-Type @"
  using System;
  using System.Runtime.InteropServices;
  public class Tricks {
     [DllImport("user32.dll")]
     [return: MarshalAs(UnmanagedType.Bool)]
     public static extern bool SetForegroundWindow(IntPtr hWnd);
  }
"@

ii "C:\Program Files (x86)\TMSchrank\TM_Schrank.exe"
start-sleep -s 5
$TMwindow = (Get-Process tm_schrank).MainWindowHandle
[void] [Tricks]::SetForegroundWindow($TMwindow)

add-type -AssemblyName microsoft.VisualBasic
add-type -AssemblyName System.Windows.Forms
start-sleep -Milliseconds 500

[System.Windows.Forms.SendKeys]::SendWait("~~092214~{tab}{tab}~{tab}~+({tab}{tab}{tab}{tab})~{tab}{tab}{tab}{tab}{tab}{tab}")
start-sleep -m 200


######Load Quote info#######################################################
[void] [Tricks]::SetForegroundWindow($PSwindow)
write-host "Loading Quote info"
$QN = $QN.Trim("Q")
$file = get-item "\\10.51.0.12\Tool Shop\05. Share\11. Tool Shop Quotes\01_New-parts\03. Quotes 2015\$QN*\$QN*.xlsx"
$sheetname = "part_1"
$file
$before = @(Get-Process [e]xcel | %{$_.Id})
$xl = New-Object -ComObject Excel.Application
$ExcelId = Get-Process excel | %{$_.Id} | ?{$before -notcontains $_}
$wb = $xl.Workbooks.Open($file)
$ws = $wb.Worksheets.Item($sheetName)
$xl.Visible=$false
$Jobname = $wb.Worksheets.Item("Sheet1").Cells.Item(18,"B").text
$Partname = ($ws.Cells.Item(6,4).text).split(" ")
$Quantity = $ws.Cells.Item(7,2).text
$Material = $ws.Cells.Item(8,2).text
$drawing = $part = $null
foreach($item in $partname){

    if($item -match '\d+'){
        $drawing = "$drawing $item"
    }
    else{
        $part = "$part $item"
    }

}
$Part = $part.TrimStart(" ")
$drawing = $drawing.TrimStart(" ")

#positions
$positions = $null
for($X=35; $X -le 42; $X++){
    $positions += ,($ws.Cells.Item($X,1).text, $ws.Cells.Item($X,3).text)
}

Stop-Process -Id $ExcelId -Force -ErrorAction SilentlyContinue


#########Job Loader ########################################################
[void] [Tricks]::SetForegroundWindow($TMwindow)
#Load up fist job:

$poscnt = 3
$operation = 10

#Auftrag
[System.Windows.Forms.SendKeys]::SendWait("joe$(<#$z[0][2]#>){tab}$($z[0][1]){tab}$($z[0][3]){tab}$($z[0][0]){tab}$($z[0][4]){tab}$($z[0][5]){tab}{tab}$Jobname{tab}{tab}$Quantity{tab}$($z[0][6]){tab}Offen{tab}{tab}{down}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}Material: $material{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
start-sleep -s 1
[System.Windows.Forms.SendKeys]::SendWait("~~")
start-sleep -m 50
for($i = 0; $i -lt 33; $i++){[System.Windows.Forms.SendKeys]::SendWait("+{TAB}")}

#position
[System.Windows.Forms.SendKeys]::SendWait("joe$(<#$z[0][2]#>){tab}$($z[0][1]){tab}$($z[0][3]){tab}$($z[0][0]){tab}{tab}$($z[0][5]){tab}0100{tab}$Part{tab}$drawing{tab}$Quantity{tab}$($z[0][6]){tab}Offen{tab}{tab}{down}{down}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
start-sleep -s 1
[System.Windows.Forms.SendKeys]::SendWait("~~")
start-sleep -m 50
for($i = 0; $i -lt 33; $i++){[System.Windows.Forms.SendKeys]::SendWait("+{TAB}")}

#material
[System.Windows.Forms.SendKeys]::SendWait("joe$(<#$z[0][2]#>){tab}$($z[0][1]){tab}$($z[0][3]){tab}$($z[0][0]){tab}{tab}$($z[0][5]){tab}0100{tab}$Part{tab}$drawing{tab}$Quantity{tab}$($z[0][6]){tab}Offen{tab}0005{tab}{down}{down}{tab}{tab}{down}{down}{down}{down}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
start-sleep -s 1
[System.Windows.Forms.SendKeys]::SendWait("~~")
start-sleep -m 50
for($i = 0; $i -lt 33; $i++){[System.Windows.Forms.SendKeys]::SendWait("+{TAB}")}
start-sleep -s 5


for($XX = 0; $XX -le 7; $XX++){
    $opname = $positions[$XX][0]
    $optime = $positions[$XX][1]
    if($optime -gt 0){
        [System.Windows.Forms.SendKeys]::SendWait("joe$(<#$z[0][2]#>){tab}$($z[0][1]){tab}$($z[0][3]){tab}$($z[0][0]){tab}{tab}$($z[0][5]){tab}0100{tab}$Part{tab}$drawing{tab}$Quantity{tab}$($z[0][6]){tab}Offen{tab}00$operation{tab}$opname{tab}{tab}Raw bandsaw{tab}{tab}{tab}$optime{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
        start-sleep -s 2
        [System.Windows.Forms.SendKeys]::SendWait("~")
        start-sleep -s 2
        [System.Windows.Forms.SendKeys]::SendWait("~")
        for($i = 0; $i -lt 33; $i++){[System.Windows.Forms.SendKeys]::SendWait("+{TAB}")}
        $operation=$operation+10
    }
}


#end
[System.Windows.Forms.SendKeys]::SendWait("joe$(<#$z[0][2]#>)~")


[void] [Tricks]::SetForegroundWindow($PSwindow)
Start-Sleep -s 15
stop-transcript | out-null