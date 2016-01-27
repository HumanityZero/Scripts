<#$network = "T:"

$file = $file = get-item "$network\05. Share\$('$ToolShopReportPkg')\Metrics\2015\ToolShopMetrics_*.xlsm" | sort LastWriteTime | Select -last 1
#>
$file = "C:\Users\bredestegej\Documents\Clean up.xlsm"

attrib +r $file
$file = $file
$sheetname = "PriorityOpen"
$before = @(Get-Process [e]xcel | %{$_.Id})
$xl = New-Object -ComObject Excel.Application
$ExcelId = Get-Process excel | %{$_.Id} | ?{$before -notcontains $_}
$wb = $xl.Workbooks.Open($file)
$ws = $wb.Worksheets.Item($sheetName)
$xl.Visible=$false

$z=$null
$search = "Closed"
try{ # Tried to search tracking sheet for the entered quote number. if not found returns error and is cause to tell user 'not found'
    $mainRng = $ws.usedRange
    $mainRng.Select() | out-null
    $objSearch = $mainrng.Find($Search)
    $First = $objSearch
    Do{
        $objSearch = $mainrng.FindNext($objSearch)
        $row = $objSearch.row
        $z += ,($ws.Cells.Item($row,"A").text, $ws.Cells.Item($row,"C").text, $ws.Cells.Item($row,"D").text, $("{0:D4}" -f [int]$ws.Cells.Item($row,"K").text), $("{0:D4}" -f [int]$ws.Cells.Item($row,"N").text))
    }
    While ($objSearch -ne $NULL -and $objSearch.AddressLocal() -ne $First.AddressLocal())
}catch{exit}

Stop-Process -Id $ExcelId -Force -ErrorAction SilentlyContinue



Add-Type @"
  using System;
  using System.Runtime.InteropServices;
  public class Tricks {
     [DllImport("user32.dll")]
     [return: MarshalAs(UnmanagedType.Bool)]
     public static extern bool SetForegroundWindow(IntPtr hWnd);
  }
"@
#[void] [Tricks]::SetForegroundWindow($PSwindow)
#[void] [Tricks]::SetForegroundWindow($TMwindow)

get-process TM_schrank -ErrorAction SilentlyContinue | stop-process -ErrorAction SilentlyContinue
for($5 = 0; $5 -lt $z.count; $5++){
$PSwindow = (Get-Process -PID $pid).MainWindowHandle
    for($c = 1; $c -le 1; $c++){
        ii "C:\Program Files (x86)\TMSchrank\TM_Schrank.exe"
        start-sleep -s 3
        $TMwindow = (Get-Process tm_schrank).MainWindowHandle
        [void] [Tricks]::SetForegroundWindow($TMwindow)

        add-type -AssemblyName microsoft.VisualBasic
        add-type -AssemblyName System.Windows.Forms
        start-sleep -Milliseconds 500

        [System.Windows.Forms.SendKeys]::SendWait("~~092214~{tab}{tab}~{tab}~+({tab}{tab}{tab}{tab})~{tab}{tab}{tab}{tab}{tab}{tab}") # Standard quick log in (as toolshop intern)
        Start-Sleep -m 500
        try{get-process tm_schrank | out-null}catch{$c = 0}
    }

    "$($z[$5][0])  $($z[$5][1])  $($z[$5][2])  $($z[$5][3])  $($z[$5][4])"


    [System.Windows.Forms.SendKeys]::SendWait("$($z[$5][2]){tab}{tab}{tab}{tab}{tab}{tab}$($z[$5][3]){tab}{tab}{tab}{tab}{tab}{tab}$($z[$5][4])~")
    [System.Windows.Forms.SendKeys]::SendWait("+({tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}){left}")
    [System.Windows.Forms.SendKeys]::SendWait("{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}Erledigt")
    [System.Windows.Forms.SendKeys]::SendWait("{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}~")
    start-sleep -m 1000
    [System.Windows.Forms.SendKeys]::SendWait("{right}~")
    start-sleep -m 1000
    [System.Windows.Forms.SendKeys]::SendWait("{tab}")

    
    get-process TM_schrank -ErrorAction SilentlyContinue | stop-process -ErrorAction SilentlyContinue
}