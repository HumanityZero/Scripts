Add-Type @"
  using System;
  using System.Runtime.InteropServices;
  public class Tricks {
     [DllImport("user32.dll")]
     [return: MarshalAs(UnmanagedType.Bool)]
     public static extern bool SetForegroundWindow(IntPtr hWnd);
  }
"@

Write-host "Follow these instructions very carefully:
1. Open geuhring
2. Load the job you wish to edit
3. Select the first item that you want to edit`n
Press enter to skip an item:"

$count = 500 #Read-Host "Approximate number of lines to edit"
#if($count -gt 20){$count = 20}
<#$jobNo = Read-host "Job Number"
$Xvas = Read-host "Xvas Number"
$order = Read-host "Order Number"
$Customer = Read-host "Customer"
$OrderValue = Read-host "Order Value"
$OrderDate = Read-host "Order Date"
$EndDate = Read-host "End Date"
$JobPosition = Read-host "Job Position"
$Component = Read-host "Component"
$DrawingNo = Read-host "Drawing Number"
$Quantity = Read-host "Quantity"
$Status = Read-host "Status"#>
$Timeest = 1.01#Read-Host "Estimated Time"
#$IdentNo = Read-host "Ident Number"
#$WzNo = Read-host "WzNo"
$speed = 1000#Read-Host "Run speed (default 500)"
#if($speed -eq ""){$speed=500}
Write-host "Ready to start" -nonewline
read-host

add-type -AssemblyName microsoft.VisualBasic
add-type -AssemblyName System.Windows.Forms


$PSwindow = (Get-Process -PID $pid).MainWindowHandle
$TMwindow = (Get-Process tm_schrank).MainWindowHandle
[void] [Tricks]::SetForegroundWindow($TMwindow)

start-sleep -Milliseconds 500

#[Microsoft.VisualBasic.Interaction]::AppActivate("TM_Schrank")
#start-sleep -s 2
for($t = 0; $t -lt $count; $t++){
    #$stamp = "[AutoJobCard] $(get-content env:username) $(get-date)"
    [System.Windows.Forms.SendKeys]::SendWait("{tab}$JobNo")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$Xvas")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$order")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$Customer")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$OrderValue")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$OrderDate")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$JobPosition")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$Component")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$DrawingNo")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$Quantity")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$EndDate")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$Status")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$Operation")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$Activity")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$ActivityDescription")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$Machine")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")#costperhour
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")#costcenter
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$Timeest")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$Timeact")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")#costest
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")#costact
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$IdentNo")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$WzNo")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")#creator
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}$stamp")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")#start
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")#newsearch
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")#showconfirmation
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")#totalcost
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")#copyjob
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")#copyposition
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")#copy
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")#save
    start-sleep -m 100 #$speed
    [System.Windows.Forms.SendKeys]::SendWait("~~~{ESC}{ESC}")
    #start-sleep -m $speed
    #for($pddy -eq 0; $pddy -le 34; $pddy++){[System.Windows.Forms.SendKeys]::SendWait("+{TAB}")}
    #[System.Windows.Forms.SendKeys]::SendWait("{down}")
    start-sleep -m $speed
    "$t > $count"
}

[void] [Tricks]::SetForegroundWindow($PSwindow)
read-host "Complete, press enter to exit"