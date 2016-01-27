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

$jobNo = Read-host "Job Number"
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
$Status = Read-host "Status"
$IdentNo = Read-host "Ident Number"
$WzNo = Read-host "WzNo"
$count = Read-Host "Approximate number of jobs to edit"
Write-host "Ready to start" -nonewline
read-host

add-type -AssemblyName microsoft.VisualBasic
add-type -AssemblyName System.Windows.Forms



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
    [System.Windows.Forms.SendKeys]::SendWait("~")
    start-sleep -m 5
    [System.Windows.Forms.SendKeys]::SendWait("{down}")
}

read-host "Complete, press enter to exit"