$ToolShop = '\\10.51.0.12\DFSRoots\Tool Shop'



#Set Window to Front
$signature = @’ 
[DllImport("user32.dll")] 
public static extern bool SetWindowPos( 
    IntPtr hWnd, 
    IntPtr hWndInsertAfter, 
    int X, 
    int Y, 
    int cx, 
    int cy, 
    uint uFlags); 
‘@
$type = Add-Type -MemberDefinition $signature -Name SetWindowPosition -Namespace SetWindowPos -Using System.Text -PassThru
$handle = (Get-Process -id $Global:PID).MainWindowHandle 
$alwaysOnTop = New-Object -TypeName System.IntPtr -ArgumentList (-1) 
$type::SetWindowPos($handle, $alwaysOnTop, 0, 0, 0, 0, 0x0003) | Out-Null




if((get-content env:computername) -eq "FLO-8212TRA-PC"){$time = 8000}else{$time=15000}

Write-Host "Updating Guehring download...`nPlease wait window is NOT frozen, just working..."
get-process TM_schrank -ErrorAction SilentlyContinue | stop-process -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds ($time*0.05)
for($c = 1; $c -le 1; $c++){
    ii "C:\Program Files (x86)\TMSchrank\TM_Schrank.exe"
    Start-Sleep -Milliseconds ($time*0.3)
    $TMwindow = (Get-Process tm_schrank).MainWindowHandle

    add-type -AssemblyName microsoft.VisualBasic
    add-type -AssemblyName System.Windows.Forms
    start-sleep -Milliseconds 500

    [System.Windows.Forms.SendKeys]::SendWait("~~092214~{tab}{tab}~{tab}~{tab}{tab}") # Standard quick log in (as toolshop intern)
    try{get-process tm_schrank | out-null;Write-host "TM Running" -ForegroundColor DarkGreen}catch{$c = 0,"TM Failed"}
}
Write-Host "Loading confirmations"
[System.Windows.Forms.SendKeys]::SendWait("{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
Start-Sleep -Milliseconds ($time*0.3)
[System.Windows.Forms.SendKeys]::SendWait("~1~")
Write-Host "Saving Confirmations"
Start-Sleep -Milliseconds ($time*0.05)
[System.Windows.Forms.SendKeys]::SendWait("$ToolShop\05. Share\$('$ToolShopReportPkg')\$('$Downloads')\Metrics_Guehring\Confirmations.csv{tab}c+{tab}~~")
Start-Sleep -Milliseconds ($time*0.1)
Write-Host "Confirmations saved"
Start-sleep -s 1
Write-Host "Loading jobs"
[System.Windows.Forms.SendKeys]::SendWait("~{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}~")
Start-Sleep -Milliseconds ($time*0.2)
[System.Windows.Forms.SendKeys]::SendWait("+({tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab})~1~")
Start-Sleep -Milliseconds ($time*0.05)
Write-Host "Saving jobs"
[System.Windows.Forms.SendKeys]::SendWait("$ToolShop\05. Share\$('$ToolShopReportPkg')\$('$Downloads')\Orders_Guehring\Guehring.csv{tab}c+{tab}~~")
Write-Host "Jobs saved"
Start-sleep -s 1
Write-Host "Closing TM"
Start-Sleep -Milliseconds ($time*0.8)
get-process TM_schrank -ErrorAction SilentlyContinue | stop-process -ErrorAction SilentlyContinue