#version 3.0 for Quote sheet version 5.1

$ToolShop = "T:"
$ToolShop = '\\10.51.0.12\DFSRoots\Tool Shop\'

$host.UI.RawUI.ForegroundColor = "green"
$host.UI.RawUI.BackgroundColor = "black"

$ErrorActionPreference = "Continue"#Stop"

#keygen - Generates random serial number for log file
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
    $log = "$ToolShop\recovery\Toolshop\Root\Script\Logs\TM-$serial.log"
    Start-Transcript -path $log -append | out-null
}
catch{
    try{
        $log = "$ToolShop\05. Share\35. Time Studies\Logs\TM-$serial.log"
        Start-Transcript -path $log -append | out-null
    }
    catch{
        $log = "C:\Windows\Temp\TM-$serial.log"
        Start-Transcript -path $log -append | out-null
    }
}

if($PSVersionTable.PSVersion.major -lt 4){
    write-host "This script requires powershell version 4.0+ (you have v$($PSVersionTable.PSVersion.major))"
    write-host "Update your computer to .net framework 4.0 to continue"
    read-host "Press enter to exit"
    exit
}

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
#[void] [Tricks]::SetForegroundWindow($PSwindow)
#[void] [Tricks]::SetForegroundWindow($TMwindow)

#disclaimer
if((get-content env:username) -ne "bredestegej"){
    Clear-Host
    Write-Host "Session ID: TM-$serial" -ForegroundColor DarkGreen
    Write-Host "`nNOTE: It is very important that you do not click with your mouse.`nIf the program starts to run haywire hit ALT+F4 several times. `nTM will also be restarted, ensure it is closed or everything is saved." -ForegroundColor Yellow
    Write-Host "`nJob Card Creator version 3.0`n" -ForegroundColor Cyan
    Pause
}

function TM-return { # Saves position and returns to job No. space
    param([float]$speed)
    start-sleep -m $speed
    [System.Windows.Forms.SendKeys]::SendWait("~")
    start-sleep -m ($speed/5)
    [System.Windows.Forms.SendKeys]::SendWait("{ESC}")
    start-sleep -m ($speed/5)
    [System.Windows.Forms.SendKeys]::SendWait("{ESC}")
    start-sleep -m ($speed/10)
    [System.Windows.Forms.SendKeys]::SendWait("{ESC}")
    for($i = 0; $i -lt 33; $i++){[System.Windows.Forms.SendKeys]::SendWait("+{TAB}")}
}

function TM-respond {
    param($ExcelId, $PSwindow)
    for($r=0; $r -lt 5; $r++){
        try{
            if((get-process TM_Schrank -ea stop).responding){
                return
            }
        }
        catch{
            Write-Host "[Diagnostics] TM NOT FOUND" -ForegroundColor Black
            break
        }
        start-sleep -s 2
    }
    Add-Type @"
  using System;
  using System.Runtime.InteropServices;
  public class Tricks {
     [DllImport("user32.dll")]
     [return: MarshalAs(UnmanagedType.Bool)]
     public static extern bool SetForegroundWindow(IntPtr hWnd);
  }
"@
    [void] [Tricks]::SetForegroundWindow($PSwindow)
    Write-host "TM is not responding" -ForegroundColor Yellow
    Read-host "Press Enter to Exit"
    Stop-Process -Id $ExcelId -Force -ErrorAction SilentlyContinue
    Stop-Process -Id $ExcelId2 -Force -ErrorAction SilentlyContinue
    Stop-Transcript
    Add-Content -Path $log -Value '[Diagnostics] $z Part info'
    Add-Content -Path $log -Value $z
    Add-Content -Path $log -Value '[Diagnostics] $bookmark'
    Add-Content -Path $log -Value $bookmark
    Exit
}

Function CheckBox{
#==============================================================================================
# XAML Code - Imported from Visual Studio Express WPF Application
#==============================================================================================
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="New Part" Height="350" Width="447.911">
    <Grid Margin="0,0,-0.2,0.2">
        <Border BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="72" Margin="48,231,0,0" VerticalAlignment="Top" Width="240"/>
        
        <CheckBox Name="MAT" Content="Material" HorizontalAlignment="Left" Margin="10,98,0,0" VerticalAlignment="Top" IsChecked="True"/>
        <CheckBox Name="SAW" Content="Sawing" HorizontalAlignment="Left" Margin="75,98,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="CAM" Content="CAM" HorizontalAlignment="Left" Margin="10,147,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="ST" Content="Standard Turning" HorizontalAlignment="Left" Margin="75,119,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="CN" Content="CNC Turning" HorizontalAlignment="Left" Margin="75,140,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="SM" Content="Standard Milling" HorizontalAlignment="Left" Margin="75,161,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="MAM" Content="Multiaxial Milling" HorizontalAlignment="Left" Margin="75,182,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="CG" Content="Circular grinding" HorizontalAlignment="Left" Margin="75,203,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="SG" Content="Surface Grinding" HorizontalAlignment="Left" Margin="189,98,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="EDM" Content="Wire Cut EDM" HorizontalAlignment="Left" Margin="189,119,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="HM" Content="Hard Milling" HorizontalAlignment="Left" Margin="189,140,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="HT" Content="Hard Turning" HorizontalAlignment="Left" Margin="189,161,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="HMF" Content="Hard Milling Finish" HorizontalAlignment="Left" Margin="189,182,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="HTF" Content="Hard Turning Finsh" HorizontalAlignment="Left" Margin="189,203,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="MW" Content="Manual Workstations" HorizontalAlignment="Left" Margin="299,98,0,0" VerticalAlignment="Top" IsChecked="True"/>
        
        <Button Name="btnExit" Content="Ok" HorizontalAlignment="Left" Margin="353,267,0,0" VerticalAlignment="Top" Width="75"/>
        
        <CheckBox Name="NIT" Content="Nitriding" HorizontalAlignment="Left" Margin="168,259,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="COT" Content="Coating" HorizontalAlignment="Left" Margin="168,238,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="Heatreat" Content="Heat Treatment" HorizontalAlignment="Left" Margin="57,238,0,0" VerticalAlignment="Top" Width="103"/>
        <CheckBox Name="BOF" Content="Black Oxid Finish" HorizontalAlignment="Left" Margin="168,280,0,0" VerticalAlignment="Top"/>
        
        <TextBox Name="PN" HorizontalAlignment="Left" Height="38" Margin="10,10,0,0" TextWrapping="Wrap" Text="Part Name" VerticalAlignment="Top" Width="341" FontSize="18.667"/>
        <TextBox Name="Quantity" HorizontalAlignment="Left" Height="38" Margin="356,10,0,0" TextWrapping="Wrap" Text="Quantity" VerticalAlignment="Top" Width="74" FontSize="16"/>
        <TextBox Name="DN" HorizontalAlignment="Left" Height="30" Margin="10,53,0,0" TextWrapping="Wrap" Text="Drawing Number" VerticalAlignment="Top" Width="420" AutoWordSelection="True" FontSize="16"/>
        <TextBox Name="Material" HorizontalAlignment="Left" Height="23" Margin="10,119,0,0" TextWrapping="Wrap" Text="Material" VerticalAlignment="Top" Width="60"/>
        <TextBox Name="HRD" HorizontalAlignment="Left" Height="23" Margin="57,259,0,0" TextWrapping="Wrap" Text="Hardness" VerticalAlignment="Top" Width="106" FontSize="13.333"/>
        <TextBox Name="XPPS" HorizontalAlignment="Left" Height="28" Margin="338,234,0,0" TextWrapping="Wrap" Text="$XPPS1" VerticalAlignment="Top" Width="90" FontSize="13.333"/>
        <CheckBox Name="Print" Content="Print?" HorizontalAlignment="Left" Margin="376,294,0,0" VerticalAlignment="Top" Width="52" IsChecked="True"/>
        <CheckBox Name="More" Content="More" HorizontalAlignment="Left" Margin="322,294,0,0" VerticalAlignment="Top" Width="49" IsChecked="True"/>

    </Grid>
</Window>
"@
#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; exit}

#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================

$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)}


#===========================================================================
# Shows the form
#===========================================================================
$btnExit.Add_Click({
    if($XPPS.text.Length -ge 11 -and $PN.text -ne "Part Name" -and $Quantity.text -ne "Quantity" -and $DN.text -ne "Drawing Number"){
        $form.Close()
    }
})
$Form.ShowDialog() | out-null
$checkbox = @($PN.text,$DN.text,$XPPS.text,$Position.text,$Material.text,$Quantity.text,$Print.IsChecked,$Mat.IsChecked,$SAW.IsChecked,$CAM.IsChecked,$ST.IsChecked,$CN.IsChecked,$SM.IsChecked,$MAM.IsChecked,$CG.IsChecked,$SG.IsChecked,$EDM.IsChecked,$HM.IsChecked,$HT.IsChecked,$HMF.IsChecked,$HTF.IsChecked,$MW.IsChecked,$Heatreat.IsChecked,$HRD.text,$NIT.IsChecked,$COT.IsChecked,$BOF.IsChecked,$More.IsChecked)

return $checkbox
}

#if((get-content env:username) -ne "bredestegej"){$ErrorActionPreference= "Silently Continue"}
$checkbox = $null
$XPPS1 = 1450
Clear-Host
$QN = read-host "Enter Quote Number"
if($QN -eq "QHC"){
    Write-host "WARNING: This function is in beta and may not work right." -ForegroundColor Yellow
    $PO = Read-host "Enter Customer PO"
    $Search = $PO
}
elseif($QN -eq "Custom"){
    for($i=0;$i -lt 3000;$i++){
        $checkbox += ,(CheckBox $XPPS1)
        $search += ,($checkbox[$i][2])
        $XPPS1 = $checkbox[$i][2]
        if(!$checkbox[$i][27]){break}
    }
}
else{$Search = $QN}
$safemode = Read-Host "Safe mode? (y/n)" # safe mode names jobs "Joe" and pauses for 3 seconds before saving each position
$host.ui.RawUI.WindowTitle = "$(if($safemode -ne "n"){"[SAFE-MODE] "})$QN$(if($safemode -ne "n"){" [SAFE-MODE]"})"
$trackingsheet = get-item "$ToolShop\05. Share\$('$ToolShopReportPkg')\Orders\2015\ToolShopOrders_*.xlsm" | sort LastWriteTime | Select -last 1
if($safemode -eq "n"){
    $speed = 50
}
else{
    $host.UI.RawUI.ForegroundColor = "White"
    Write-Host "Safe Mode Enabled"
    $speed = 3000
    #$stamp = "[AutoJobCard] $(get-content env:username) $(get-date)"
    if((get-content env:username) -eq "bredestegej"){$speed = 1000}
}
$machine = @("Raw bandsaw","CAM","manual workstation","CNC Lathe","manual workstation","5-axis Mill","ID/OD grinder","Surface grinder","Wire cut EDM","5-axis Mill","CNC Lathe","5-axis Mill","CNC Lathe","manual workstation")
$cost1 = @("30","","55","85","55","105","85","65","55","105","85","105","85","55")
$cost2 = @("6410255750","","6410255770","6410255710","6410255770","6410255717","6410255735","6410255730","6410255745","6410255717","6410255710","6410255717","6410255710","6410255770")
$z = $null

Write-host "`nLoading tracking sheet info..." -ForegroundColor DarkGreen
try{
    attrib +r $trackingsheet
    $sheetname = "active"
    $before = @(Get-Process [e]xcel | %{$_.Id})
    $xl = New-Object -ComObject Excel.Application
    $ExcelId = Get-Process excel | %{$_.Id} | ?{$before -notcontains $_}
    $wb = $xl.Workbooks.Open($trackingsheet)
    $ws = $wb.Worksheets.Item($sheetName)
    $xl.Visible=$false
}
catch{
    Write-Host "Error: Cannot Access server! [ER3]" -ForegroundColor yellow
    $trackingsheet
    pause
    exit
}

foreach ($item in $Search){
    try{ # Tried to search tracking sheet for the entered quote number. if not found returns error and is cause to tell user 'not found'
        $mainRng = $ws.usedRange
        $mainRng.Select() | out-null
        $objSearch = $mainrng.Find($item)
        $First = $objSearch
        Do{
            $objSearch = $mainrng.FindNext($objSearch)
            $row = $objSearch.row
            Write-Host "$($ws.Cells.Item($row,4).text)" -ForegroundColor DarkGreen
            $z += ,($ws.Cells.Item($row,1).text, $ws.Cells.Item($row,3).text, $ws.Cells.Item($row,4).text, $ws.Cells.Item($row,6).text, $ws.Cells.Item($row,10).text, $ws.Cells.Item($row,"BN").text, $ws.Cells.Item($row,"BO").text, $ws.Cells.Item($row,"BR").text, $ws.Cells.Item($row,"BS").text, $ws.Cells.Item($row,"BT").text, $ws.Cells.Item($row,8).text)
            if($QN -eq "Custom"){$quote = $ws.Cells.Item($row,5).text}
        }
        While ($objSearch -ne $NULL -and $objSearch.AddressLocal() -ne $First.AddressLocal())
    }
    catch{ # ಠ_ಠ
        if($QN -eq "Custom"){
            Write-Host "No info for $item" -ForegroundColor DarkGreen
        }
        else{
            if($objSearch -eq $null){
                Write-Host "Error: Quote not found in tracking sheet" -ForegroundColor yellow
            }
            else{
                Write-Host "Unknown Error: Program must close" -ForegroundColor Red
            }
            Read-Host "Press Enter to Exit"
            Stop-Process -Id $ExcelId -Force -ErrorAction SilentlyContinue
            Exit
        }
    }
}
Stop-Process -Id $ExcelId -Force -ErrorAction SilentlyContinue

if(!$checkbox){
    write-host "Loading Quote..." -ForegroundColor DarkGreen
    if($QN -eq "QHC"){
        $Quotesheet = get-item "$ToolShop\05. Share\11. Tool Shop Quotes\01_New-parts\03. Quotes 2015\$QN*.xlsm"
    }
    else{
        $QN = $QN.Trim("Q")
        $qnx = $qn
        if($qn -match "B"){
            $qnx = $qn.trim("-B")
        }
        if($qn -match "C"){
            $qnx = $qn.trim("-C")
        }
        $Quotesheet = get-item "$ToolShop\05. Share\11. Tool Shop Quotes\01_New-parts\03. Quotes 2015\$QNx*\$QN*.xlsm"
        if($Quotesheet -eq $null){
            Write-host "Quote not found, press enter to exit" -foregroundcolor yellow
            Read-Host
            exit
        }
    }
    $before = @(Get-Process [e]xcel | %{$_.Id})
    $xl = New-Object -ComObject Excel.Application
    $ExcelId = Get-Process excel | %{$_.Id} | ?{$before -notcontains $_}
    $wb = $xl.Workbooks.Open($Quotesheet)
    $xl.Visible=$false
    $poscnt = $null

    $before = @(Get-Process [e]xcel | %{$_.Id})
    $xl2 = New-Object -ComObject Excel.Application
    $ExcelId2 = Get-Process excel | %{$_.Id} | ?{$before -notcontains $_}
    $wb2 = $xl2.Workbooks.Open("T:\04. Blue Prints\Translations.xlsx")
    $ws2 = $wb2.Worksheets.Item("Translations")
    $xl2.Visible=$false

    if($wb.Worksheets.Item("complete").Cells.Item(1,"B").text -ne 5.1){ # Makes sure quote sheet is newest acceptable version
        Write-Host "$Quotesheet" -ForegroundColor Yellow
        $continue = Read-host "Required: Quote version 5.1 or higher`nActivity description may be messed up`nContinue? (y/n)"
        if($continue -ne "y"){
            Stop-Process -Id $ExcelId -Force -ErrorAction SilentlyContinue
            Stop-Process -Id $ExcelId2 -Force -ErrorAction SilentlyContinue
            exit
        }
    }

    [int]$poscnt = $wb.Worksheets.Item("complete").Cells.Item(2,"B").value() # gets position count
    $Jobname = $wb.Worksheets.Item("Sheet1").Cells.Item(17,"B").text
    $position = 0
}

if($QN -eq "QHC"){
    Write-Host "Gathing Standard Part List" -ForegroundColor DarkGreen
    $STDparts=$null
    for($x = 2; $x -le $poscnt; $x++){
        $ws = $wb.Worksheets.Item("part_$x")
        $STDparts += ,("$($x-1)) $($ws.Cells.Item(6,5).text)")
    }
    
    $maxcnt = $z.count
    $j = 0
    foreach ($item in $z){$j++
        clear-host
        Write-host "$maxcnt Matched jobs for PO $PO (Press Enter to skip)`n"
        for($i=0; $i -lt $STDparts.Count; $i++){Write-host $STDparts[$i] -ForegroundColor Cyan}
        Write-host "`n$($item[2]) / $($item[9]) / $($item[7]) / $($item[8]) / $($item[10])" -ForegroundColor green
        $x=$null
        [int]$x = (Read-Host "($j/$maxcnt) Enter the number that best matches")
        if($x -ne ""){$x++}
        if($x -ne ""){
            if($item[10] -lt 1){
                $item[10] = Read-Host "Enter Quantity"
            }
            if($($item[9]) -match ","){$sugg = $($z[$5][9]).Substring(0, $($z[$5][9]).IndexOf(','))}else{$sugg = $($item[9])}
            Write-Host "Suggested Part name: '$sugg'"
            $part = Read-Host "Enter name"
            if($part -eq ""){$part = $sugg}
            $drawing = Read-Host "Enter Drawing Number"
        }
        $STDP += ,($x,$part,$drawing)
        $drawing = $null
    }
    $i=0
    foreach ($item in $z){
        Write-host "$($item[2]) - $($item[9]) - $($stdp[$i])"
        $i++
    }
    Clear-Host
}else{
    $maxcnt = $poscnt
}

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
$type::SetWindowPos($handle, $alwaysOnTop, 0, 0, 0, 0, 0x0003)


write-host "Launch and Load TM" -ForegroundColor DarkGreen
get-process TM_schrank -ErrorAction SilentlyContinue | stop-process -ErrorAction SilentlyContinue
for($c = 1; $c -le 1; $c++){
    ii "C:\Program Files (x86)\TMSchrank\TM_Schrank.exe"
    start-sleep -s 3
    $TMwindow = (Get-Process tm_schrank).MainWindowHandle
    [void] [Tricks]::SetForegroundWindow($TMwindow)

    add-type -AssemblyName microsoft.VisualBasic
    add-type -AssemblyName System.Windows.Forms
    start-sleep -Milliseconds 500

    [System.Windows.Forms.SendKeys]::SendWait("~~092214~{tab}{tab}~{tab}~+({tab}{tab}{tab}{tab})~{tab}{tab}{tab}{tab}{tab}{tab}") # Standard quick log in (as toolshop intern)
    try{get-process tm_schrank | out-null;Write-host "TM Running" -ForegroundColor DarkGreen}catch{$c = 0,"TM Failed"}
}
[void] [Tricks]::SetForegroundWindow($PSwindow)

for($5 = 0; $5 -lt $z.count; $5++){Add-Content -Path "$ToolShop\05. Share\35. Time Studies\Logs\$QN" -Value "$($5+1)) $($z[$5][2]) / $($z[$5][9]) / $($z[$5][7]) / $($z[$5][8])"}
for($count = 1; $count -le $maxcnt; $count++){
    TM-respond $ExcelId $PSwindow
    Write-Host "        [Diagnostics] 5: $5  x: $x" -foregroundcolor black
    try{get-process tm_schrank | out-null}catch{Write-host "TM check failed" -ForegroundColor RED}
    write-host "Loading quote info for part $count/$maxcnt..."
    if(!$checkbox){
        if($QN -eq "QHC"){
            $5 = $count-1
            try{
                $ws = $wb.Worksheets.Item("part_$($STDP[$5][0])")
                $Quantity = $z[$5][10]
                $part = $STDP[$5][1]
                $drawing = $STDP[$5][2]
            }catch{$5=$null}
        }
        else{
            $ws = $wb.Worksheets.Item("part_$count")
            $Quantity = $ws.Cells.Item(7,2).text
            $Part = ($ws.Cells.Item(6,5).text).trim(" ")
            $drawing = $ws.Cells.Item(7,5).text
        }
        $Material = $ws.Cells.Item(8,2).text
        if($QN -eq "QHC"){Write-Host ""}
        try{ 
            $mainRng = $ws2.usedRange
            $mainRng.Select() | out-null
            $objSearch = $mainrng.Find($Part)
            $First = $objSearch
            Do{
                $objSearch = $mainrng.FindNext($objSearch)
                $row = $objSearch.row
                if($part -contains $ws2.Cells.Item($row,2).text){
                    $part = "$($ws2.Cells.Item($row,3).text) ($part)"
                }
            }
            While ($objSearch -ne $NULL -and $objSearch.AddressLocal() -ne $First.AddressLocal())
        }catch{}
        $heatreat = $ws.Cells.Item(24,"C").text
        $blackoxid = $ws.Cells.Item(26,"C").text
        $nitriding = $ws.Cells.Item(28,"C").text
        $coating = $ws.Cells.Item(30,"C").text
        $purchase = $ws.Cells.Item(51,"D").text
        if($purchase -eq "y"){
            $material = "Purchase Part"
        }#elseif($safemode -eq "n"){$material = $null}
        $positions = $null
        for($X=35; $X -le 48; $X++){
            $positions += ,($ws.Cells.Item($X,1).text, $ws.Cells.Item($X,4).text, $ws.Cells.Item($X,7).text)
        }
    }else{

        

    }
    try{
        if($z[1][0] -and $QN -ne "QHC"){
            for($v = 0; $v -lt $z.count; $v++){
                $drawsugg = $null
                if($drawing -ne "" -and ($z[$v][7] -match $drawing -or $z[$v][8] -match $drawing -or $z[$v][9] -match $drawing)){
                    $drawsugg = "$($v+1)) $($z[$v][7]) $($z[$v][8]) $($z[$v][9])"
                    if($part -ne "" -and ($z[$v][7] -match $part -or $z[$v][8] -match $part -or $z[$v][9] -match $part)){
                        $5 = $v
                        break
                    }
                }
                else{$5 = $null}
            }
            if($5 -eq $null -and ($part -ne "" -and $drawing -ne "")){
                if($drawsugg){Write-Host "Suggested Part: $drawsugg" -ForegroundColor Magenta}else{for($5 = 0; $5 -lt $z.count; $5++){Write-host "$($5+1)) $($z[$5][2]) / $($z[$5][9]) / $($z[$5][7]) / $($z[$5][8])" -ForegroundColor Cyan}}
                $5 = (Read-Host "Enter the number that '$part [$drawing]' best matches")-1
            }else{$5--}
            if($5 -ge 0){
                if($Quantity -ne $z[$5][10] -or $Quantity -eq "" -or $z[$5][10] -eq ""){
                    Write-Host "`nQuantities invalid! [ER1]" -ForegroundColor Yellow
                    Write-Host "Tracking sheet: $($z[$5][10])`nQuote: $quantity" -ForegroundColor Cyan
                    $Quantity = Read-Host "Enter quantity for $part"
                    $jobname = $part
                }else{Write-Host "Quantities Match: Tracking sheet: $($z[$5][10]) Quote: $quantity"}
            }
        }
        if($bookmark -contains $($z[$5][2])){
            foreach ($item in $bookmark){
                if($item[0] -eq "$($z[$5][2])"){
                    $item[1]=$item[1]+100
                    $position = $item[1]
                }
            }
        }
        elseif($5 -ge 0){
            $position = 100
            $bookmark += ,($($z[$5][2]), $position, $false)
        }
    }
    catch{
        if($QN -ne "QHC"){
            $5 = 0
            $bookmark = ,($($z[$5][2]), $position, $false)
        }
        $position = $position + 100
        if($Quantity -ne $z[$5][10] -and $z[$5][10] -gt 10){
            Write-Host "Quantities invalid! [ER2]" -ForegroundColor Yellow
            Write-Host "Tracking sheet: $($z[$5][10])`nQuote: $quantity"
            $Quantity = Read-Host "Enter quantity for $part"
        }
    }
    IF($5 -ge "0"){
        if($drawing -eq "?"){
            Write-Host "`nNo Drawing Number! [ER4]" -ForegroundColor Yellow
            $drawing = Read-Host "Enter Drawing number for $part"
        }
        if($drawing -ne "?"){
            $drawings += ,(($drawing -replace [regex]::Escape(' '),'_' -replace [regex]::Escape('-'),'_'), $($z[$5][2]), $position)
        }


        if($heatreat -eq "y" -and $maxcnt -lt 20){Write-host "`n$Part $drawing";$heatreat = Read-Host "Enter Hardness"}
        if($heatreat -notmatch "HRC" -and $heatreat -ne "None"){$heatreat = "$heatreat HRC"}
        $heatreat = $heatreat -replace [regex]::Escape('+'),'{+}'
        if($blackoxid -eq "y"){$blackoxid = "Black Oxid Finish"}
        if($nitriding -eq "y"){$nitriding = "Nitriding"}
        if($coating -ne "None"){
            if($coating -eq "y"){$coating = "Coating"}
            elseif($coating = "Coating"){}
            else{$coating = "Coating: $coating"}
        }
        if($Jobname -eq "?"){$jobname = $part}

        if($z[$5][7].length -lt 8 -and $z[$5][7] -notmatch "/"){$ID = $z[$5][7]}else{$id = $null}

        if($z[$5][8] -match "/" -and ($z[$5][8]).length -lt 8){$WZ = $z[$5][8]}
        elseif($z[$5][7] -gt 10 -and $z[$5][7] -match "/"){
            $splitter = $z[$5][7].split(" ")
            if($splitter[0].length -lt 8 -and $splitter[0] -notmatch "/" -and $splitter[0].length -gt 3){$ID = $splitter[0]}
            if($splitter[1] -match "/"){$WZ = $splitter[1]}
        }
        else{$WZ=$null}

        $part = $part -replace [regex]::escape('('),'{(}' -replace [regex]::escape(')'),'{)}'
        [void] [Tricks]::SetForegroundWindow($TMwindow)
        start-sleep -Milliseconds 500
        $operation = 10
        for($XX = 0; $xx -le 12; $XX++){
            $opname = $positions[$XX][0]
            [float]$optime = $positions[$XX][1]
            $opmchn = $machine[$XX]
            $CPH = $cost1[$XX]
            $CostCenter = $cost2[$XX]
            if($QN -eq "QHC"){$optime = $optime*$Quantity}
            if($xx -gt 7 -and $xx -lt 12){
                $activitydesc = $positions[$XX][0]
                if($XX -eq 9 -or $xx -eq 11){$opname = "CNC Turning"}else{$opname = "Multiaxial Milling"}
            }elseif($positions[$XX][2] -eq $opmchn){$activitydesc = $null
            }else{$activitydesc = $positions[$XX][2]}
            if($XX -eq 5 -and $heatreat -ne "None"){
                TM-respond $ExcelId $PSwindow
                Write-host "$($z[$5][2]) $("{0:D4}" -f ($position+$operation)) Heat Treat: $heatreat" -ForegroundColor DarkGreen
                [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$z[$5][2]}else{"joe"}){tab}$($z[$5][1]){tab}$($z[$5][3]){tab}$($z[$5][0]){tab}{tab}$($z[$5][5]){tab}$("{0:D4}" -f $position){tab}$Part{tab}$drawing{tab}$Quantity{tab}$($z[$5][6]){tab}Offen{tab}$("{0:D4}" -f $operation){tab}External Heat Treatment{tab}$(if($heatreat -notmatch "y"){$heatreat}){tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}$ID{tab}$WZ{tab}{tab}$stamp{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
                TM-return $speed
                $operation=$operation+10
            }
            if($xx -eq 12 -and $material -match "390" -and $part -match "die"){ #Tempering
                TM-respond $ExcelId $PSwindow
                Write-host "$($z[$5][2]) $("{0:D4}" -f ($position+$operation)) Tempering: $part" -ForegroundColor DarkGreen
                [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$z[$5][2]}else{"joe"}){tab}$($z[$5][1]){tab}$($z[$5][3]){tab}$($z[$5][0]){tab}{tab}$($z[$5][5]){tab}$("{0:D4}" -f $position){tab}$Part{tab}$drawing{tab}$Quantity{tab}$($z[$5][6]){tab}Offen{tab}$("{0:D4}" -f $operation){tab}External Heat Treatment{tab}Tempering{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}$ID{tab}$WZ{tab}{tab}$stamp{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
                TM-return $speed
                $operation=$operation+10
            }
            if($xx -eq 12 -and $coating -ne "None"){
                TM-respond $ExcelId $PSwindow
                Write-host "$($z[$5][2]) $("{0:D4}" -f ($position+$operation)) $Coating" -ForegroundColor DarkGreen
                [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$z[$5][2]}else{"joe"}){tab}$($z[$5][1]){tab}$($z[$5][3]){tab}$($z[$5][0]){tab}{tab}$($z[$5][5]){tab}$("{0:D4}" -f $position){tab}$Part{tab}$drawing{tab}$Quantity{tab}$($z[$5][6]){tab}Offen{tab}$("{0:D4}" -f $operation){tab}External Heat Treatment{tab}$coating{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}$ID{tab}$WZ{tab}{tab}$stamp{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
                TM-return $speed
                $operation=$operation+10
            }
            if($xx -eq 12 -and $nitriding -ne "None"){
                TM-respond $ExcelId $PSwindow
                Write-host "$($z[$5][2]) $("{0:D4}" -f ($position+$operation)) Nitriding: $part" -ForegroundColor DarkGreen
                [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$z[$5][2]}else{"joe"}){tab}$($z[$5][1]){tab}$($z[$5][3]){tab}$($z[$5][0]){tab}{tab}$($z[$5][5]){tab}$("{0:D4}" -f $position){tab}$Part{tab}$drawing{tab}$Quantity{tab}$($z[$5][6]){tab}Offen{tab}$("{0:D4}" -f $operation){tab}External Heat Treatment{tab}$Nitriding{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}$ID{tab}$WZ{tab}{tab}$stamp{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
                TM-return $speed
                $operation=$operation+10
            }
            if($optime -ne "0" -and $optime -ne "0.0"){
                TM-respond $ExcelId $PSwindow
                if($xx -eq 12){$stamp = "Additional info: $($z[$5][7])   $($z[$5][8])   $($z[$5][9])"}
                Write-host "$($z[$5][2]) $("{0:D4}" -f ($position+$operation)) $($opname): $opmchn  Time: $optime $activitydesc" -ForegroundColor DarkGreen
                [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$z[$5][2]}else{"joe"}){tab}$($z[$5][1]){tab}$($z[$5][3]){tab}$($z[$5][0]){tab}{tab}$($z[$5][5]){tab}$("{0:D4}" -f $position){tab}$Part{tab}$drawing{tab}$Quantity{tab}$($z[$5][6]){tab}Offen{tab}$("{0:D4}" -f $operation){tab}$opname{tab}$activitydesc{tab}$opmchn{tab}$CPH{tab}$CostCenter{tab}$optime{tab}{tab}{tab}{tab}$ID{tab}$WZ{tab}{tab}$stamp{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
                if($xx -eq 12){$Stamp = $null}
                TM-return $speed
                $operation=$operation+10
            }
        }
        if($position -eq 100){
            #Auftrag
            TM-respond $ExcelId $PSwindow
            $stamp = "$(get-content env:username)   $(get-date)   $QN"
            Write-host "$($z[$5][2]) 0000 Auftrag: $Jobname" -ForegroundColor DarkGreen
            [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$z[$5][2]}else{"joe"}){tab}$($z[$5][1]){tab}$($z[$5][3]){tab}$($z[$5][0]){tab}$($z[$5][4]){tab}$($z[$5][5]){tab}{tab}$Jobname{tab}{tab}{tab}$($z[$5][6]){tab}Offen{tab}{tab}AUFTRAG GESAMT{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}$ID{tab}$WZ{tab}{tab}$stamp{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
            $stamp = $null
            TM-return $speed
        }

        #position
        TM-respond $ExcelId $PSwindow
        Write-host "$($z[$5][2]) $("{0:D4}" -f $position) Position: $("{0:D4}" -f $position)" -ForegroundColor DarkGreen
        [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$z[$5][2]}else{"joe"}){tab}$($z[$5][1]){tab}$($z[$5][3]){tab}$($z[$5][0]){tab}{tab}$($z[$5][5]){tab}$("{0:D4}" -f $position){tab}$Part{tab}$drawing{tab}$Quantity{tab}$($z[$5][6]){tab}Offen{tab}{tab}POSITION GESAMT{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}$ID{tab}$WZ{tab}{tab}$stamp{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
        TM-return $speed

        #material
        TM-respond $ExcelId $PSwindow
        Write-host "$($z[$5][2]) $("{0:D4}" -f ($position+5)) Material: $material" -ForegroundColor DarkGreen
        [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$z[$5][2]}else{"joe"}){tab}$($z[$5][1]){tab}$($z[$5][3]){tab}$($z[$5][0]){tab}{tab}$($z[$5][5]){tab}$("{0:D4}" -f $position){tab}$Part{tab}$drawing{tab}$Quantity{tab}$($z[$5][6]){tab}Offen{tab}0005{tab}Material{tab}$material{tab}manual workstation{tab}{tab}{tab}{tab}{tab}{tab}{tab}$ID{tab}$WZ{tab}{tab}$stamp{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
        TM-return $speed
    }
    [void] [Tricks]::SetForegroundWindow($PSwindow)
}
Stop-Process -Id $ExcelId -Force -ErrorAction SilentlyContinue
Stop-Process -Id $ExcelId2 -Force -ErrorAction SilentlyContinue
Write-Host "        [Diagnostics] Bookmark count: $($bookmark.count)" -ForegroundColor Black
$Print = $null
if($QN -ne "Custom"){$Print = read-host "Print? (y/n)"}else{$print = $checkbox[0][6]}
start-sleep -m 500
if($Print -eq "y" -or $print -eq $true){
    get-process TM_schrank -ErrorAction SilentlyContinue | stop-process -ErrorAction SilentlyContinue
    for($c = 1; $c -le 1; $c++){
        ii "C:\Program Files (x86)\TMSchrank\TM_Schrank.exe"
        start-sleep -s 3
        $TMwindow = (Get-Process tm_schrank).MainWindowHandle
        [void] [Tricks]::SetForegroundWindow($TMwindow)

        add-type -AssemblyName microsoft.VisualBasic
        add-type -AssemblyName System.Windows.Forms
        start-sleep -Milliseconds 500

        [System.Windows.Forms.SendKeys]::SendWait("~~092214~{tab}{tab}~{tab}~+({tab}{tab}{tab}{tab})~{tab}{tab}{tab}{tab}{tab}{tab}") # Standard quick log in (as toolshop intern)
        try{get-process tm_schrank | out-null;Write-host "TM Running" -ForegroundColor DarkGreen}catch{$c = 0,"TM Failed"}
    }
    $action = "Printing"
    for($5 = 0; $5 -lt $bookmark.count; $5++){
        [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$bookmark[$5][0]}else{"joe"})~+({tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab})")
        for($b = 100; $b -le $bookmark[$5][1]; $b = $b + 100){
            TM-respond $ExcelId $PSwindow
            if($b -eq 100 -and $($bookmark[$5][1]) -gt 100){
                start-sleep -s 1
                [System.Windows.Forms.SendKeys]::SendWait("~")
                start-sleep -s 1
                [System.Windows.Forms.SendKeys]::SendWait("~")
                start-sleep -s 2
                [System.Windows.Forms.SendKeys]::SendWait("5~")
                start-sleep -s 2
                [System.Windows.Forms.SendKeys]::SendWait("~")
                start-sleep -s 1
                Write-Host "Printing: $($bookmark[$5][0]) Cover Sheet " -ForegroundColor DarkGreen
            }
            write-host "Printing: $($bookmark[$5][0]) $("{0:D4}" -f $b)/$($bookmark[$5][1])" -ForegroundColor DarkGreen
            [System.Windows.Forms.SendKeys]::SendWait("~")
            start-sleep -s 1
            [System.Windows.Forms.SendKeys]::SendWait("~")
            start-sleep -s 2
            [System.Windows.Forms.SendKeys]::SendWait("1~$b~")
            start-sleep -s 2.5
            [System.Windows.Forms.SendKeys]::SendWait("~")
            start-sleep -s 1
            foreach ($item in $drawings){
                if($item[1] -eq $($bookmark[$5][0]) -and $QN -ne "QHC"){
                    if($item[2] -eq $b){
                        get-childitem "$ToolShop\05. Share\11. Tool Shop Quotes\01_New-parts\03. Quotes 2015\$QN*\" -recurse -Exclude "*quote*" | where {$_.extension -eq ".pdf" -and $_.BaseName -match $item[0]} | % {
                            Start-Process -FilePath $_.FullName –Verb Print -PassThru | %{sleep 10;$_} | kill
                            Write-host "Printing: $($_.BaseName)"
                        }
                    }
                }
            }
        }
        $bookmark[$5][2] = $true
        [System.Windows.Forms.SendKeys]::SendWait("{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
    }
    
    <#foreach ($item in $drawings){
        get-childitem "$ToolShop\05. Share\11. Tool Shop Quotes\01_New-parts\03. Quotes 2015\$QN*\" -recurse -Exclude "*quote*" | where {$_.extension -eq ".pdf" -and $_.BaseName -match $item} | % {
            Start-Process -FilePath $_.FullName –Verb Print -PassThru #| %{sleep 5;$_} | kill
        }
    }#>
}

#Print check
if($Print -eq "y"){
    TM-respond $ExcelId $PSwindow
    if($bookmark -contains $false){Write-Host "`nSome items failed to print!" -ForegroundColor Yellow}
    else{Write-host "`nPrint Successful"}
}
[void] [Tricks]::SetForegroundWindow($PSwindow)
Read-host "Press Enter to Exit"
stop-transcript


Add-Content -Path $log -Value '[Diagnostics] $z Part info'
Add-Content -Path $log -Value $z
Add-Content -Path $log -Value '[Diagnostics] $bookmark'
Add-Content -Path $log -Value $bookmark
Add-Content -Path $log -Value '[Diagnostics] Errors:'
Add-Content -Path $log -Value $Error