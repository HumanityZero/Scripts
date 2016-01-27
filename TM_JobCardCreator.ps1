#version 3.0 for Quote sheet version 5.3
$year = get-date -f yyyy
$ToolShop = "T:"
$ToolShop = '\\10.51.0.12\DFSRoots\Tool Shop'

$host.UI.RawUI.ForegroundColor = "green"
$host.UI.RawUI.BackgroundColor = "black"

$ErrorActionPreference = "Stop"

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
    $log = "$ToolShop\05. Share\35. Time Studies\Program Files\Logs\TM-$serial.log"
    Start-Transcript -path $log -append | out-null
}
catch{
    $log = "C:\Windows\Temp\TM-$serial.log"
    Start-Transcript -path $log -append | out-null
}


if($PSVersionTable.PSVersion.major -lt 3){
    write-host "This script requires powershell version 3.0+ (you have V$("{0:N1}" -f $PSVersionTable.PSVersion.major))"
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

cd "$ToolShop\05. Share\35. Time Studies\Program Files"

#disclaimer
if((get-content env:username) -ne "bredestegej"){
    Clear-Host
    if($PSVersionTable.PSVersion.major -lt 4){Write-Host "ATTENTION: You are running an outdated version of powershell (V$("{0:N1}" -f $PSVersionTable.PSVersion.major)) consider updating." -ForegroundColor Red}
    Write-Host "Session ID: TM-$serial" -ForegroundColor DarkGreen
    Write-Host "`nNOTE: It is very important that you do not click with your mouse.`nIf the program starts to run haywire close TM." -ForegroundColor Yellow
    Write-Host "`nJob Card Creator version 4.0" -ForegroundColor Cyan
    Write-Host "`nJoe Bredestege`n513-373-9531`n" -ForegroundColor DarkGreen
    Pause
}

function 11x17 {
    if((get-content env:computername) -eq "FLO-8212TRA-PC"){
        $strComputer = "."
        $X = (get-wmiobject -class "Win32_Printer" -Filter "" -namespace "root\CIMV2" ` -computername $strComputer) | where {$_.Name -eq "\\flofile01.us.mubea.net\8212 - Toolshop Canon iR3235"}
        $X.SetDefaultPrinter() | Out-Null
    }
}

function letter {
    if((get-content env:computername) -eq "FLO-8212TRA-PC"){
        $strComputer = "."
        $X = (get-wmiobject -class "Win32_Printer" -Filter "" -namespace "root\CIMV2" ` -computername $strComputer) | where {$_.Name -eq "\\flofile01\8212 - Toolshop Canon iR3235"}
        $X.SetDefaultPrinter() | Out-Null
    }
}

function TM-return { # Saves position and returns to job No. space
    param([float]$speed)
    start-sleep -m $speed
    [System.Windows.Forms.SendKeys]::SendWait("~~~")
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
        <CheckBox Name="HM" Content="Milling Finish" HorizontalAlignment="Left" Margin="189,140,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="HT" Content="Turning Finish" HorizontalAlignment="Left" Margin="189,161,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="HMF" Content="Jig Grinding" HorizontalAlignment="Left" Margin="189,182,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="HTF" Content="Grinding Finish" HorizontalAlignment="Left" Margin="189,203,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="MW" Content="Manual Workstations" HorizontalAlignment="Left" Margin="299,98,0,0" VerticalAlignment="Top" IsChecked="True"/>
        
        <Button Name="btnExit" Content="Ok" HorizontalAlignment="Left" Margin="353,267,0,0" VerticalAlignment="Top" Width="75"/>
        
        <CheckBox Name="NIT" Content="Nitriding" HorizontalAlignment="Left" Margin="168,259,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="COT" Content="Coating" HorizontalAlignment="Left" Margin="168,238,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="Heatreat" Content="Heat Treatment" HorizontalAlignment="Left" Margin="57,238,0,0" VerticalAlignment="Top" Width="103"/>
        <CheckBox Name="BOF" Content="Black Oxide Finish" HorizontalAlignment="Left" Margin="168,280,0,0" VerticalAlignment="Top"/>
        
        <TextBox Name="PN" HorizontalAlignment="Left" Height="38" Margin="10,10,0,0" TextWrapping="Wrap" Text="Part Name" VerticalAlignment="Top" Width="341" FontSize="18.667"/>
        <TextBox Name="Quantity" HorizontalAlignment="Left" Height="38" Margin="356,10,0,0" TextWrapping="Wrap" Text="Quantity" VerticalAlignment="Top" Width="74" FontSize="16"/>
        <TextBox Name="DN" HorizontalAlignment="Left" Height="30" Margin="10,53,0,0" TextWrapping="Wrap" Text="Drawing Number" VerticalAlignment="Top" Width="420" AutoWordSelection="True" FontSize="16"/>
        <TextBox Name="Material" HorizontalAlignment="Left" Height="23" Margin="10,119,0,0" TextWrapping="Wrap" Text="Material" VerticalAlignment="Top" Width="60"/>
        <TextBox Name="HRD" HorizontalAlignment="Left" Height="23" Margin="57,259,0,0" TextWrapping="Wrap" Text="Hardness" VerticalAlignment="Top" Width="106" FontSize="13.333"/>
        <TextBox Name="XPPS" HorizontalAlignment="Left" Height="28" Margin="338,234,0,0" TextWrapping="Wrap" Text="$XPPS1" VerticalAlignment="Top" Width="90" FontSize="13.333"/>
        <CheckBox Name="Print" Content="Print?" HorizontalAlignment="Left" Margin="376,294,0,0" VerticalAlignment="Top" Width="52" IsChecked="True"/>
        <CheckBox Name="More" Content="More" HorizontalAlignment="Left" Margin="322,294,0,0" VerticalAlignment="Top" Width="49" IsChecked="False"/>

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
$checkbox = @($PN.text.trim(" "),$DN.text.trim(" "),$XPPS.text.trim(" "),$Material.text.trim(" "),$Quantity.text.trim(" "),$Print.IsChecked,$Mat.IsChecked,$CAM.IsChecked,$SAW.IsChecked,$ST.IsChecked,$CN.IsChecked,$SM.IsChecked,$MAM.IsChecked,$CG.IsChecked,$SG.IsChecked,$EDM.IsChecked,$HM.IsChecked,$HT.IsChecked,$HMF.IsChecked,$HTF.IsChecked,$MW.IsChecked,$Heatreat.IsChecked,$HRD.text.trim(" "),$NIT.IsChecked,$COT.IsChecked,$BOF.IsChecked,$More.IsChecked)

return $checkbox
}

#if((get-content env:username) -ne "bredestegej"){$ErrorActionPreference= "Silently Continue"}
$checkbox = $null
$XPPS1 = 1450
Clear-Host
Write-Host "Enter quote number to pull from quote`nType 'Standard' to pull from standard parts list`nType 'Custom' to enter custom job`nType 'Load' to load last custom job" -ForegroundColor White
$QN = read-host "Sudo"
if($QN -match "2015"){
    $QuoteFolder = "03. Quotes 2015"
}elseif($QN -match "2016"){
    $QuoteFolder = "04. Quotes 2016"
}elseif($QN -match "2017"){
    $QuoteFolder = "05. Quotes 2017"
}else{$QuoteFolder = "03. Quotes $(Get-Date -f yyyy)"}
if($QN -eq "Standard"){
    Write-host "WARNING: This function is in beta and may not work right." -ForegroundColor Yellow
    $PO = Read-host "Enter Customer PO"
    $Search = $PO
}
elseif($QN -eq "Custom"){
    for($i=0;$i -lt 3000;$i++){
        $checkbox += ,(CheckBox $XPPS1)
        $search += ,($checkbox[$i][2])
        $XPPS1 = $checkbox[$i][2]
        $maxcnt++
        if(!$checkbox[$i][26]){break}
    }
    $checkbox
    clear-host
}elseif($QN -eq "Load"){
    . ./variables.ps1
    foreach($item in $checkbox){Write-host "$($item[0]) $($item[1]) $($item[2])" -ForegroundColor Cyan}
    pause
    $QN = "Custom"
    foreach($item in $checkbox){$Search += ,($item[2])}
}
else{$Search = $QN}
if($QN -eq "Custom"){$safemode = "n"}else{$safemode = Read-Host "Safe mode? (y/n)"} # safe mode names jobs "Joe" and pauses for 3 seconds before saving each position
$host.ui.RawUI.WindowTitle = "$(if($safemode -ne "n"){"[SAFE-MODE] "})$QN$(if($safemode -ne "n"){" [SAFE-MODE]"})"
$trackingsheet = get-item "$ToolShop\05. Share\$('$ToolShopReportPkg')\Orders\2015\ToolShopOrders_*.xlsm" | sort LastWriteTime | Select -last 1
if($safemode -eq "n"){
    $speed = 150
}
else{
    $host.UI.RawUI.ForegroundColor = "White"
    Write-Host "Safe Mode Enabled"
    $speed = 3000
    #$stamp = "[AutoJobCard] $(get-content env:username) $(get-date)"
    if((get-content env:username) -eq "bredestegej"){$speed = 1000}
}


#saver!
if($QN -eq "custom"){
    Remove-Item -Path variables.ps1 -Force -ErrorAction SilentlyContinue
    Add-Content variables.ps1 -value ('$maxcnt' + " = $maxcnt")
    Add-Content variables.ps1 -value ('$checkbox' + " = 1..$maxcnt")
    for($u=0;$u-lt$maxcnt;$u++){
        Add-Content variables.ps1 -value ('$checkbox'+"[$u] = '$($checkbox[$u][0])','$($checkbox[$u][1])','$($checkbox[$u][2])','$($checkbox[$u][3])','$($checkbox[$u][4])','$($checkbox[$u][5])','$($checkbox[$u][6])','$($checkbox[$u][7])','$($checkbox[$u][8])','$($checkbox[$u][9])','$($checkbox[$u][10])','$($checkbox[$u][11])','$($checkbox[$u][12])','$($checkbox[$u][13])','$($checkbox[$u][14])','$($checkbox[$u][15])','$($checkbox[$u][16])','$($checkbox[$u][17])','$($checkbox[$u][18])','$($checkbox[$u][19])','$($checkbox[$u][20])','$($checkbox[$u][21])','$($checkbox[$u][22])','$($checkbox[$u][23])','$($checkbox[$u][24])','$($checkbox[$u][25])','$($checkbox[$u][26])'")
    }
}

#load times!@!
. ./MachineTimes.ps1
$z = $null


Write-host "Loading tracking sheet info..." -ForegroundColor DarkGreen
$Search = $Search | select -Unique
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
            if($ws.Cells.Item($row,4).text -ne ""){$z += ,($ws.Cells.Item($row,1).text, $ws.Cells.Item($row,3).text, $ws.Cells.Item($row,4).text, $ws.Cells.Item($row,6).text, $ws.Cells.Item($row,10).text, $ws.Cells.Item($row,"BN").text, $ws.Cells.Item($row,"BO").text, $ws.Cells.Item($row,"BR").text, $ws.Cells.Item($row,"BS").text, $ws.Cells.Item($row,"BT").text, $ws.Cells.Item($row,8).text)}
            if($QN -eq "Custom"){$quote = $ws.Cells.Item($row,5).text}
            $LoopErrorTest += ,($ws.Cells.Item($row,4).text)
            continue
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
    if($LoopErrorTest.length -gt 5){
        if($LoopErrorTest[0] -eq $LoopErrorTest[1] -and $LoopErrorTest[1] -eq $LoopErrorTest[2] -and $LoopErrorTest[2] -eq $LoopErrorTest[3]){
            Write-host "Are you stuck in a loop? The tracking sheet may contain multiple highlighted items! Consider updating the tracking sheet!! [ER22]" -foregroundcolor yellow
        }
    }
}
Stop-Process -Id $ExcelId -Force -ErrorAction SilentlyContinue
if($QN -eq "Custom"){$z += ,("[DIAGNOSTICS] Place Holder")}elseif($z.count -eq 1){$z += ,("[DIAGNOSTICS] Place Holder")}
$z = $z | select -uniq


$before2 = @(Get-Process [e]xcel | %{$_.Id})
$xl2 = New-Object -ComObject Excel.Application
$ExcelId2 = Get-Process excel | %{$_.Id} | ?{$before2 -notcontains $_}
$wb2 = $xl2.Workbooks.Open("$ToolShop\04. Blue Prints\Translations.xlsx")
$ws2 = $wb2.Worksheets.Item("Translations")
$xl2.Visible=$false
if(!$checkbox){
    write-host "Loading Quote..." -ForegroundColor DarkGreen
    if($QN -eq "Standard"){
        $QHC = "Q$($z[0][0])"
        $Quotesheet = get-item "$ToolShop\05. Share\11. Tool Shop Quotes\01_New-parts\$QuoteFolder\$QHC*.xlsm"
    }
    else{
        $QN = $QN.Trim("Q")
        $qnx = ($qn -replace ('[^0-9-]','')).trim('-')
        $Quotesheet = get-item "$ToolShop\05. Share\11. Tool Shop Quotes\01_New-parts\$QuoteFolder\$QNx*\$QN*.xlsm"
        if($Quotesheet -eq $null){
            "$ToolShop\05. Share\11. Tool Shop Quotes\01_New-parts\$QuoteFolder\$QNx*\$QN*.xlsm"
            Write-host "Quote not found, press enter to exit. [ER18]" -foregroundcolor yellow
            Read-Host
            exit
        }
    }
    $quote = $QN
    $before = @(Get-Process [e]xcel | %{$_.Id})
    $xl = New-Object -ComObject Excel.Application
    $ExcelId = Get-Process excel | %{$_.Id} | ?{$before -notcontains $_}
    $wb = $xl.Workbooks.Open($Quotesheet)
    $xl.Visible=$false
    $poscnt = $null

    if($wb.Worksheets.Item("complete").Cells.Item(1,"B").text -ne 5.3){ # Makes sure quote sheet is newest acceptable version
        Write-Host "This Quote is version $($wb.Worksheets.Item("complete").Cells.Item(1,"B").text)" -ForegroundColor Yellow
        $continue = Read-host "Required: Quote version 5.3 or higher`nWill mostly work for 5.2`nContinue? (y/n)"
        if($continue -ne "y"){
            Stop-Process -Id $ExcelId -Force -ErrorAction SilentlyContinue
            Stop-Process -Id $ExcelId2 -Force -ErrorAction SilentlyContinue
            exit
        }
    }

    [int]$poscnt = $wb.Worksheets.Item("complete").Cells.Item(2,"B").value() # gets position count
    $position = 0
}
else{
    
    foreach ($item in $checkbox){
        $Customer = $(
            if($Checkbox[0][2] -match "145024"){"HC"}
            elseif($Checkbox[0][2] -match "145025"){"TS"}
            elseif($Checkbox[0][2] -match "145026"){"SB"}
            elseif($Checkbox[0][2] -match "145027"){"CS"}
            elseif($Checkbox[0][2] -match "145028"){"TRB"}
            elseif($Checkbox[0][2] -match "145029"){"VS"}
            elseif($Checkbox[0][2] -match "145030"){"ITSW"}
            elseif($Checkbox[0][2] -match "145031"){"Rework"}
            elseif($Checkbox[0][2] -match "145032"){"HC-IN"}
        )
        $before3 = @(Get-Process [e]xcel | %{$_.Id})
        $xl3 = New-Object -ComObject Excel.Application
        $ExcelId3 = Get-Process excel | %{$_.Id} | ?{$before3 -notcontains $_}
        $wb3 = $xl3.Workbooks.Open("$ToolShop\05. Share\35. Time Studies\XPPS\$Customer.xlsx")
        $ws3 = $wb3.Worksheets.Item("Machine Times")
        $xl3.Visible=$false
        $rowMax = $ws3.UsedRange.Rows.count

        $times=$null
        try{
            Write-Host "Loading Times for $($item[0]) in $customer" -ForegroundColor DarkGreen
            $mainRng = $ws3.usedRange
            $mainRng.Select() | out-null
            $objSearch = $mainrng.Find($item[0])
            $First = $objSearch
            Do{
                $objSearch = $mainrng.FindNext($objSearch)
                $row = $objSearch.row
                Write-Progress -activity “Searching for $($item[0]) in $customer...” -id 42 -PercentComplete (($row / $rowMax)*100) -CurrentOperation "$row/$rowMax"
                $times += ,([float]($ws3.Cells.Item($row,11).text).split(" ")[0],"",[float]($ws3.Cells.Item($row,9).text).split(" ")[0],[float]($ws3.Cells.Item($row,3).text).split(" ")[0],[float]($ws3.Cells.Item($row,12).text).split(" ")[0],[float]($ws3.Cells.Item($row,2).text).split(" ")[0],[float]($ws3.Cells.Item($row,5).text).split(" ")[0],[float]($ws3.Cells.Item($row,4).text).split(" ")[0],[float]($ws3.Cells.Item($row,10).text).split(" ")[0],"","","","",[float]($ws3.Cells.Item($row,7).text).split(" ")[0])
            }
            While ($objSearch -ne $NULL -and $objSearch.AddressLocal() -ne $First.AddressLocal())
            Write-progress -activity "x” -id 42 -status "x"  -Completed
            for($i=7; $i -le 20; $i++){
                if($item[$i]){
                    for($j=0; $j -lt $times.length; $j++){
                        if($times[$j][$i-7] -gt 0.01){
                            if($item[$i] -eq $true){
                                $item[$i] = $times[$j][$i-7]
                            }
                            else{
                                $item[$i] = (($times[$j][$i-7]+$item[$i])/2)
                            }
                        }
                    }
                    if($i -eq 15 -and $item[15] -eq $true){$item[15] = 10.42}
                    if($item[$i] -eq $true){
                        $item[$i] = 0.01
                    }elseif($item[$i] -ne $false){
                        $item[$i] = $item[$i]*[float]$item[4]
                    }elseif($item[$i] -eq "False"){
                        $item[$i] = 0
                    }
                }
            }
            Write-Host "Total estimated time: $([math]::Round(($item[7]+$item[8]+$item[9]+$item[10]+$item[11]+$item[12]+$item[13]+$item[14]+$item[15]+$item[16]+$item[17]+$item[18]+$item[19]+$item[20]),2)) hours" -ForegroundColor DarkGreen
        }
        catch{
            Write-progress -activity "x” -id 42 -status "x"  -Completed
            Write-Host "No times found" -ForegroundColor DarkYellow
            for($i=7; $i -le 20; $i++){
                if($item[$i] -eq $true){
                        $item[$i] = 0.01
                }
            }
        }
        Stop-Process -Id $ExcelId3 -Force -ErrorAction SilentlyContinue
    }
}
if($QN -match "Standard"){
    Write-Host "Gathering Standard Part List" -ForegroundColor DarkGreen
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
            Write-Host "Suggested Drawing number: '$($item[8])'"
            $drawing = Read-Host "Enter Drawing Number"
            if($drawing -eq ""){$drawing = $item[8]}
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
}elseif(!$maxcnt){
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
$type::SetWindowPos($handle, $alwaysOnTop, 0, 0, 0, 0, 0x0003) | Out-Null


write-host "Launch and Load TM" -ForegroundColor DarkGreen
for($c = 1; $c -le 1; $c++){
    get-process TM_schrank -ErrorAction SilentlyContinue | stop-process -ErrorAction SilentlyContinue
    ii "C:\Program Files (x86)\TMSchrank\TM_Schrank.exe"
    start-sleep -s 3
    $TMwindow = (Get-Process tm_schrank).MainWindowHandle
    [void] [Tricks]::SetForegroundWindow($TMwindow)

    add-type -AssemblyName microsoft.VisualBasic
    add-type -AssemblyName System.Windows.Forms
    start-sleep -Milliseconds 500

    [System.Windows.Forms.SendKeys]::SendWait("~~092214~{tab}{tab}~{tab}~{tab}{tab}") # Standard quick log in (as toolshop intern)
    try{get-process tm_schrank | out-null;Write-host "TM Running" -ForegroundColor DarkGreen}catch{$c = 0,"TM Failed"}
    start-sleep -m 500
    [void] [Tricks]::SetForegroundWindow($PSwindow)
    if($QN -eq "Custom"){
        $continue = Read-Host "Is TM running with the Job No. box selected? (y/n)"
        if($continue -eq "n"){$c = 0}
    }
}

for($5 = 0; $5 -lt $z.count; $5++){Add-Content -Path "$ToolShop\05. Share\35. Time Studies\Program Files\Logs\$QN.txt" -Value "$($5+1)) $($z[$5][2]) / $($z[$5][9]) / $($z[$5][7]) / $($z[$5][8])"}
for($count = 1; $count -le $maxcnt; $count++){
    TM-respond $ExcelId $PSwindow
    Write-Host "        [Diagnostics] 5: $5  x: $x $($z.count)" -foregroundcolor black
    try{get-process tm_schrank | out-null}catch{Write-host "TM check failed" -ForegroundColor RED}
    write-host "Loading info for part $count/$maxcnt..."
    $5 = $count-1
    if(!$checkbox){
        if($QN -eq "Standard"){
            try{
                if(!($STDP[$5][0] -ge 0)){Write-Host "Part Skipped";continue}
                $ws = $wb.Worksheets.Item("part_$($STDP[$5][0])")
                $Quantity = $z[$5][10]
                $part = $STDP[$5][1]
                $drawing = $STDP[$5][2]
            }catch{$5=$null}
        }
        else{
            try{$ws = $wb.Worksheets.Item("part_$count")}catch{Write-host "Invalid Quote Part Count [ER42]" -ForegroundColor Red;continue}
            $Quantity = $ws.Cells.Item(7,2).text
            $Part = ($ws.Cells.Item(6,5).text).trim(" ")
            $Jobname = $wb.Worksheets.Item("Sheet1").Cells.Item(17,"B").text
            $drawing = $ws.Cells.Item(7,5).text
            $price = "$(($ws.Cells.Item(61,5).text))"
            $QuotePrice = $wb.Worksheets.Item("complete").Cells.Item(2,"F")
        }
    }else{
        $OrderNumber = "XXX"
        $XVAS = "290"
        $positions = $null
        for($i=7; $i -le 20; $i++){
            $positions += ,($activity[$i-7],$checkbox[$5][$i],"")
        }
        $Part = $checkbox[$5][0]
        $drawing = $checkbox[$5][1]
        $XPPS = $checkbox[$5][2]
        if($checkbox[$5][3] -ne "Material"){$Material = $checkbox[$5][3]}
        $Quantity = $checkbox[$5][4]
        if($checkbox[$5][21] -eq $true){$Heatreat = "y"}else{$heatreat = "None"}
        if($checkbox[$5][21] -eq $true -and $checkbox[$5][22] -notmatch 'Hardness'){$heatreat = $checkbox[$5][22]}
        if($checkbox[$5][23] -eq $true){$nitriding = "y"}else{$nitriding = "None"}
        if($checkbox[$5][24] -eq $true){$coating = "y"}else{$coating = "None"}
        if($checkbox[$5][25] -eq $true){$blackoxid = "y"}else{$blackoxid = "None"}
        $Customer = $(if($XPPS -match "145024"){"HC"}elseif($XPPS -match "145025"){"TS"}elseif($XPPS -match "145026"){"SB"}elseif($XPPS -match "145027"){"CS"}elseif($XPPS -match "145028"){"TRB"}elseif($XPPS -match "145029"){"VS"}elseif($XPPS -match "145030"){"ITSW"}elseif($XPPS -match "145031"){"Rework"}elseif($XPPS -match "145032"){"HC-IN"})
    }
    "$Part $drawing"
    if($QN -match "Standard"){
        $Material = $ws.Cells.Item(8,2).text
        if($QN -eq "Standard"){Write-Host ""}
        if(Get-Process excel | %{$_.Id} | ?{$_ -match $ExcelId}){
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
            if($positions[44] -contains "Hard Milling"){
                $positions[44][0] = "Milling Finish"
            }
            if($positions[45] -contains "Hard Turning"){
                $positions[45][0] = "Turning Finish"
            }
            if($positions[46] -contains "Hard Milling Finish"){
                Write-host "Quote is messed up and cannot proceed!"
                pause
                exit
            }
            if($positions[47] -contains "Hard Turning Finish"){
                Write-host "Quote is messed up and cannot proceed!"
                pause
                exit
            }
        }
    }
    try{
        if($z[1][0] -and $QN -ne "Standard"){
            $NumberSuggestion = $null
            $drawsugg = $null
            for($v = 0; $v -lt $z.count; $v++){
                if($XPPS -ne $null){
                    if($XPPS -eq $z[$v][2]){
                        $5 = $v
                        break
                    }
                }
                elseif($z -contains "[DIAGNOSTICS] Place Holder" -and $QN -ne "custom"){
                    $5 = 0
                    break
                }
                if($drawing -ne "" -and ($z[$v][7] -match $drawing -or $z[$v][8] -match $drawing -or $z[$v][9] -match $drawing)){
                    $drawsugg += ,("$($v+1)) $($z[$v][7]) $($z[$v][8]) $($z[$v][9])")
                    $NumberSuggestion = $v
                    if($part -ne "" -and ($z[$v][7] -match $part -or $z[$v][8] -match $part -or $z[$v][9] -match $part)){
                        $5 = $v
                        break
                    }
                }
                else{$5 = $null}
            }
            if($drawsugg.length -eq 1){$5 = $NumberSuggestion}
            elseif($z -contains "[DIAGNOSTICS] Place Holder"){$5 = 0}
            elseif($5 -eq $null -and ($part -ne "" -and $drawing -ne "")){
                if($drawsugg){
                    Write-Host "Suggested Parts:" -ForegroundColor Magenta;
                    foreach ($item in $drawsugg){
                        Write-Host "$item" -ForegroundColor Magenta
                    }
                }else{
                    for($5 = 0; $5 -lt $z.count; $5++){Write-host "$($5+1)) $($z[$5][2]) / $($z[$5][4]) / $($z[$5][9]) / $($z[$5][7]) / $($z[$5][8])" -ForegroundColor Cyan}
                }
                $5 = (Read-Host "Enter the number that '$part [$drawing] $price' best matches")-1
            }
            if($5 -ge 0){
                if(($Quantity -ne $z[$5][10] -or $Quantity -eq "" -or $z[$5][10] -eq "") -and $z -notcontains "[DIAGNOSTICS] Place Holder"){
                    Write-Host "`nQuantities invalid! [ER1]" -ForegroundColor Yellow
                    Write-Host "Tracking sheet: $($z[$5][10])`nQuote: $quantity" -ForegroundColor Cyan
                    $Quantity = Read-Host "Enter quantity for $part"
                }#else{Write-Host "Quantities Match: Tracking sheet: $($z[$5][10]) Quote: $quantity"}
                $Material = $ws.Cells.Item(8,2).text
                if($QN -eq "Standard"){Write-Host ""}
                if(Get-Process excel | %{$_.Id} | ?{$_ -match $ExcelId}){
                    $contact = $ws.Cells.Item(4,"B").text
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
                    if($positions[44] -contains "Hard Milling"){
                        $positions[44][0] = "Milling Finish"
                    }
                    if($positions[45] -contains "Hard Turning"){
                        $positions[45][0] = "Turning Finish"
                    }
                    if($positions[46] -contains "Hard Milling Finish"){
                        Write-host "Quote is messed up and cannot proceed!"
                        pause
                        exit
                    }
                    if($positions[47] -contains "Hard Turning Finish"){
                        Write-host "Quote is messed up and cannot proceed!"
                        pause
                        exit
                    }
                }
            }
        }
    }
    catch{
        if($QN -ne "Standard" -and $QN -ne "Custom"){
            $5 = 0
        }
        if($QuotePrice -ne $z[0][10] -and $QN -ne "Custom"){
            $breaker = $null
            Write-Host "Prices NOT match!" -ForegroundColor RED
            $breaker = Read-Host "Build part $part $drawing ?? (y/n)"
            if($breaker -eq "n"){continue}
        }
        try{
            if($Quantity -ne $z[$5][10] -and $z[$5][10] -gt 10){
                Write-Host "Quantities invalid! [ER2]" -ForegroundColor Yellow
                Write-Host "Tracking sheet: $($z[$5][10])`nQuote: $quantity"
                $Quantity = Read-Host "Enter quantity for $part"
            }
        }catch{}
    }
    if($part -eq ""){$SkipReason = "- Part missing from quote"}
    if($drawing -eq ""){$SkipReason = "- Drawing missing from quote"}
    if($checkbox -and ($part -eq "" -or $drawing -eq "") -and $5 -eq $null){
        Write-Host "Error: Part or drawing name is blank! $part $drawing $XPPS" -ForegroundColor Yellow
    }
    if($5 -eq $null){Write-Host "Part Skipped $SkipReason" -ForegroundColor Red;continue}
    IF($5 -ge 0){
        if($drawing -eq "?"){
            Write-Host "`nNo Drawing Number! [ER4]" -ForegroundColor Yellow
            $drawing = Read-Host "Enter Drawing number for $part"
        }
        

        try{
            if($heatreat -eq "y" -and $maxcnt -lt 20){Write-host "`n$Part $drawing";$heatreat = Read-Host "Enter Hardness"}
            if($heatreat -notmatch "HRC" -and $heatreat -ne "None"){$heatreat = "$heatreat HRC"}
            foreach($item in $z){
                $item = $item -replace [regex]::Escape('+'),'{+}' -replace [regex]::Escape('%'),'{%}' -replace [regex]::Escape('^'),'{^}' -replace [regex]::Escape('~'),'{~}' -replace [regex]::Escape('('),'{(}' -replace [regex]::Escape(')'),'{)}'
            }
            $heatreat = $heatreat-replace [regex]::Escape('+'),'{+}' -replace [regex]::Escape('%'),'{%}' -replace [regex]::Escape('^'),'{^}' -replace [regex]::Escape('~'),'{~}' -replace [regex]::Escape('('),'{(}' -replace [regex]::Escape(')'),'{)}'
            if($blackoxid -eq "y"){$blackoxid = "Black Oxide Finish"}
            if($nitriding -eq "y"){$nitriding = "Nitriding"}
            if($nitriding -ne "None"){
                if($nitriding -notmatch "Nitriding"){$nitriding = "Nitriding: $nitriding"}
                $nitriding = $nitriding -replace [regex]::Escape('+'),'{+}' -replace [regex]::Escape('%'),'{%}' -replace [regex]::Escape('^'),'{^}' -replace [regex]::Escape('~'),'{~}' -replace [regex]::Escape('('),'{(}' -replace [regex]::Escape
            }
            if($blackoxid -ne "None"){
                if($blackoxid -notmatch "Black"){$blackoxid = "Black oxide: $blackoxid"}
            }
            if($coating -ne "None"){
                if($coating -eq "y"){$coating = "Coating"}
                elseif($coating = "Coating"){}
                else{$coating = "Coating: $coating"}
            }

            if($z[$5][7].length -le 8 -and $z[$5][7] -notmatch "/"){$ID = $z[$5][7]}else{$id = $null}

            if($z[$5][8] -match "/" -and ($z[$5][8]).length -lt 12){$WZ = $z[$5][8]}
            elseif($z[$5][7] -gt 10 -and $z[$5][7] -match "/"){
                $splitter = $z[$5][7].split(" ")
                if($splitter[0].length -le 8 -and $splitter[0] -notmatch "/" -and $splitter[0].length -gt 3){$ID = $splitter[0]}
                if($splitter[1] -match "/"){$WZ = $splitter[1]}
            }
            else{$WZ=$null}
            if(!$checkbox){if($z[$5][2] -match "1450"){$XPPS = $z[$5][2]}}
            if($z[$5][1] -match '^[0-9]*$'){$XVAS = $z[$5][1]}
            if($z[$5][3] -match '^[0-9]*$'){$OrderNumber = $z[$5][3]}
            $Customer = ($z[$5][0]).trim("!")
            $OrderDate = $z[$5][5]
            $EndDate = $z[$5][6]
            $TotalPrice = ($z[$5][4]).trim("$")
        }
        
        catch{
            if($QN -ne "Standard"){
                $5 = 0
            }
        }

        try{
            $mainRng = $ws2.usedRange
            $mainRng.Select() | out-null
            $objSearch = $mainRng.Find($Part)
            $First = $objSearch
            Do{
                $objSearch = $mainrng.FindNext($objSearch)
                $row = $objSearch.row
                if($part -match $ws2.Cells.Item($row,2).text){
                    $part = "$($ws2.Cells.Item($row,3).text) ($part)"
                }
            }
            While ($objSearch -ne $NULL -and $objSearch.AddressLocal() -ne $First.AddressLocal())
        }catch{<#Write-Host "No Translation"#>}
        $dont = $false
            foreach ($item in $bookmark){
                if($item[0] -contains $XPPS){
                    $item[1]=$item[1]+100
                    $position = $item[1]
                    $dont = $true
                }
            }
        if($5 -ge 0 -and $dont -eq $false){
            $position = 100
            $bookmark += ,("$XPPS", $position, $false)
        }
        if($drawing -ne "?" -and $drawing -ne "N/A" -and $drawing -ne ""){
            $drawings += ,(($drawing -replace [regex]::Escape(' '),'.' -replace [regex]::Escape('-'),'.'), $XPPS, $position)
        }
        $Material = $Material -replace [regex]::Escape('+'),'{+}' -replace [regex]::Escape('%'),'{%}' -replace [regex]::Escape('^'),'{^}' -replace [regex]::Escape('~'),'{~}' -replace [regex]::Escape('('),'{(}' -replace [regex]::Escape(')'),'{)}'
        if($Jobname -eq $null -or $Jobname -eq "" -or $Jobname -eq "?" -or $Jobname -eq "project / reference"){$jobname = $part}
        $part = $part -replace [regex]::Escape('+'),'{+}' -replace [regex]::Escape('%'),'{%}' -replace [regex]::Escape('^'),'{^}' -replace [regex]::Escape('~'),'{~}' -replace [regex]::Escape('('),'{(}' -replace [regex]::Escape(')'),'{)}'
        [void] [Tricks]::SetForegroundWindow($TMwindow)
        start-sleep -Milliseconds 600
        $operation = 10
        for($XX = 0; $xx -le 13; $XX++){$speed++
            $activitydesc = $null
            $opname = $positions[$XX][0]
            $optime = $positions[$XX][1]
            if($optime -eq 0.01){
                if($price -gt 10){
                    $optime = $price/1000
                }
                else{
                   $optime = $totalPrice/1000 
                }
            }
            $opmchn = $machine[$XX]
            $CPH = $cost1[$XX]
            $CostCenter = $cost2[$XX]
            if($QN -eq "Standard"){$optime = $optime*$Quantity}
            $activitydesc = $positions[$XX][2]

            if($XX -eq 6 -and $heatreat -ne "None"){
                TM-respond $ExcelId $PSwindow
                Write-host "$($XPPS) $("{0:D4}" -f ($position+$operation)) Heat Treat: $heatreat" -ForegroundColor DarkGreen
                [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$XPPS}else{"joe"}){tab}$($XVAS){tab}$($OrderNumber){tab}$($Customer){tab}{tab}$($OrderDate){tab}$("{0:D4}" -f $position){tab}$Part{tab}$drawing{tab}$Quantity{tab}$($EndDate){tab}Offen{tab}$("{0:D4}" -f $operation){tab}External Heat Treatment{tab}$(if($heatreat -notmatch "y"){$heatreat}){tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}$ID{tab}$WZ{tab}{tab}$stamp{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
                TM-return $speed
                $operation=$operation+10
            }
            if($xx -eq 13 -and $material -match "390" -and $part -match "die"){ #Tempering
                TM-respond $ExcelId $PSwindow
                Write-host "$($XPPS) $("{0:D4}" -f ($position+$operation)) Tempering: $part" -ForegroundColor DarkGreen
                [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$XPPS}else{"joe"}){tab}$($XVAS){tab}$($OrderNumber){tab}$($Customer){tab}{tab}$($OrderDate){tab}$("{0:D4}" -f $position){tab}$Part{tab}$drawing{tab}$Quantity{tab}$($EndDate){tab}Offen{tab}$("{0:D4}" -f $operation){tab}Heat Treat Finish{tab}Tempering{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}$ID{tab}$WZ{tab}{tab}$stamp{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
                TM-return $speed
                $operation=$operation+10
            }
            if($xx -eq 13 -and $coating -ne "None"){
                TM-respond $ExcelId $PSwindow
                Write-host "$($XPPS) $("{0:D4}" -f ($position+$operation)) $Coating" -ForegroundColor DarkGreen
                [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$XPPS}else{"joe"}){tab}$($XVAS){tab}$($OrderNumber){tab}$($Customer){tab}{tab}$($OrderDate){tab}$("{0:D4}" -f $position){tab}$Part{tab}$drawing{tab}$Quantity{tab}$($EndDate){tab}Offen{tab}$("{0:D4}" -f $operation){tab}Heat Treat Finish{tab}$coating{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}$ID{tab}$WZ{tab}{tab}$stamp{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
                TM-return $speed
                $operation=$operation+10
            }
            if($xx -eq 13 -and $nitriding -ne "None"){
                TM-respond $ExcelId $PSwindow
                Write-host "$($XPPS) $("{0:D4}" -f ($position+$operation)) Nitriding: $nitriding" -ForegroundColor DarkGreen
                [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$XPPS}else{"joe"}){tab}$($XVAS){tab}$($OrderNumber){tab}$($Customer){tab}{tab}$($OrderDate){tab}$("{0:D4}" -f $position){tab}$Part{tab}$drawing{tab}$Quantity{tab}$($EndDate){tab}Offen{tab}$("{0:D4}" -f $operation){tab}Heat Treat Finish{tab}$Nitriding{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}$ID{tab}$WZ{tab}{tab}$stamp{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
                TM-return $speed
                $operation=$operation+10
            }
            if($xx -eq 13 -and $blackoxid -ne "None"){
                TM-respond $ExcelId $PSwindow
                Write-host "$($XPPS) $("{0:D4}" -f ($position+$operation)) Black Oxide: $part" -ForegroundColor DarkGreen
                [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$XPPS}else{"joe"}){tab}$($XVAS){tab}$($OrderNumber){tab}$($Customer){tab}{tab}$($OrderDate){tab}$("{0:D4}" -f $position){tab}$Part{tab}$drawing{tab}$Quantity{tab}$($EndDate){tab}Offen{tab}$("{0:D4}" -f $operation){tab}Heat Treat Finish{tab}$blackoxid{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}$ID{tab}$WZ{tab}{tab}$stamp{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
                TM-return $speed
                $operation=$operation+10
            }
            if($Part -match "(Spreizdorn)" -and $xx -eq 4){
                TM-respond $ExcelId $PSwindow
                Write-host "$($XPPS) $("{0:D4}" -f ($position+$operation)) Sawing: $opmchn  Time: $optime $activitydesc" -ForegroundColor DarkGreen
                [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$XPPS}else{"joe"}){tab}$($XVAS){tab}$($OrderNumber){tab}$($Customer){tab}{tab}$($OrderDate){tab}$("{0:D4}" -f $position){tab}$Part{tab}$drawing{tab}$Quantity{tab}$($EndDate){tab}Offen{tab}$("{0:D4}" -f $operation){tab}Sawing{tab}$activitydesc{tab}Raw bandsaw{tab}30.00{tab}6410255750{tab}$($Quantity*0.5){tab}{tab}{tab}{tab}$ID{tab}$WZ{tab}{tab}$stamp{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
                TM-return $speed
                $operation=$operation+10
            }
            if($optime -ne "0" -and $optime -ne "0.0" -and $optime -ne $false){
                TM-respond $ExcelId $PSwindow
                if($xx -eq 13 -and ($z[$5][7] -ne "" -or $z[$5][8] -ne "" -or $z[$5][9] -ne "")){$stamp = "$($z[$5][7])   $($z[$5][8])   $($z[$5][9])"}
                if($opname -eq "CAM"){$opname = "CAM Programming"}
                Write-host "$($XPPS) $("{0:D4}" -f ($position+$operation)) $($opname): $opmchn  Time: $optime $activitydesc" -ForegroundColor DarkGreen
                [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$XPPS}else{"joe"}){tab}$($XVAS){tab}$($OrderNumber){tab}$($Customer){tab}{tab}$($OrderDate){tab}$("{0:D4}" -f $position){tab}$Part{tab}$drawing{tab}$Quantity{tab}$($EndDate){tab}Offen{tab}$("{0:D4}" -f $operation){tab}$opname{tab}$activitydesc{tab}$opmchn{tab}$CPH{tab}$CostCenter{tab}$([math]::Round($optime,2)){tab}{tab}{tab}{tab}$ID{tab}$WZ{tab}{tab}$stamp{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
                if($xx -eq 13){$Stamp = $null}
                TM-return $speed
                $operation=$operation+10
            }
        }
        if($position -eq 100){
            #Auftrag
            TM-respond $ExcelId $PSwindow
            $stamp = "$(get-content env:username)   $QN   TM-$serial"
            Write-host "$($XPPS) 0000 Auftrag: $Jobname" -ForegroundColor DarkRed
            [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$XPPS}else{"joe"}){tab}$($XVAS){tab}$($OrderNumber){tab}$($Customer){tab}$($TotalPrice){tab}$($OrderDate){tab}{tab}$Jobname{tab}{tab}{tab}$($EndDate){tab}Offen{tab}{tab}AUFTRAG GESAMT{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}$ID{tab}$WZ{tab}{tab}$stamp{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
            $stamp = $null
            TM-return $speed
        }

        #position
        TM-respond $ExcelId $PSwindow
        Write-host "$($XPPS) $("{0:D4}" -f $position) Position: $("{0:D4}" -f $position)" -ForegroundColor DarkGreen
        [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$XPPS}else{"joe"}){tab}$($XVAS){tab}$($OrderNumber){tab}$($Customer){tab}$price{tab}$($OrderDate){tab}$("{0:D4}" -f $position){tab}$Part{tab}$drawing{tab}$Quantity{tab}$($EndDate){tab}Offen{tab}{tab}POSITION GESAMT{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}$ID{tab}$WZ{tab}{tab}$stamp{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
        TM-return $speed

        if(!$checkbox -or $checkbox[$5][6]){
            #Material
            TM-respond $ExcelId $PSwindow
            Write-host "$($XPPS) $("{0:D4}" -f ($position+5)) Material: $material" -ForegroundColor DarkGreen
            [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$XPPS}else{"joe"}){tab}$($XVAS){tab}$($OrderNumber){tab}$($Customer){tab}{tab}$($OrderDate){tab}$("{0:D4}" -f $position){tab}$Part{tab}$drawing{tab}$Quantity{tab}$($EndDate){tab}Offen{tab}0005{tab}Material{tab}$material{tab}manual workstation{tab}$CPH{tab}$CostCenter{tab}{tab}{tab}{tab}{tab}$ID{tab}$WZ{tab}{tab}$stamp{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
            TM-return $speed
        }
    }
    $5 = $XPPS = $XVAS = $OrderNumber = $Customer = $price = $OrderDate = $EndDate = $TotalPrice = $jobname = $null
    [void] [Tricks]::SetForegroundWindow($PSwindow)
}
Stop-Process -Id $ExcelId -Force -ErrorAction SilentlyContinue
if($QN -ne "Custom"){Stop-Process -Id $ExcelId2 -Force -ErrorAction SilentlyContinue}




Write-Host "        [Diagnostics] Bookmark count: $($bookmark.count)" -ForegroundColor Black
$Print = $null
if($QN -ne "Custom"){$Print = read-host "Print? (y/n)"}else{$print = $checkbox[0][5]}
start-sleep -m 500
if($Print -eq "y" -or $print -eq $true){
    if($quote -ne $null){$QN = $quote}
    get-process TM_schrank -ErrorAction SilentlyContinue | stop-process -ErrorAction SilentlyContinue
    for($c = 1; $c -le 1; $c++){
        ii "C:\Program Files (x86)\TMSchrank\TM_Schrank.exe"
        start-sleep -s 3
        $TMwindow = (Get-Process tm_schrank).MainWindowHandle
        [void] [Tricks]::SetForegroundWindow($TMwindow)

        add-type -AssemblyName microsoft.VisualBasic
        add-type -AssemblyName System.Windows.Forms
        start-sleep -Milliseconds 500

        [System.Windows.Forms.SendKeys]::SendWait("~~092214~{tab}{tab}~{tab}~{tab}{tab}") # Standard quick log in (as toolshop intern)
        try{get-process tm_schrank | out-null;Write-host "TM Running" -ForegroundColor DarkGreen}catch{$c = 0,"TM Failed"}
    }
    $action = "Printing"
    for($5 = 0; $5 -lt $bookmark.count; $5++){
        Write-Progress -activity “$($bookmark[$5][0])” -id 43 -PercentComplete ((($5+1) / $bookmark.count)*100)
        [System.Windows.Forms.SendKeys]::SendWait("$(if($safemode -eq "n"){$bookmark[$5][0]}else{"joe"})~+({tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab})")
        for($b = 100; $b -le $bookmark[$5][1]; $b = $b + 100){
            TM-respond $ExcelId $PSwindow
            if($b -eq 100 -and $($bookmark[$5][1]) -gt 100){
                Write-Progress -activity “Printing” -id 44 -PercentComplete (($b / ($bookmark[$5][1]*2))*100) -status "$("{0:D4}" -f $b)/$($bookmark[$5][1])" -CurrentOperation "Cover Sheet"
                Write-Host "Printing: $($bookmark[$5][0]) Cover Sheet " -ForegroundColor DarkGreen
                start-sleep -s 1
                [System.Windows.Forms.SendKeys]::SendWait("~")
                start-sleep -s 1
                [System.Windows.Forms.SendKeys]::SendWait("~")
                start-sleep -s 2
                [System.Windows.Forms.SendKeys]::SendWait("5~")
                start-sleep -s 2
                [System.Windows.Forms.SendKeys]::SendWait("~")
                start-sleep -s 1
            }
            Write-Progress -activity “Printing” -id 44 -PercentComplete ((($b-50) / ($bookmark[$5][1]))*100) -status "$("{0:D4}" -f $b)/$($bookmark[$5][1])" -CurrentOperation "Job Card"
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
                Write-Progress -activity “Print” -id 44 -PercentComplete (($b / $bookmark[$5][1])*100) -status "$("{0:D4}" -f $b)/$($bookmark[$5][1])" -CurrentOperation "Drawing"
                if($item[1] -eq $($bookmark[$5][0]) -and $QN -ne "Standard"){
                    if($item[2] -eq $b){
                        try{
                            get-childitem "$ToolShop\05. Share\11. Tool Shop Quotes\01_New-parts\$QuoteFolder\$QNx*\" -recurse -EA stop -Exclude "*quote*" | where {$_.extension -eq ".pdf" -and $_.BaseName -match $item[0]} | % {
                                if(($_.length -gt 89000 -and $_.BaseName -match "0881") -or ($_.length -gt 66000 -and $_.BaseName -match "900") -or ($_.length -gt 26000 -and $_.BaseName -match "6.645") -or ($_.length -gt 34000 -and $_.BaseName -match "6.820")){
                                    11x17
                                    Write-host "Printing: [11x17] $($_.BaseName)" -ForegroundColor DarkGreen
                                    Start-Process -FilePath $_.FullName –Verb Print -PassThru | %{sleep 10;$_} | kill
                                    letter
                                }else{
                                    Write-host "Printing: $($_.BaseName)" -ForegroundColor DarkGreen
                                    Start-Process -FilePath $_.FullName –Verb Print -PassThru | %{sleep 10;$_} | kill
                                }
                            }
                        }catch{<#
                            try{
                            get-childitem "$ToolShop\05. Share\11. Tool Shop Quotes" -recurse -EA stop -Exclude "*quote*" | where {$_.extension -eq ".pdf" -and $_.BaseName -match $item[0]} | % {
                                if(($_.length -gt 89000 -and $_.BaseName -match "0881") -or ($_.length -gt 66000 -and $_.BaseName -match "900") -or ($_.length -gt 26000 -and $_.BaseName -match "6.645") -or ($_.length -gt 34000 -and $_.BaseName -match "6.820")){
                                    11x17
                                    Write-host "Printing: [11x17] $($_.BaseName)" -ForegroundColor DarkGreen
                                    Start-Process -FilePath $_.FullName –Verb Print -PassThru | %{sleep 10;$_} | kill
                                    letter
                                }else{
                                    Write-host "Printing: $($_.BaseName)" -ForegroundColor DarkGreen
                                    Start-Process -FilePath $_.FullName –Verb Print -PassThru | %{sleep 10;$_} | kill
                                }
                                continue
                            }
                            }catch{
                            get-childitem "$ToolShop\04. Blue Prints" -recurse -EA SilentlyContinue -Exclude "*quote*" | where {$_.extension -eq ".pdf" -and $_.BaseName -match $item[0]} | % {
                                if(($_.length -gt 89000 -and $_.BaseName -match "0881") -or ($_.length -gt 66000 -and $_.BaseName -match "900") -or ($_.length -gt 26000 -and $_.BaseName -match "6.645") -or ($_.length -gt 34000 -and $_.BaseName -match "6.820")){
                                    11x17
                                    Write-host "Printing: 11x17 $($_.BaseName)" -ForegroundColor DarkGreen
                                    Start-Process -FilePath $_.FullName –Verb Print -PassThru | %{sleep 10;$_} | kill
                                    letter
                                }else{
                                    Write-host "Printing: $($_.BaseName)" -ForegroundColor DarkGreen
                                    Start-Process -FilePath $_.FullName –Verb Print -PassThru | %{sleep 10;$_} | kill
                                }
                                continue
                            }
                            }#>
                        }
                    }
                }
            }
        }
        $bookmark[$5][2] = $true
        [System.Windows.Forms.SendKeys]::SendWait("{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}")
    }
    Write-Progress -activity “Print” -id 44 -status "x" -Completed
    Write-progress -activity "x” -id 43 -status "x"  -Completed
    
    <#foreach ($item in $drawings){
        get-childitem "$ToolShop\05. Share\11. Tool Shop Quotes\01_New-parts\$QuoteFolder\$QN*\" -recurse -Exclude "*quote*" | where {$_.extension -eq ".pdf" -and $_.BaseName -match $item} | % {
            Start-Process -FilePath $_.FullName –Verb Print -PassThru #| %{sleep 5;$_} | kill
        }
    }#>
}


#Print check
if($Print -eq "y"){
    TM-respond $ExcelId $PSwindow
    if($bookmark -contains $false){Write-Host "`nSome items failed to print!" -ForegroundColor Yellow}
    else{Write-host "`nPrint Operation Complete"}
}


<#
if(($Print -eq "y" -or $print -eq $true) -and $safemode -eq "n"){
Write-Host "Updating Guehring download...`nPlease wait window is NOT frozen, just working..."
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
[System.Windows.Forms.SendKeys]::SendWait("~")
Sleep 5
[System.Windows.Forms.SendKeys]::SendWait("+({tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab})~1~")
Start-Sleep -Milliseconds 500
[System.Windows.Forms.SendKeys]::SendWait("$ToolShop\05. Share\$('$ToolShopReportPkg')\$('$Downloads')\Orders_Guehring\Guehring.csv{tab}c{tab}~~")
sleep 5
get-process TM_schrank -ErrorAction SilentlyContinue | stop-process -ErrorAction SilentlyContinue
}
#>



[void] [Tricks]::SetForegroundWindow($PSwindow)
Read-host "Press Enter to Exit"
stop-transcript


Add-Content -Path $log -Value '[Diagnostics] $z Part info'
Add-Content -Path $log -Value $z
if($checkbox){
    Add-Content -Path $log -Value '[Diagnostics] $checkbox info'
    Add-Content -Path $log -Value $checkbox
}
Add-Content -Path $log -Value '[Diagnostics] $bookmark'
Add-Content -Path $log -Value $bookmark
Add-Content -Path $log -Value '[Diagnostics] $drawing'
Add-Content -Path $log -Value $drawings
Add-Content -Path $log -Value '[Diagnostics] Errors:'
Add-Content -Path $log -Value $Error







<#

CHANGE LOG

8/21/15

-Added change log

-Machine times load from separate changable file



8/27/15

-FIXED Compatability with quote version 5.2


9/24/15

-Change QHC to Standard


10/01/15

-Triple confirm on TM-return (fixes date error)

-De-nulled XPPS and added condition to prevent for checkbox


10/8/15

-Added progress bar to print

-fixed missing quote sheet issue


10/15/15

-Added contact to additional info


10/22/15

-fixed translations on checkbox



#>