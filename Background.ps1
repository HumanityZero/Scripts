#serial Generator
$ErrorActionPreference = "silentlycontinue"
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


Start-Transcript -path "C:\Windows\Temp\Radio-$Serial.log" -append
$host.ui.RawUI.WindowTitle = "Background $serial"

Add-Type @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PInvoke.Win32 {

    public static class UserInput {

        [DllImport("user32.dll", SetLastError=false)]
        private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        [StructLayout(LayoutKind.Sequential)]
        private struct LASTINPUTINFO {
            public uint cbSize;
            public int dwTime;
        }

        public static DateTime LastInput {
            get {
                DateTime bootTime = DateTime.UtcNow.AddMilliseconds(-Environment.TickCount);
                DateTime lastInput = bootTime.AddMilliseconds(LastInputTicks);
                return lastInput;
            }
        }

        public static TimeSpan IdleTime {
            get {
                return DateTime.UtcNow.Subtract(LastInput);
            }
        }

        public static int LastInputTicks {
            get {
                LASTINPUTINFO lii = new LASTINPUTINFO();
                lii.cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO));
                GetLastInputInfo(ref lii);
                return lii.dwTime;
            }
        }
    }
}
'@

function clear-recyclebin($days) {
    $Shell = New-Object -ComObject Shell.Application
    $Global:Recycler = $Shell.NameSpace(0xa)

    foreach($item in $Recycler.Items()){
        $DeletedDate = $Recycler.GetDetailsOf($item,2) -replace "\u200f|\u200e",""
        $dtDeletedDate = get-date $DeletedDate 
        If($dtDeletedDate -lt (Get-Date).AddDays(($days)*(-1))){
            Remove-Item -Path $item.Path -Confirm:$false -Force -Recurse
        }#EndIF
    }#EndForeach item
}$days=$null

cd "C:\Users\Joe\GOOGLEDRIVE\Dropbox\Scripts\Radio recorder"

function start-npr {
    $IE=new-object -com internetexplorer.application
    $IE.navigate2("http://www.npr.org/player/v2/mediaPlayer.html?action=3&t=live1&islist=false")
    $IE.visible=$true
    $itunesid = get-process itunes | %{$_.Id} | ?{$before -notcontains $_}
    stop-process -id $itunesid
}
function stop-npr {
    Get-Process iexplore | Foreach-Object { $_.CloseMainWindow() }
}


    ii C:\Program Files\PeerBlock\peerblock.exe -CONFIRM
    ii C:\Users\Joe\AppData\Roaming\uTorrent\utorrent.exe
    if((get-date).dayofweek -eq "monday"){
        clear-recyclebin 7
    }
    if((get-date).dayofweek.value__ -lt 6){
    Sleep -S (New-TimeSpan -End "6am").TotalSeconds
        start-npr
    Sleep -S (New-TimeSpan -End "8am").TotalSeconds
        stop-npr
    Sleep -S (New-TimeSpan -End "4pm").TotalSeconds
        start-npr
    Sleep -S (New-TimeSpan -End "6pm").TotalSeconds
        stop-npr
    }Get-Process peerblock | Foreach-Object { $_.CloseMainWindow() }
    ####
    
    #timer to next iteration
    write-host (get-date)


exit

<#
$folder = "\\10.51.0.12\Tool Shop\recovery\Toolshop\Root\Script\Logs\"
$wait = 1

[void][reflection.assembly]::LoadWithPartialName("System.Windows.Forms")
$form = new-object Windows.Forms.Form
$form.Text = "Image Viewer"
$form.WindowState= "Maximized"
$form.controlbox = $false
$form.formborderstyle = "0"
$form.BackColor = [System.Drawing.Color]::black

$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.dock = "fill"
$pictureBox.sizemode = 4
$form.controls.add($pictureBox)
$form.Add_Shown( { $form.Activate()} )
$form.Show()

do
{
    $files = (get-childitem $folder | where { ! $_.PSIsContainer})
    foreach ($file in $files)
    {
        $pictureBox.Image = [System.Drawing.Image]::Fromfile($file.fullname)
        Start-Sleep -Seconds $wait
        $form.bringtofront()
    }
}
While ($running -ne 1)
#>