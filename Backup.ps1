#shortcutinfo
#C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File "C:\Users\bredestegej\Documents\backups\Rename_Script.ps1" -WindowsStyle Hidden
#to enable call joe: 513-373-9531

$serial=$null

#serial Generator
for($i=1; $i -le 10; $i++){

$safe=(26/36)*100
$randomfyer=get-random -min 0 -max 99
if($randomfyer -le $safe){
$a=Get-Random -Count 1 -InputObject (65..90) | % -begin {$aa=$null} -process {$aa += [char]$_} -end {$aa}}
else{$a=get-random -max 9 -min 0}

#write-host "$a" -nonewline
$serial="$serial$a"}

#Renaming target:
$directory = "A:\04. Blue Prints"

##change directory
set-location C:\BU

#Various variables:
$blanks=0
$quiet="/quiet" # want to run quiet? (clear to run loud)
$namelog="B$serial.log"

. "C:\BU\log.ps1"
Log-Start -LogPath "C:\Windows\Temp" -logname "$namelog" -ScriptVersion "1.0" -SessionID "$serial"
Log-Write -LogPath "C:\Windows\Temp\$namelog" -LineValue "[$(get-date -f hh:mm:ss)] Program initiated"
clear-host
Write-host "Server Backup"
Write-host "Version 1.0"
Write-host "Session ID: $serial"
Write-host "Compiled by Joe Bredestege"
write-host "Creation date: 1/24/15"
start-sleep -s 10


#error action (change to "silentlycontinue" to ignore errors, inquire to ask before proceeding)
$ErrorActionPreference = "stop"
Log-Write -LogPath "C:\Windows\Temp\$namelog" -LineValue "[$(get-date -f hh:mm:ss)] Error action: $erroractionpreference"

# chcp 1252

Write-Host "`nBacking up - This will take a few minutes..."
start-sleep -s 4
#copy directory
Log-Write -LogPath "C:\Windows\Temp\$namelog" -LineValue "[$(get-date -f hh:mm:ss)] Starting Backup"
#[reflection.assembly]::loadfile("C:\BU\LongPath\Bin\Microsoft.Experimental.IO.dll")
#[microsoft.experimental.io.longpathfile]::Copy((gi $directory).fullname, "D:BP\$(get-date -f dd.MM.yy)",$true)
#Copy-Item -path "$directory" -destination D:BP\$(get-date -f dd.MM.yy) -recurse -force
Log-Write -LogPath "C:\Windows\Temp\$namelog" -LineValue "[$(get-date -f hh:mm:ss)] Defaulting to Robocopy backup - Initiated"
robocopy.exe $directory D:BP\$(get-date -f dd.MM.yy) "*.*" /e 
Log-Write -LogPath "C:\Windows\Temp\$namelog" -LineValue "[$(get-date -f hh:mm:ss)] Robocopy complete"

Write-Progress -Activity "Backing up" -completed
Write-host "Complete"
Log-Write -LogPath "C:\Windows\Temp\$namelog" -LineValue "[$(get-date -f hh:mm:ss)] Backup Complete - Progress Bar exited"
Log-Write -LogPath "C:\Windows\Temp\$namelog" -LineValue "[$(get-date -f hh:mm:ss)] Program Complete - Exiting"
Log-Finish -LogPath "C:\Windows\Temp\$namelog" -NoExit $True
exit