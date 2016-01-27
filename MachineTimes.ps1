$CAMProgramming = "CAM Programming"
$CAMProgrammingMachine = "CAM"
$CAMProgrammingCostPerHour = "0.00"
$CAMProgrammingCostCenter = ""

$Sawing = "Sawing"
$SawingMachine = "Raw bandsaw"
$SawingCostPerHour = "30.00"
$SawingCostCenter = "6410255750"

$StandardTurning = "Standard Turning"
$StandardTurningMachine = "manual workstation"
$StandardTurningCostPerHour = "55.00"
$StandardTurningCostCenter = "6410255770"

$CNCTurning = "CNC Turning"
$CNCTurningMachine = "CNC Lathe"
$CNCTurningCostPerHour = "85.00"
$CNCTurningCostCenter = "6410255710"

$StandardMilling = "Standard Milling"
$StandardMillingMachine = "manual workstation"
$StandardMillingCostPerHour = "55.00"
$StandardMillingCostCenter = "6410255770"

$MultiaxialMilling = "Multiaxial Milling"
$MultiaxialMillingMachine = "5-axis Mill"
$MultiaxialMillingCostPerHour = "105.00"
$MultiaxialMillingCostCenter = "6410255717"

$CircularGrinding = "Circular grinding"
$CircularGrindingMachine = "ID/OD grinder"
$CircularGrindingCostPerHour = "85.00"
$CircularGrindingCostCenter = "6410255735"

$SurfaceGrinding = "Surface Grinding"
$SurfaceGrindingMachine = "Surface grinder"
$SurfaceGrindingCostPerHour = "65.00"
$SurfaceGrindingCostCenter = "6410255730"

$WireCutEDM = "Wire Cut EDM"
$WireCutEDMMachine = "Wire cut EDM"
$WireCutEDMCostPerHour = "55.00"
$WireCutEDMCostCenter = "6410255745"

$MillingFinish = "Milling Finish"
$MillingFinishMachine = "5-axis Mill"
$MillingFinishCostPerHour = "105.00"
$MillingFinishCostCenter = "6410255717"

$TurningFinish = "Turning Finish"
$TurningFinishMachine = "CNC Lathe"
$TurningFinishCostPerHour = "85.00"
$TurningFinishCostCenter = "6410255710"

$JIGGrinding = "JIG Grinding"
$JIGGrindingMachine = "JIG Grinder"
$JIGGrindingCostPerHour = "55.00"
$JIGGrindingCostCenter = "6410255770"

$GrindingFinish = "Grinding Finish"
$GrindingFinishMachine = "Surface grinder"
$GrindingFinishCostPerHour = "65.00"
$GrindingFinishCostCenter = "6410255730"

$ManualWorkStations = "Manual Workstations"
$ManualWorkStationsMachine = "manual workstation"
$ManualWorkStationsCostPerHour = "55.00"
$ManualWorkStationsCostCenter = "6410255770"


$activity  = @(           $CAMProgramming,           $Sawing,           $StandardTurning,           $CNCTurning,           $StandardMilling,           $MultiaxialMilling,           $CircularGrinding,           $SurfaceGrinding,           $WireCutEDM,           $MillingFinish,           $TurningFinish,           $JIGGrinding,           $GrindingFinish,           $ManualWorkstations)
$machine   = @(    $CAMProgrammingMachine,    $SawingMachine,    $StandardTurningMachine,    $CNCTurningMachine,    $StandardMillingMachine,    $MultiaxialMillingMachine,    $CircularGrindingMachine,    $SurfaceGrindingMachine,    $WireCutEDMMachine,    $MillingFinishMachine,    $TurningFinishMachine,    $JIGGrindingMachine,    $GrindingFinishMachine,    $ManualWorkStationsMachine)
$cost1     = @($CAMProgrammingCostPerHour,$SawingCostPerHour,$StandardTurningCostPerHour,$CNCTurningCostPerHour,$StandardMillingCostPerHour,$MultiaxialMillingCostPerHour,$CircularGrindingCostPerHour,$SurfaceGrindingCostPerHour,$WireCutEDMCostPerHour,$MillingFinishCostPerHour,$TurningFinishCostPerHour,$JIGGrindingCostPerHour,$GrindingFinishCostPerHour,$ManualWorkStationsCostPerHour)
$cost2     = @( $CAMProgrammingCostCenter, $SawingCostCenter, $StandardTurningCostCenter, $CNCTurningCostCenter, $StandardMillingCostCenter, $MultiaxialMillingCostCenter, $CircularGrindingCostCenter, $SurfaceGrindingCostCenter, $WireCutEDMCostCenter, $MillingFinishCostCenter, $TurningFinishCostCenter, $JIGGrindingCostCenter, $GrindingFinishCostCenter, $ManualWorkStationsCostCenter)



<#  OLD WORKING NOTES
$activity = @("CAM Programming","Sawing","Standard Turning","CNC Turning","Standard Milling","Multiaxial Milling","Circular grinding","Surface Grinding","Wire Cut EDM","Milling Finish","Turning Finish","JIG Grinding","Grinding Finish","Manual Workstations")
$machine = @("CAM","Raw bandsaw","manual workstation","CNC Lathe","manual workstation","5-axis Mill","ID/OD grinder","Surface grinder","Wire cut EDM","5-axis Mill","CNC Lathe","JIG Grinder","Surface grinder","manual workstation")
$cost1 = @("0.00","30.00","55.00","85.00","55.00","105.00","85.00","65.00","55.00","105.00","85.00","55.00","65.00","55.00")
$cost2 = @("","6410255750","6410255770","6410255710","6410255770","6410255717","6410255735","6410255730","6410255745","6410255717","6410255710","6410255770","6410255730","6410255770")
#>