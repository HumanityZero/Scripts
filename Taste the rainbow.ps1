#Constantly test path for connection


$colors = @("Black","Blue","Cyan","DarkBlue","DarkCyan","DarkGray","DarkGreen","DarkMagenta","DarkRed","DarkYellow","Gray","Green","Magenta","Red","White","Yellow")
for($y=0; $y -lt 5000){
    for($u = 0; $u -le 15; $u++){
        $test = Test-Path '\\10.51.0.12\Tool Shop\05. Share\$ToolShopReportPkg\ToolShopOrders.xlsm'
        write-host "$($y) $test" -ForegroundColor $colors[$u]
        start-sleep -m 20
        $y++
    }
}
