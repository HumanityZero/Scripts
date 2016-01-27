Write-Host "Opening in read-only mode..."
$ToolShop = '\\10.51.0.12\DFSRoots\Tool Shop\'

$trackingsheet = get-item "$ToolShop\05. Share\$('$ToolShopReportPkg')\Orders\2015\ToolShopOrders_*.xlsm" | sort LastWriteTime | Select -last 1
$trackingsheet
attrib +r $trackingsheet
$sheetname = "active"
$xl = New-Object -ComObject Excel.Application
$xl.Visible=$true
$wb = $xl.Workbooks.Open($trackingsheet)
$ws = $wb.Worksheets.Item($sheetName)
$xl.Activewindow.Zoom = 90