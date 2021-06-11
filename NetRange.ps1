#Excel Document Params (do not edit)

$erroractionpreference = "SilentlyContinue"
$a = New-Object -comobject Excel.Application
$a.visible = $True 

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "MAC"
$c.Cells.Item(1,2) = "IP"
$c.Cells.Item(1,3) = "Device"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True
$d.EntireColumn.AutoFit($True)

$intRow = 1

#Define and filter

$netcol = get-netneighbor | where ipaddress -like "138.136.*.*"

foreach($PC in $NetCol){
if(test-connection -ComputerName $PC.IPAddress -Count 1 -Quiet)
{
$mac = $PC.linklayeraddress 
$IP =  $PC.ipaddress
$device = (get-wmiobject -class win32_computersystem -ComputerName $pc.IPAddress).name

if($device -eq $null){
$device = "No Device Name"}

$intRow = $intRow + 1
$c.Cells.Item($intRow, 1) = $mac.ToUpper()
$c.Cells.Item($intRow, 2) = $IP
$c.Cells.Item($intRow, 3) = $device
}
}

$d.EntireColumn.AutoFit()


