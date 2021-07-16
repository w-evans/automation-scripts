#OUR DEVICE ARRAY
$CompArray = 'adadNB0220VJJ', 'afsdafNB0220VQD', 'RgsgWK93524SF', 'RKasdWK93524QR', 'RasdFNB0220VGW', 'KsdMFWK9352RP', '14NB020VBK', '23FWK7312J97', 

#Excel Document Params (do not edit)

$a = New-Object -COMobject Excel.Application
$a.visible = $True  
$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)
$d = $c.UsedRange
$a.Columns.item(1).ColumnWidth = 30; 

#add date/time to top row
$intRow = 1
$date = Get-Date
$c.Cells.Item($intRow, 1) = $date
$intRow = $intRow + 1
###################################

#FOR EACH LOOP
for ($num = 0; $num -le $CompArray.Length-1; $num++) {

if (Test-Connection -ComputerName $CompArray[$num] -Count 1 -ErrorAction SilentlyContinue) { 
    
    $c.Cells.Item($intRow, 1).Font.ColorIndex = 10;
    $c.Cells.Item($intRow, 1) = $CompArray[$num] + ' is up'
    $intRow = $intRow + 1

}else {
    
    $c.Cells.Item($intRow, 1).Font.ColorIndex = 3;
    $c.Cells.Item($intRow, 1).Font.Bold = $true;
    $c.Cells.Item($intRow, 1) = $CompArray[$num] + ' is down'
    $intRow = $intRow + 1

    }
}
