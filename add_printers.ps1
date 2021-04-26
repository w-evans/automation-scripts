<#
Will's fancy printer-add script
Copyright © 2021 PrimeTrust LLC

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without 
limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

# PRINTER OBJECT CONSTRUCTOR
class PrinterTempl {
    [string]$PrinterName
    [string]$PrinterPort
    [string]$PrinterIP
    [string]$PrinterDriver
}

# OUR LOCALLY INSTALLED DRIVERS
$BroDriver = 'Brother Generic Jpeg Type2 Class Driver'
$HPDriver = 'HP Universal Printing PCL 6'

# INITIALIZE/SET PRINTER VALUES ###########
$RedPrinter = [PrinterTempl]::new(); $RedPrinter.PrinterName = 'PT-Red_Printer'; $RedPrinter.PrinterPort = 'PT-Red_Printer'; $RedPrinter.PrinterIP = '192.168.26.5'; $RedPrinter.PrinterDriver = $BroDriver
$BluePrinter = [PrinterTempl]::new(); $BluePrinter.PrinterName = 'PT-Blue_Printer'; $BluePrinter.PrinterPort = 'PT-Blue_Printer'; $BluePrinter.PrinterIP = '192.168.26.6'; $BluePrinter.PrinterDriver = $BroDriver
$OrangePrinter = [PrinterTempl]::new(); $OrangePrinter.PrinterName = 'PT-Orange_Printer'; $OrangePrinter.PrinterPort = 'PT-Orange_Printer'; $OrangePrinter.PrinterIP = '192.168.26.7'; $OrangePrinter.PrinterDriver = $BroDriver
$RainbowPrinter = [PrinterTempl]::new(); $RainbowPrinter.PrinterName = 'PT-Rainbow_Printer'; $RainbowPrinter.PrinterPort = 'PT-Rainbow_Printer'; $RainbowPrinter.PrinterIP = '192.168.26.8'; $RainbowPrinter.PrinterDriver = $BroDriver
$BlackPrinter = [PrinterTempl]::new(); $BlackPrinter.PrinterName = 'PT-Black_Printer'; $BlackPrinter.PrinterPort = 'PT-Black_Printer'; $BlackPrinter.PrinterIP = '192.168.26.9'; $BlackPrinter.PrinterDriver = $BroDriver
$PurplePrinter = [PrinterTempl]::new(); $PurplePrinter.PrinterName = 'PT-Purple_Printer'; $PurplePrinter.PrinterPort = 'PT-Purple_Printer'; $PurplePrinter.PrinterIP = '192.168.26.10'; $PurplePrinter.PrinterDriver = $BroDriver
$GrayPrinter = [PrinterTempl]::new(); $GrayPrinter.PrinterName = 'PT-Gray_Printer'; $GrayPrinter.PrinterPort = 'PT-Gray_Printer'; $GrayPrinter.PrinterIP = '192.168.26.11'; $GrayPrinter.PrinterDriver = $BroDriver
$YellowPrinter = [PrinterTempl]::new(); $YellowPrinter.PrinterName = 'PT-Yellow_Printer'; $YellowPrinter.PrinterPort = 'PT-Yellow_Printer'; $YellowPrinter.PrinterIP = '192.168.26.12'; $YellowPrinter.PrinterDriver = $BroDriver

# Place All Initilized Printer Instanced objects into an Array List
$PrinterArray = $RedPrinter, $BluePrinter, $OrangePrinter, $RainbowPrinter, $BlackPrinter, $PurplePrinter, $GrayPrinter, $YellowPrinter

# ITERATE THROUGH OUR PRINTER ARRAY TO ADD PRINTERS ON LOCAL MACHINE
for ($num = 0; $num -le $PrinterArray.Length-1; $num++) {
Write-Host $PrinterArray[$num].PrinterName.toString()
Write-Host $PrinterArray[$num].PrinterPort.toString()
Write-Host $PrinterArray[$num].PrinterIP.toString()
Write-Host '================================'


###Add-PrinterPort -Name $PrinterArray[$num].PrinterName.toString() -PrinterHostAddress $PrinterArray[$num].PrinterIP.toString()
##Start-Sleep 1
##Add-PrinterDriver -Name $PrinterArray[$num].PrinterDriver.toString() 
##Start-Sleep 1
##Add-Printer -Name $PrinterArray[$num].PrinterName.toString() -DriverName $PrinterArray[$num].PrinterDriver.toString() -Port $PrinterArray[$num].PrinterPort.toString()
}


