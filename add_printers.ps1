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

# OUR LOCALLY INSTALLED DRIVER(S)
$BroDriver = 'Brother Generic Jpeg Type2 Class Driver'

# INITIALIZE/SET PRINTER VALUES ###########
$RedPrinter = [PrinterTempl]::new(); $RedPrinter.PrinterName = 'PT-Red_Printer'; $RedPrinter.PrinterPort = 'PT-Red_Printer'; $RedPrinter.PrinterIP = '192.168.26.XXX'; $RedPrinter.PrinterDriver = $BroDriver
$BluePrinter = [PrinterTempl]::new(); $BluePrinter.PrinterName = 'PT-Blue_Printer'; $BluePrinter.PrinterPort = ''; $BluePrinter.PrinterIP = '192.168.26.XXX'; $BluePrinter.PrinterDriver = $BroDriver
$OrangePrinter = [PrinterTempl]::new(); $OrangePrinter.PrinterName = 'PT-Orange_Printer'; $OrangePrinter.PrinterPort = ''; $OrangePrinter.PrinterIP = '192.168.26.XXX'; $OrangePrinter.PrinterDriver = $BroDriver
$YellowPrinter = [PrinterTempl]::new();$Yellowrinter.PrinterName = 'PT-Yellow_Printer'; $YellowPrinter.PrinterPort = ''; $YellowPrinter.PrinterIP = '192.168.26.XXX'; $YellowPrinter.PrinterDriver = $BroDriver
$RainbowPrinter = [PrinterTempl]::new(); $RainbowPrinter.PrinterName = 'PT-Rainbow_Printer'; $RainbowPrinter.PrinterPort = ''; $RainbowPrinter.PrinterIP = '192.168.26.XXX'; $RainbowPrinter.PrinterDriver = $BroDriver
$BlackPrinter = [PrinterTempl]::new(); $BlackPrinter.PrinterName = 'PT-Black_Printer'; $BlackPrinter.PrinterPort = ''; $BlackPrinter.PrinterIP = '192.168.26.XXX'; $BlackPrinter.PrinterDriver = $BroDriver
$PurplePrinter = [PrinterTempl]::new(); $PurplePrinter.PrinterName = 'PT-Purple_Printer'; $PurplePrinter.PrinterPort = ''; $PurplePrinter.PrinterIP = '192.168.26.XXX'; $PurplePrinter.PrinterDriver = $BroDriver
$GrayPrinter = [PrinterTempl]::new(); $GrayPrinter.PrinterName = 'PT-Gray_Printer'; $GrayPrinter.PrinterPort = ''; $GrayPrinter.PrinterIP = '192.168.26.XXX'; $GrayPrinter.PrinterDriver = $BroDriver

# Place All Initilized Printer Instanced objects into an Array List
$PrinterArray = $RedPrinter, $BluePrinter, $OrangePrinter, $YellowPrinter, $RainbowPrinter, $BlackPrinter, $PurplePrinter, $GrayPrinter

# ITERATE THROUGH OUR PRINTER ARRAY TO ADD PRINTERS ON LOCAL MACHINE
for ($num = 0, $num -le 8 ; $num++) {
Add-PrinterPort -Name $PrinterArray[$num].PrinterName -PrinterHostAddress $PrinterArray[$num].PrinterIP && Start-Sleep 1
Add-PrinterDriver -Name $PrinterArray[$num].PrinterDriver && Start-Sleep 1
Add-Printer -Name $PrinterArray[$num].PrinterName -DriverName $PrinterArray[$num].PrinterDriver -Port $PrinterArray[$num].PrinterPort
}


