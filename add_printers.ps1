############### Will Evans fancy printer script ################
############### 4/21/2021 - PrimeTrust LLC #####################

##### PRINTER OBJECT CONSTRUCTORS #####
class PrinterTemplate {
    [string]$PrinterName
    [string]$PrinterPort
    [string]$PrinterIP
    [string]$PrinterDriver
}

###### INITIALIZE/SET PRINTER VALUES ###########
$RedPrinter = [PrinterTemplate]::new(); $RedPrinter.PrinterName = 'Red Printer'; $RedPrinter.PrinterPort = ''; $RedPrinter.PrinterIP = '192.168.26.XXX'; $RedPrinter.PrinterDriver = ''
$BluePrinter = [PrinterTemplate]::new(); $BluePrinter.PrinterName = 'Blue Printer'; $BluePrinter.PrinterPort = ''; $BluePrinter.PrinterIP = ''; $BluePrinter.PrinterDriver = ''
$OrangePrinter = [PrinterTemplate]::new(); $OrangePrinter.PrinterName = 'Orange Printer'; $OrangePrinter.PrinterPort = ''; $OrangePrinter.PrinterIP = ''; $OrangePrinter.PrinterDriver = ''
$YellowPrinter = [PrinterTemplate]::new();$Yellowrinter.PrinterName = 'Yellow Printer'; $YellowPrinter.PrinterPort = ''; $YellowPrinter.PrinterIP = ''; $YellowPrinter.PrinterDriver = ''
$RainbowPrinter = [PrinterTemplate]::new(); $RainbowPrinter.PrinterName = 'Rainbow Printer'; $RainbowPrinter.PrinterPort = ''; $RainbowPrinter.PrinterIP = ''; $RainbowPrinter.PrinterDriver = ''
$BlackPrinter = [PrinterTemplate]::new(); $BlackPrinter.PrinterName = 'Black Printer'; $BlackPrinter.PrinterPort = ''; $BlackPrinter.PrinterIP = ''; $BlackPrinter.PrinterDriver = ''
$PurplePrinter = [PrinterTemplate]::new(); $PurplePrinter.PrinterName = 'Purple Printer'; $PurplePrinter.PrinterPort = ''; $PurplePrinter.PrinterIP = ''; $PurplePrinter.PrinterDriver = ''
$GrayPrinter = [PrinterTemplate]::new(); $GrayPrinter.PrinterName = 'Gray Printer'; $GrayPrinter.PrinterPort = ''; $GrayPrinter.PrinterIP = ''; $GrayPrinter.PrinterDriver = ''

#### Place All Initilized Printer Instanced objects into an Array List
$PrinterArray = $RedPrinter, $BluePrinter, $OrangePrinter, $YellowPrinter, $RainbowPrinter, $BlackPrinter, $PurplePrinter, $GrayPrinter

######ITERATE THROUGH ARRAY TO ADD PRINTERS ON LOCAL MACHINE
for ($num = 0, $num -le 8 ; $num++) {

Add-PrinterPort -Name $PrinterArray[$num].PrinterName() -PrinterHostAddress $PrinterArray[$num].PrinterIP()
Add-Printer -DriverName $PrinterArray[$num].PrinterDriver() -Name $PrinterArray[$num].PrinterName() -Port $PrinterArray[$num].PrinterIP()

}


