############### Will Evans test script for adding printers to Windows PCs################
############### 4/21/2021 - PrimeTrust LLC #####################

##### PRINTER OBJECT CONSTRUCTORS #####
class PrinterTemplate {
    [string]$PrinterName
    [string]$PrinterPort
    [string]$PrinterIP
    [string]$PrinterDriver
}

###### INITIALIZE/SET PRINTER VALUES ###########
$OrangePrinter = [PrinterTemplate]::new(); $OrangePrinter.PrinterName = 'Orange Printer'; $OrangePrinter.PrinterPort = '192.168.26.3'; $OrangePrinter.PrinterIP = '192.168.26.3'; $OrangePrinter.PrinterDriver = 'HP Universal Printing PCL 6'


#### Place All Initilized Printer Instanced objects into an Array List
$PrinterArray = $OrangePrinter


######ITERATE THROUGH ARRAY TO ADD PRINTERS ON LOCAL MACHINE (WONT NEED FOR LOOP FOR TEST)
Add-PrinterPort -Name $PrinterArray[$num].PrinterName() -PrinterHostAddress $PrinterArray[$num].PrinterIP()
Add-Printer -DriverName $PrinterArray[$num].PrinterDriver() -Name $PrinterArray[$num].PrinterName() -Port $PrinterArray[$num].PrinterIP()