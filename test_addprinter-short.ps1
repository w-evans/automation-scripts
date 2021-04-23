##### PRINTER OBJECT CONSTRUCTORS #####
class PrinterTemplate {
    [string]$PrinterName
    [string]$PrinterPort
    [string]$PrinterIP
    [string]$PrinterDriver
}

$OrangePrinter = [PrinterTemplate]::new(); $OrangePrinter.PrinterName = 'PT-Orange_Printer'; $OrangePrinter.PrinterPort = 'PT-Orange_Printer'; $OrangePrinter.PrinterIP = '192.168.26.3'; $OrangePrinter.PrinterDriver = "Brother Generic Jpeg Type2 Class Driver"

Add-PrinterPort -Name $OrangePrinter.PrinterName -PrinterHostAddress $OrangePrinter.PrinterIP
Add-PrinterDriver -Name $OrangePrinter.PrinterDriver
Start-Sleep 2
Add-Printer -Name $OrangePrinter.PrinterName -DriverName $OrangePrinter.PrinterDriver -Port $OrangePrinter.PrinterPort

