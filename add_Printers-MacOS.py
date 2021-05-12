#!/usr/bin/python
import os
import time
#Will's fancy printer-add script
#Copyright 2021 PrimeTrust LLC

#Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without 
#limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
#The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
#THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
#AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

#PRINTER OBJECT CONSTRUCTOR
class PrinterTempl:  
    def __init__(self, printer_name, printer_ip, printer_port, printer_driver):  
        self.printer_name = printer_name
        self.printer_ip = printer_ip
        self.printer_port = printer_port
        self.printer_driver = printer_driver

#LOCALLY INSTALLED DRIVERS
bro_driver = "Universal Brother Jepg Image type 2"
hp_drivers = "HP Universal PCL 6 Driver"

#CREATING OUR PRINTERTEMPL OBJECTS
red_printer = PrinterTempl("PT-RedPrinter", "192.168.100.8", "PT-RedPrinter", bro_driver)
blue_printer = PrinterTempl("PT-BluePrinter", "192.168.100.23", "PT-BluePrinter", bro_driver) 
orange_printer = PrinterTempl("PT-OrangePrinter", "192.168.100.28", "PT-OrangePrinter", bro_driver)  
rainbow_printer = PrinterTempl("PT-RainbowPrinter", "192.168.26.13", "PT-RainbowPrinter", bro_driver)  
black_printer = PrinterTempl("PT-BlackPrinter", "192.168.26.14", "PT-BlackPrinter", bro_driver)  
purple_printer = PrinterTempl("PT-PurplePrinter", "192.168.26.15", "PT-PurplePrinter", bro_driver)  
gray_printer = PrinterTempl("PT-GrayPrinter", "192.168.26.16", "PT-GrayPrinter", bro_driver)  
yellow_printer = PrinterTempl("PT-YellowPrinter", "192.168.26.17", "PT-YellowPrinter", bro_driver)  

#PLACE ALL OUR PRINTER OBJECTS INTO AN ARRAY
printer_array = [red_printer, blue_printer, orange_printer, rainbow_printer, black_printer, purple_printer, gray_printer, yellow_printer]

#ITERATE THROUGH OUR PRINTER ARRAY AND ADD THE PRINTERS TO THE LOCAL MACHINE
for i in printer_array:
    os.system("lpadmin -p " + i.printer_name + " -D " + i.printer_port + " -E -v" + " ipp://" + i.printer_ip + "/ipp/print" + " -m everywhere")
    time.sleep(0.5)
