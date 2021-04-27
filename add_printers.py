class PrinterTempl:  
    def __init__(self, printer_name, printer_ip, printer_port, printer_driver):  
        self.printer_name = printer_name
        self.printer_ip = printer_ip
        self.printer_port = printer_port
        self.printer_driver = printer_driver

    def display(self):  
        print("ID: %s \nName: %s \nDriver: %s \nPort: %s" % (self.printer_ip, self.printer_name, self.printer_driver, self.printer_port))  



bro_driver = "Universal Brother Jepg Image type 2"

red_printer = PrinterTempl("PT-RedPrinter", "192.168.26.10", "PT-RedPrinter", bro_driver)  
  
# accessing display() method to print employee 1 information  
  
red_printer.display()   
