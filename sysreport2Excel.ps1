$ErrorActionPreference = “SilentlyContinue"

#obj constructor
class Seven24_obj {

    [string]$MACaddress
    [string]$IPaddress
    [string]$HostName
    [string]$Location
    [string]$DiskFree
    [string]$DiskMax
    [string]$LastHotfix
    [string]$hotfixDate

}

#intilize 724 objects
$EXAMPLE_OBJ = [Seven24_obj]::new(); $EXAMPLE_724.MACaddress = ''; $EXAMPLE_724.IPaddress = ''; $EXAMPLE_724.Hostname = ''; $EXAMPLE_724.Location = ''; $EXAMPLE_724.DiskFree = ''; $EXAMPLE_724.DiskMax = ''; $EXAMPLE_724.lastHotFix = ''; $EXAMPLE_724.hotfixdate = ''

#Array of 724 Objects
$724_Array = $EXAMPLE1, $EXAMPLE2

#email variables initilize
$emailSmtpServer = ""
$emailFrom = ""
$emailDistro = ""
$emailAlertsDistro = "", ""
$date = Get-Date -Format FileDate
$time_now = Get-Date -format "MM.dd.yyyy @ HH:mm"

#Excel Workbook initilize
$724_Table = New-Object -COMobject Excel.Application
$724_Table.visible = $False
$724_init = $724_Table.Workbooks.Add()

#add top row of excel (headers)
$intRow = 1

$724_Table.Cells.Item($intRow, 1).ColumnWidth = 18;
$724_Table.Cells.Item($intRow, 1) = "HOSTNAME"
$724_Table.Cells.Item($intRow, 2).ColumnWidth = 19;
$724_Table.Cells.Item($intRow, 2) = "IP ADDRESS"
$724_Table.Cells.Item($intRow, 3).ColumnWidth = 22;
$724_Table.Cells.Item($intRow, 3) = "MAC ADDRESS"
$724_Table.Cells.Item($intRow, 4).ColumnWidth = 13;
$724_Table.Cells.Item($intRow, 4) = "STATUS"
$724_Table.Cells.Item($intRow, 5).ColumnWidth = 18;
$724_Table.Cells.Item($intRow, 5) = "LOCATION"
$724_Table.Cells.Item($intRow, 6).ColumnWidth = 16;
$724_Table.Cells.Item($intRow, 6) = "DISK FREE"
$724_Table.Cells.Item($intRow, 7).ColumnWidth = 16;
$724_Table.Cells.Item($intRow, 7) = "DISK MAX"
$724_Table.Cells.Item($intRow, 8).ColumnWidth = 16;
$724_Table.Cells.Item($intRow, 8) = "HOTFIX"
$724_Table.Cells.Item($intRow, 9).ColumnWidth = 19;
$724_Table.Cells.Item($intRow, 9) = "HOTFIX DATE"

$intRow = $intRow + 1

#top row formatting (excel)
$724_Table.Rows.Range("A1:I1").RowHeight = 25;
$724_Table.Rows.Range("A1:I1").Font.Size = 18;
$724_Table.Rows.Range("A1:I1").Interior.ColorIndex = 23;
$724_Table.Rows.Range("A1:I1").Font.Bold = $true;

#color formatting - alternating colors
$724_Table.Rows.Range("A2:I2").Interior.ColorIndex = 37;
$724_Table.Rows.Range("A3:I3").Interior.ColorIndex = 34;
$724_Table.Rows.Range("A4:I4").Interior.ColorIndex = 2;
$724_Table.Rows.Range("A5:I5").Interior.ColorIndex = 37;
$724_Table.Rows.Range("A6:I6").Interior.ColorIndex = 34;
$724_Table.Rows.Range("A7:I7").Interior.ColorIndex = 2;
$724_Table.Rows.Range("A8:I8").Interior.ColorIndex = 37;
$724_Table.Rows.Range("A9:I9").Interior.ColorIndex = 34;
$724_Table.Rows.Range("A10:I10").Interior.ColorIndex = 2;
$724_Table.Rows.Range("A11:I11").Interior.ColorIndex = 37;
$724_Table.Rows.Range("A12:I12").Interior.ColorIndex = 34;

#horizontal formatting
$724_Table.Cells.Range("A1:I12").HorizontalAlignment = -4108;
$724_Table.Cells.item(15, 1).HorizontalAlignment = -4108;
$724_Table.Cells.item(14, 1).HorizontalAlignment = -4108;

#'checked @' blurb
$724_Table.Cells.Item(14, 1) = "Checked @"
$724_Table.Cells.Item(15, 1) = $time_now
$724_Table.Cells.item(14, 1).Font.Bold = $true;

#netOps blurb
$724_Table.Cells.Item(14, 3) = ""
$724_Table.Cells.Item(15, 3) = ""
$724_Table.Cells.Item(16, 3) = ""

#log$ blurb
$724_Table.Cells.item(14, 6).HorizontalAlignment = -4108;
$724_Table.Cells.item(15, 6).HorizontalAlignment = -4108;
$724_Table.Cells.Item(14, 6).Font.Bold = $true;
$724_math = (Get-ChildItem '' -Recurse | measure length -sum).Sum /1gb
$724_Table.Cells.Item(14, 6).Font.Size = 13;
$724_Table.Cells.Item(15, 6).Font.Size = 13;
$724_Table.Cells.Item(14, 6) = "Log$ Size(GB):"
$724_Table.Cells.Item(15, 6)= [math]::Truncate($724_math)

#netOps formatting
$724_Table.Cells.Item(14, 3).Font.Bold = $true;
$724_Table.Cells.Item(15, 3).Font.Italic = $true;
$724_Table.Cells.Item(14, 3).RowHeight = 25;
$724_Table.Cells.Item(14, 3).Font.Size = 16;

#table border formatting
$724_Table.Range("A2:I12").BorderAround(1,4,1)
$724_Table.Range("D2:D12").BorderAround(1,4,1)
$724_Table.Range("F2:F12").BorderAround(1,4,1)

#misc formatting
$724_Table.Cells.Range("A2:A12").Font.Bold = $true;
$724_Table.Rows.Range("A2:E12").RowHeight = 22;

###################################

#FOR EACH LOOP TO ADD ITEMS TO EXCEL
for ($num = 0; $num -le $724_Array.Length-1; $num++) {

if (Test-Connection -ComputerName $724_Array[$num].HostName -Count 1 -ErrorAction SilentlyContinue) { 
    
    #win-rm activate
    $ComputerName = $724_Array[$num].HostName
    $args = "SC \\$ComputerName\ Start WinRM" 
    & cmd /c $args

#disk ninja jujitsu
$drive_max = (Get-WmiObject Win32_LogicalDisk -ComputerName $ComputerName -Filter DriveType=3 | Select-Object DeviceID, @{'Name'='Size (GB)'; 'Expression'={[string]::Format('{0:N0}',[math]::truncate($_.size / 1GB))}}, @{'Name'='Freespace (GB)'; 'Expression'={[string]::Format('{0:N0}',[math]::truncate($_.freespace / 1GB))}}).'Size (GB)' 
$freespace = (Get-WmiObject Win32_LogicalDisk -ComputerName $ComputerName -Filter DriveType=3 | Select-Object DeviceID, @{'Name'='Size (GB)'; 'Expression'={[math]::truncate($_.size / 1GB)}}, @{'Name'='Freespace (GB)'; 'Expression'={[math]::truncate($_.freespace / 1GB)}}).'Freespace (GB)'

#kb hotfixes
$lasthotfix = (Get-HotFix -ComputerName $ComputerName | sort -descending | select -first 1).HotFixID
$hotfixDate = (Get-HotFix -ComputerName $ComputerName | sort -descending | select -first 1).InstalledOn.toString('MM/dd/yyyy')

    $724_Table.Cells.Item($intRow, 1) = $724_Array[$num].HostName
    $724_Table.Cells.Item($intRow, 2) = $724_Array[$num].IPaddress
    $724_Table.Cells.Item($intRow, 3) = $724_Array[$num].MACaddress
    $724_Table.Cells.Item($intRow, 4) = "UP"
    $724_Table.Cells.Item($intRow, 5) = $724_Array[$num].Location
    $724_Table.Cells.Item($intRow, 6) = $freespace
    $724_Table.Cells.Item($intRow, 7) = $drive_max
    $724_Table.Cells.Item($intRow, 8) = $lasthotfix
    $724_Table.Cells.Item($intRow, 9) = $hotfixdate

    #formatting
    $724_Table.Cells.Item($intRow, 4).Font.ColorIndex = 16;
    $724_Table.Cells.Item($intRow, 4).Interior.ColorIndex = 4;
    $724_Table.Cells.Item($intRow, 4).Font.Bold = $true;

    #disk size check (less than 50GB free)
    if($freespace -lt 50) {
    $724_Table.Cells.Item($intRow, 6).Font.Bold = $true;
    $724_Table.Cells.Item($intRow, 6).Interior.ColorIndex = 3;
    }

    #more than 50
    if($freespace -gt 50) {
    $724_Table.Cells.Item($intRow, 6).Interior.ColorIndex = 45;
    }

    #more than 110
    if($freespace -gt 110) {
    $724_Table.Cells.Item($intRow, 6).Interior.ColorIndex = 4;
    }

    $intRow = $intRow + 1

}else {    

    $724_Table.Cells.Item($intRow, 1) = $724_Array[$num].HostName
    $724_Table.Cells.Item($intRow, 2) = $724_Array[$num].IPaddress
    $724_Table.Cells.Item($intRow, 3) = $724_Array[$num].MACaddress
    $724_Table.Cells.Item($intRow, 4) = "DOWN"
    $724_Table.Cells.Item($intRow, 5) = $724_Array[$num].Location
    $724_Table.Cells.Item($intRow, 6) = "UNKNOWN"
    $724_Table.Cells.Item($intRow, 7) = "UNKNOWN"

    #formatting
    $724_Table.Cells.Item($intRow, 4).Font.Bold = $true;
    $724_Table.Cells.Item($intRow, 4).Font.ColorIndex = 1;
    $724_Table.Cells.Item($intRow, 4).Interior.ColorIndex = 3;

    #red formatting
    $724_Table.Range("F11:G11").Font.Bold = $true;
    $724_Table.Range("F11:G11").Interior.ColorIndex = 3;

    $intRow = $intRow + 1

    #emails to alert distro list we provided at top of script
    $emailSubject_Alert = "724 DOWN ALERT!! ->" + $724_Array[$num].HostName  + " -> LOCATION: " + $724_Array[$num].Location
    $emailBody_Alert = "724 DOWN ALERT!! -> " + $724_Array[$num].HostName + " -> LOCATION: " + $724_Array[$num].Location
    Send-MailMessage -To $emailAlertsDistro -From $emailFrom -Subject $emailSubject_Alert -Body $emailBody_Alert -BodyAsHTML -SmtpServer $emailSmtpServer

    }
}

#save our excel workbook ($724_Table) dynamically
$fileName = '724_Log-' + $date + '.xlsx'
$savePath = '' + $fileName
$724_Table.Workbooks['book1'].SaveAs($savePath)
$724_Table.Close
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($724_Table)
Stop-Process -n Excel

############################################################
#sends daily report to sysadmin group

    #finds last log file (that we just created)
    $path = ""
    $contents = Get-ChildItem -Path $path | Sort LastWriteTime | Select -Last 1

   #emails to distro list we initilized at top of script
    $emailSubject = "Daily 724 Tracking Report"
    $emailBody = "Daily 724 Tracking Report. Please See Attached. This is an automated message"
    $attachment = $path + $contents.name

    Send-MailMessage -To $emailDistro -From $emailFrom -Subject $emailSubject -Body $emailBody -BodyAsHTML -Attachments $attachment -SmtpServer $emailSmtpServer
############################################################
