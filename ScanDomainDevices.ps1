#create template for our 724 objects
class Seven24_obj {
    [string]$MACaddress
    [string]$IPaddress
    [string]$HostName
    [string]$Location
}

#intilize 724 objects
$PACU_724 = [Seven24_obj]::new(); $PACU_724.MACaddress = ''; $PACU_724.IPaddress = ''; $PACU_724.Hostname ='; $PACU_724.Location = 'PACU'
$Admin_724 = [Seven24_obj]::new(); $Admin_724.MACaddress = ''; $Admin_724.IPaddress = ''; $Admin_724.Hostname = ''; $Admin_724.Location = 'Admin (Kevins desk)'
$OR_724 = [Seven24_obj]::new(); $OR_724.MACaddress = ''; $OR_724.IPaddress = ''; $OR_724.Hostname = ''; $OR_724.Location = 'OR'
$MSU_724 = [Seven24_obj]::new(); $MSU_724.MACaddress = ''; $MSU_724.IPaddress = ''; $MSU_724.Hostname = ''; $MSU_724.Location = 'MSU'
$Lab_724 = [Seven24_obj]::new(); $Lab_724.MACaddress = ''; $Lab_724.IPaddress = ''; $Lab_724.Hostname = ''; $Lab_724.Location = 'Lab'
$BloodBank_724 = [Seven24_obj]::new(); $BloodBank_724.MACaddress = ''; $BloodBank_724.IPaddress = ''; $BloodBank_724.Hostname = ''; $BloodBank_724.Location = 'Blood Bank'
$LD_724 = [Seven24_obj]::new(); $LD_724.MACaddress = ''; $LD_724.IPaddress = '';$LD_724.Hostname = ''; $LD_724.Location = 'LD'
$Pharm_724 = [Seven24_obj]::new(); $Pharm_724.MACaddress = ''; $Pharm_724.IPaddress = ''; $Pharm_724.Hostname = ''; $Pharm_724.Location = 'Pharmacy'
$CCU_724 = [Seven24_obj]::new(); $CCU_724.MACaddress = ''; $CCU_724.IPaddress = ''; $CCU_724.Hostname = ''; $CCU_724.Location = 'CCU'
$MicroLab_724 = [Seven24_obj]::new(); $MicroLab_724.MACaddress = ''; $MicroLab_724.IPaddress = ''; $MicroLab_724.Hostname = ''; $MicroLab_724.Location = 'Micro Lab'
$ER_724 = [Seven24_obj]::new(); $ER_724.MACaddress = ''; $ER_724.IPaddress = ''; $ER_724.Hostname = ''; $ER_724.Location = 'ER'

#Array of 724 Objects
$724_Array = $PACU_724, $Admin_724, $OR_724, $MSU_724, $Lab_724, $BloodBank_724, $LD_724, $CCU_724, $Pharm_724, $MicroLab_724, $ER_724

#Excel Document Params (do not edit)
$a = New-Object -COMobject Excel.Application
$a.visible = $False
$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)
$d = $c.UsedRange
$a.Columns.item(1).ColumnWidth = 20;
$a.Columns.item(2).ColumnWidth = 17;
$a.Columns.item(3).ColumnWidth = 20;
$a.Columns.item(4).ColumnWidth = 8;

#add top row
$intRow = 2
$date = Get-Date
$c.Cells.Item(1, 1) = $date
$c.Cells.Item($intRow, 1).Font.Bold = $true;
$c.Cells.Item($intRow, 1) = "HOSTNAME"
$c.Cells.Item($intRow, 2).Font.Bold = $true;
$c.Cells.Item($intRow, 2) = "IP ADDRESS"
$c.Cells.Item($intRow, 3).Font.Bold = $true;
$c.Cells.Item($intRow, 3) = "MAC ADDRESS"
$c.Cells.Item($intRow, 4).Font.Bold = $true;
$c.Cells.item($intRow, 4) = "STATUS"
$c.Cells.Item($intRow, 5).Font.Bold = $true;
$c.Cells.item($intRow, 5) = "LOCATION"
$intRow = $intRow + 1
###################################

#FOR EACH LOOP TO ADD ITEMS TO EXCEL
for ($num = 0; $num -le $724_Array.Length-1; $num++) {

if (Test-Connection -ComputerName $724_Array[$num].HostName -Count 1 -ErrorAction SilentlyContinue) { 
    
    $c.Cells.Item($intRow, 1) = $724_Array[$num].HostName
    $c.Cells.Item($intRow, 2) = $724_Array[$num].IPaddress
    $c.Cells.Item($intRow, 3) = $724_Array[$num].MACaddress
    $c.Cells.Item($intRow, 4).Font.ColorIndex = 10;
    $c.Cells.Item($intRow, 4) = "UP"
    $c.Cells.Item($intRow, 5) = $724_Array[$num].Location
    $intRow = $intRow + 1

}else {    

    $c.Cells.Item($intRow, 1) = $724_Array[$num].HostName
    $c.Cells.Item($intRow, 2) = $724_Array[$num].IPaddress
    $c.Cells.Item($intRow, 3) = $724_Array[$num].MACaddress
    $c.Cells.Item($intRow, 4).Font.Bold = $true;
    $c.Cells.Item($intRow, 4).Font.ColorIndex = 3;
    $c.Cells.Item($intRow, 4) = "DOWN"
    $c.Cells.Item($intRow, 5) = $724_Array[$num].Location
    $intRow = $intRow + 1


    #Send Notification that device is down
    $emailSmtpServer_NOTIFICATION = "<SMTP>"
    $emailFrom_NOTIFICATION = "<EMAIL FROM>"
    $emailTo_NOTIFICATION = "<EMAIL TO>"
    $emailSubject_NOTIFICATION = "724 DOWN NOTIFICATION ->" + $724_Array[$num].HostName  + " -> LOCATION: " + $724_Array[$num].Location
    $emailBody_NOTIFICATION = "724 DOWN NOTIFICATION -> " + $724_Array[$num].HostName + " -> LOCATION: " + $724_Array[$num].Location

    Send-MailMessage -To $emailTo_NOTIFICATION -From $emailFrom_NOTIFICATION -Subject $emailSubject_NOTIFICATION -Body $emailBody_NOTIFICATION -BodyAsHTML -SmtpServer $emailSmtpServer_NOTIFICATION

    }
}

#save the workbook
$month = (Get-Date).Month
$day = (Get-Date).Day
$year = [string](Get-Date).Year
$hour = (Get-Date).Hour
$minute = (Get-Date).Minute
$a.Workbooks['book1'].SaveAs("<PATH_HERE>\724-Log_"+ $year + "-" + $month + "-" + $day + "_" + $hour + $minute + ".xlsx")


#close excel
$a.Close
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($a)
Stop-Process -n Excel


Start-Sleep 10


#IF its 8am it will fire off an email to the Sysadmins with our daily tracking report
if ($hour -eq 8) {

    $path = "<PATH_HERE>\"
    $contents = Get-ChildItem -Path $path | Sort LastWriteTime | Select -Last 1

    $emailSmtpServer = "<SMTP>"
    $emailFrom = "<EMAIL FROM>"
    $emailTo = "<EMAIL TO>"
    $emailSubject = "Daily 724 Tracking Report"
    $attachment = "<PATH>" + $contents.name
    $emailBody = "Daily 724 Tracking Report. Please See Attached. This is an automated message"

Send-MailMessage -To $emailTo -From $emailFrom -Subject $emailSubject -Body $emailBody -BodyAsHTML -Attachments $attachment -SmtpServer $emailSmtpServer

}

Start-Sleep 5
stop-process -Id $PID
