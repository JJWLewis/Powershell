#nice article about lots of basic excel stuff
#https://technet.microsoft.com/en-us/magazine/2009.01.heyscriptingguy.aspx





#Define where the sheet is
$Path = 'C:\Users\jay.lewis\Desktop\PowerShellTest.xlsx'

#Open the Excel Document and open the sheet 'Sheet1'

#set the object called Excel
$Excel = New-Object -Comobject Excel.Application

#Show excel, set visible to false if running in background
$Excel.Visible = $true
$Excel.DisplayAlerts = $False

#Create the workbook from the defined path
$Workbook = $Excel.Workbooks.Open($Path)
#define the sheet wanted
$page = 'Sheet1'

#Set the worksheet (can also use Item(1) for sheet one)
$ws = $Workbook.WorkSheets.Item("Sheet1")
$ws.Activate()
Start-Sleep 1
$Rng = $ws.UsedRange.Cells
$row = $Rng.Rows.Count

#Delete first row in the excel sheet
[void]$ws.Cells.Item(1,1).EntireRow.Delete()

#Delete a range
$Range = $ws.range("A49:F320"), $Range.Delete()  #not tried, comma should be a ; ?




#loop through and delete rows if they are empty
$xlCellTypeLastCell = 11  #need an end, or would go all the way down
$xl = New-Object -comobject excel.application 
$xl.Visible = $true 
$xl.DisplayAlerts = $False 
$wb = $xl.Workbooks.Open("C:\Scripts\BuildXLS.xls") # <-- Change as required!
$ws = $wb.Worksheets.Item(1)
$used = $ws.usedRange 
$lastCell = $used.SpecialCells($xlCellTypeLastCell) 
$row = $lastCell.row
for ($i = 0; $i -le $row; $i++) {
    If ($ws.Cells.Item($i, 1) = " ") {
        $Range = $ws.Cells.Item($i, 1).EntireRow
        $Range.Delete()
        $i = $i - 1
    }
}






#close workbook and save changes
$Workbook.Close($true)
#quit excel
$excel.Quit()


#release the excel COM object - really important
[Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)



#very basic for loop
for ($i=1; $i -le 10; $i++) {     #le is less than <=
    Write-Host "Item" $i;         #full list of conditions and loop stuff https://blog.udemy.com/powershell-for-loops/
}


#get help info
get-help (thing name) -Full/-Showwindow

#online help
get-help (thing name) -online

