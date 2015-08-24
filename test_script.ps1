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

#Delete a range
$Range = $ws.range("A:B")
$Range.Delete()  #not tried, comma should be a ; ?

