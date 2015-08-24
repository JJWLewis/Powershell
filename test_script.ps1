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
$Range = $ws.range("A:B","D:E")
#$Range.Delete()  #not tried, comma should be a ; ?

#Set a colour instead of deleting for now
$Range.Interior.Color = 200







##Attempt 2

$colList = @(1,3,5,8)


for ($i = $colList.length; $i -ge 0; $i--) {
  
      $Range = $ws.Cells.Item($colList[$i], 1).EntireRow
        #$Range.Interior.Color = 100
      $Range.Delete()   
}

  