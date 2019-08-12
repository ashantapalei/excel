Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\test.xlsx")
objExcel.Application.DisplayAlerts = False
objExcel.Application.Visible = True
'objExcel.Workbooks.Add
'objExcel.Cells(1, 1).Value = "Test value"
'objExcel.Cells(1, 2).Value = "Second value"
objExcel.Cells(1, 1).Value = TextBox1.Text
objExcel.Cells(1, 2).Value = TextBox2.Text


'objExcel.ActiveWorkbook.Save "C:\test.xlsx"
objExcel.ActiveWorkbook.SaveAs "C:\test.xlsx"
objExcel.ActiveWorkbook.Close
'objExcel.ActiveWorkbook.


'objExcel.Application.Quit
'WScript.Echo "Finished."
'WScript.Quit