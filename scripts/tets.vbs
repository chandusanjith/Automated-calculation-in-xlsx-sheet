Const xlToRight = -4161
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWB = objExcel.Workbooks.Open("C:\Users\chandu.s\Desktop\CARDRATE\in\card.xlsx")
Set objSheet = objwb.Sheets("WORKING")
objSheet.Columns("N:N").Insert xlToRight
objSheet.Cells(7, 14).Value = "?"
objSheet.Cells(8, 14).Value = "?"
objSheet.Cells(9, 14).Value = "?"
objSheet.Cells(10, 14).Value = "?"
objSheet.Cells(11, 14).Value = "?"
objSheet.Cells(12, 14).Value = "?"
objSheet.Cells(13, 14).Value = "?"
objSheet.Cells(14, 14).Value = "?"
objSheet.Cells(15, 14).Value = "?"
objSheet.Cells(16, 14).Value = "?"
objSheet.Cells(17, 14).Value = "?"
objSheet.Cells(18, 14).Value = "?"
objSheet.Cells(19, 14).Value = "?"
objSheet.Range("Q:Q").Select
objExcel.Selection.NumberFormat = "0.0000" 
objWB.Close True
objExcel.Quit