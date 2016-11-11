Attribute VB_Name = "Module3"
Sub copyToAnotherSheet()

	Dim a As Range, b As Range
  
	LR = Cells(Rows.Count, 1).End(xlUp).Row
	 
	Set b = ActiveSheet.Range(Cells(2, 9), Cells(LR, 12))
	 
	Set XLSheet = ActiveWorkbook.Worksheets.Add(, ActiveSheet, 1)
	Set a = XLSheet.Range(XLSheet.Cells(2, 1), XLSheet.Cells(LR, 4))
	 
	'Set a = Worksheets("Sheet2").Range(Worksheets("Sheet2").Cells(2, 1), Worksheets("Sheet2").Cells(LR, 4))
	 
	a.Value = b.Value

End Sub
