Attribute VB_Name = "Module3"
Sub copyToAnotherSheet()

Dim a As Range, b As Range
  

LR = Cells(Rows.Count, 1).End(xlUp).Row

 
 Set b = ActiveSheet.Range(Cells(2, 9), Cells(LR, 12))
 
 Set xlSheetN = ActiveWorkbook.Worksheets.Add(, ActiveSheet, 1)
 Set a = xlSheetN.Range(xlSheetN.Cells(2, 1), xlSheetN.Cells(LR, 4))
 
 
 
 a.Value = b.Value

xlSheetN.Cells(1, 1) = "email address"
xlSheetN.Cells(1, 2) = "first name"
xlSheetN.Cells(1, 3) = "last name"
xlSheetN.Cells(1, 4) = "password"


End Sub


