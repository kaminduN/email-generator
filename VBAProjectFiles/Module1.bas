Attribute VB_Name = "Module1"
Option Explicit
Sub AddUniqueRef()
        
    Dim LR As Long, LC As Long, a As Long, b As Long, n As Long
    Dim c As Range
    
    'for optimisation
    Application.ScreenUpdating = False
    
    'starts from col 1 row 2
    a = 7
    
    'gets the row count for the current active col
        LR = Cells(Rows.Count, a).End(xlUp).Row
    
    'starts from the 2nd row to down
        For b = 2 To LR - 1 Step 1
            n = 0
            'checks for the duplicate values in a coloum and number them uniquily
            For Each c In Range(Cells(b + 1, a), Cells(LR, a))
                If c.Value = Cells(b, a).Value Then
                    n = n + 4
                    c = c & n
                End If
            Next c
            
        Next b
        
    'Next a
    Application.ScreenUpdating = True
End Sub
