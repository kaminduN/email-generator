Attribute VB_Name = "Module2"
Sub Macro1(control As IRibbonControl)
    createUname
    AddUniqueRef
    createEmail
    prepareSheet
    createPw
    
    copyToAnotherSheet
    
     MsgBox "University Email address preparation is complete"
End Sub

Sub createUname()

    Dim lname As String, initials As String, uname As String
    Dim LR As Long

    LR = Cells(Rows.Count, 1).End(xlUp).Row
 
    For b = 2 To LR Step 1
    
        initials = Trim(Left(Cells(b, 4), 4))
        lname = Trim(Cells(b, 3))
        
        uname = initials & lname
     
             
        Cells(b, 8) = uname
        Cells(b, 7).FormulaR1C1 = "=SUBSTITUTE(RC[+1],"" "","""")"
    Next b
  
    
End Sub

Sub createPw()

    Dim pw As String
    Dim LR As Long

    LR = Cells(Rows.Count, 1).End(xlUp).Row
 
    For b = 2 To LR Step 1
        pw = "as" & Mid(Cells(b, 1), 5, 3) & Right(Cells(b, 2), 3)
        Cells(b, 12) = pw
         
    Next b


End Sub

Sub createEmail()

    Dim LR As Long

    LR = Cells(Rows.Count, 1).End(xlUp).Row
 
    For b = 2 To LR Step 1
    
             
        Cells(b, 9) = Cells(b, 7) & "@my.university.com"
           
    Next b
  

End Sub


Sub prepareSheet()

    LR = Cells(Rows.Count, 1).End(xlUp).Row

    Range(Cells(2, 3), Cells(LR, 4)).Select
    Selection.Copy
    Range("J2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False


End Sub

