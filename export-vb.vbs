''' Export vb modules out of an Excel file

Function FolderWithVBAProjectFiles()
    Dim WshShell
    Dim FSO 
    Dim SpecialPath 

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")

    ''' Dim SpecialPath
    SpecialPath = FSO.GetAbsolutePathName(".")
    
    '''SpecialPath = WshShell.SpecialFolders("MyDocuments")

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    
    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
        On Error Resume Next
        MkDir SpecialPath & "VBAProjectFiles"
        On Error GoTo 0
    End If
    
    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
        FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function

Const vbext_ct_ClassModule = 2
Const vbext_ct_Document = 100
Const vbext_ct_MSForm = 3
Const vbext_ct_StdModule = 1

	
	''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        WScript.Quit 
    End If
    
    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0
	
	Set FSO = CreateObject("scripting.filesystemobject")
	CurPath = FSO.GetAbsolutePathName(".")
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open(CurPath & "\sampledata.xlsm")
	
	If objWorkbook.VBProject.Protection = 1 Then
		MsgBox "The VBA in this workbook is protected," & _
			"not possible to export the code"
		WScript.Quit 
	'''Else
		'''MsgBox "protection ok...." & objWorkbook.VBProject.Protection
    End If
	
	
	szExportPath = FolderWithVBAProjectFiles & "\"
	
	For Each cmpComponent In objWorkbook.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name
		'''MsgBox szFileName & "  " & cmpComponent.Type
        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next 

	objExcel.Quit
    MsgBox "Export complete"
	
	