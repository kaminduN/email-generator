''' Import vb modules into a excel workbook

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
        MsgBox "import Folder not exist"
        WScript.Quit 
	Else
		MsgBox "Exists"
    End If
    
    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0
	
	Set objFSO = CreateObject("scripting.filesystemobject")
	''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles & "\"
	
	If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       WScript.Quit 
    End If
	
	
	CurPath = objFSO.GetAbsolutePathName(".")
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open(CurPath & "\sampledata.xlsm")
	
	If objWorkbook.VBProject.Protection = 1 Then
		MsgBox "The VBA in this workbook is protected," & _
			"not possible to export the code"
		WScript.Quit 
	Else
		MsgBox "protection ok...." & objWorkbook.VBProject.Protection
    End If
	
	Set cmpComponents = objWorkbook.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path 
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.Path
        End If
        
    Next
    
   	
	objExcel.Quit
    MsgBox "Import is ready"
	
	