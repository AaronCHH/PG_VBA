Sub ImportModules_Dev()
  ' IMPORT MODULE - https://stackoverflow.com/questions/11526577/loop-through-all-word-files-in-directory
  'ActiveDocument.VBProject.VBComponents.Import ThisDocument.Path & "\TEST.bas"
  'MsgBox ThisDocument.Path
  
  ' FIND THE PATH OF ACTIVATE DOCUMENT
  szSourceWorkbook = ActiveDocument.Name ' ThisDocument.Name?
  Set wkbSource = Application.Documents(szSourceWorkbook)
  docName = Split(szSourceWorkbook, ".")(0)
  szImportPath = ActiveDocument.Path + "\" + docName
  
  MsgBox szImportPath
  
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  ' Debug.Print objFSO.GetFolder(szImportPath)
  
  For Each objFile In objFSO.GetFolder(szImportPath).Files
    'Debug.Print objFile
    'Debug.Print objFile.Path
      
  '  If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
  '     (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
  '     (objFSO.GetExtensionName(objFile.Name) = "bas") Then
  '     'cmpComponents.Import objFile.Path
  '    Debug.Print objFile.Path
  '  End If
          
  Next objFile
 
End Sub
' Refs: https://stackoverflow.com/questions/11526577/loop-through-all-word-files-in-directory



Sub RemoveCode()
  Dim doc As Document
  Dim vbc As Object
  Set doc = ActiveDocument
  For Each vbc In doc.VBProject.VBComponents
    Debug.Print vbc.Name
    Select Case vbc.Name
      Case "ThisDocument"
        vbc.CodeModule.DeleteLines 1, vbc.CodeModule.CountOfLines
      Case Else
        doc.VBProject.VBComponents.Remove vbc
    End Select
  Next vbc
End Sub
' Refs: https://answers.microsoft.com/en-us/msoffice/forum/msoffice_word-mso_other-mso_archive/how-to-delete-a-word-module-using-vba/95ca58e7-4ccd-4882-8d43-f1bead011082
