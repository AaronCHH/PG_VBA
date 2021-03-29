Attribute VB_Name = "Module2"
'''  This development export the module into folder "VBAProjectFiles"   '''
Public Sub ExportModules()
    Dim bExport As Boolean
    ' Dim wkbSource As Excel.Workbook
    Dim wkbSource As Word.Document
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim docName As String
    ' Dim cmpComponent As VBIDE.VBComponent

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    
'    If FolderWithVBAProjectFiles = "Error" Then
'        MsgBox "Export Folder not exist"
'        Exit Sub
'    End If
'
'    On Error Resume Next
'        Kill FolderWithVBAProjectFiles & "\*.*"
'    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    ' szSourceWorkbook = ActiveWorkbook.Name
    szSourceWorkbook = ActiveDocument.Name ' ThisDocument.Name?
    ' szSourceWorkbook = ThisDocument.Name
    Set wkbSource = Application.Documents(szSourceWorkbook)
    
    docName = Split(szSourceWorkbook, ".")(0)
    'Debug.Print szSourceWorkbook
    'Debug.Print Split(szSourceWorkbook, ".")(0)
    'Debug.Print ActiveDocument.FullName
    
    'If wkbSource.VBProject.Protection = 1 Then
    'MsgBox "The VBA in this workbook is protected," & _
    '    "not possible to export the code"
    'Exit Sub
    'End If
    
    'szExportPath = FolderWithVBAProjectFiles(szSourceWorkbook) & "\"
    szExportPath = ActiveDocument.Path + "\" + docName
    Debug.Print szExportPath
    'Debug.Print ActiveDocument.Path + "\" + docName + "\"
    
    ' MsgBox to confirm
    Response = MsgBox("Sure to Export?", vbYesNo + vbQuestion)
    If Response = vbYes Then
      For Each cmpComponent In wkbSource.VBProject.VBComponents
          
          bExport = True
          
          szFileName = cmpComponent.Name
          ' Debug.Print cmpComponent.Type
  
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
              On Error Resume Next
              MkDir szExportPath
              On Error GoTo 0
              
              cmpComponent.Export szExportPath + "\" + szFileName + ".bas"
              'Debug.Print cmpComponent.Name
              'Debug.Print szExportPath
              'Debug.Print szFileName
                            
          ''' remove it from the project if you want
          '''wkbSource.VBProject.VBComponents.Remove cmpComponent
          
          End If
     
      Next cmpComponent
  
      MsgBox "Export is complete"
      
    Else
    
      MsgBox "Export is abort"
      
    End If

End Sub

