Attribute VB_Name = "Module1"
Sub �妸�ץX���P�ɮ׮榡()
  '// �ܼƫŧi
  Dim sht_name, i
  
  '// Turn off Alert message
'  Application.DisplayAlerts = False  '***
  
  '// Change working dir
  ChDrive ActiveWorkbook.path
  ChDir ActiveWorkbook.path
  
  '// =======================================================================================
  '// Access Target File
  Filelist = Application.GetOpenFilename("csv files,*.csv", _
  MultiSelect:=True) 'NOTE: set to select single file
  
  If VarType(Filelist) = vbBoolean Then
    Exit Sub
  End If
  
  Application.EnableEvents = False  '***
  
  '// Loop opening files and save as another format
  For i = 1 To UBound(Filelist)
    Set Wbook = Workbooks.Open(Filelist(i))
    ofile_name = Replace(Wbook.Name, ".csv", "")
    MkDir ofile_name
        
    '// create save file name
'    ofile_name = Replace(Wbook.Name, ".csv", "")   '*** extension same as input files
'    Wbook.SaveAs fileName:=ofile_name & ".xlsx", FileFormat:=51, _
'          CreateBackup:=False
'    Wbook.Close
    'ActiveWorkbook.Close
  Next
    
  ' ThisWorkbook.Close

  '// =======================================================================================
  Application.EnableEvents = True   '***
  
  '// Turn on Alert message
'  Application.DisplayAlerts = True  '***
  
  '// Shot down Excel
  'Application.Quit
  
End Sub


