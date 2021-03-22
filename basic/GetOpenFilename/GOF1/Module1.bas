Attribute VB_Name = "Module1"
Sub Test()
  Dim FileName As String

  Call GetName(FileName)

  If FileName = "False" Then
    MsgBox "You didn't select a fileName"
  Else
    MsgBox "You selected the file: " & FileName
  End If

End Sub

Sub GetName(FileName)
  Dim FFilt As String
  FFilt = "Text Files (*.txt), *.txt," & _
          "Space Delimited Files (*.prn), *.prn," & _
          "Comma Delimited Files (*.csv), *.csv,"
  FileName = Application.GetOpenFilename(FFilt, 1, "Choose a File")
End Sub
