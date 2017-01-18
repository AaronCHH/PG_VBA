Attribute VB_Name = "Module2"
Sub 選取多個檔案()  
  Dim a As Integer
  With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = True
    If .Show = 0 Then
      Exit Sub
    End If
    For a = 1 To .SelectedItems.Count
      MsgBox .SelectedItems(a)
    Next
  End With
End Sub
