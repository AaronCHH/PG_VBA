Attribute VB_Name = "Module2"
Sub ����ɮ�2()
  With Application.FileDialog(msoFileDialogFilePicker)    
    If .Show = 0 Then
      Exit Sub
    End If
    MsgBox .SelectedItems(1)
  End With
End Sub

