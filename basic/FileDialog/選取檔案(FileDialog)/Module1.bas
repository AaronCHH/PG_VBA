Attribute VB_Name = "Module1"
Sub ¿ï¨úÀÉ®×()
  Dim Selection As Integer
  With Application.FileDialog(msoFileDialogFilePicker)
    Selection = .Show
    If Selection = 0 Then
      Exit Sub
    End If
    MsgBox .SelectedItems(1)
  End With
End Sub

