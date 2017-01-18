Attribute VB_Name = "Module1"
Sub GetOpenFilename()
    Dim FileName As Variant
    FileName = Application.GetOpenFilename()
    If FileName = False Then
        Exit Sub
    End If
    MsgBox FileName
End Sub

