Attribute VB_Name = "Module1"
Sub GetOpenFilename����h���ɮ�()
    Dim FileList As Variant
    FileList = Application.GetOpenFilename(MultiSelect:=True)
    If VarType(FileList) = vbBoolean Then
        Exit Sub
    End If
    For i = 1 To UBound(FileList)
        MsgBox FileList(i)
    Next
End Sub
