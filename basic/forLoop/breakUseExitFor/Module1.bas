Attribute VB_Name = "Module1"
Sub break()
  For i = 1 To 10
    If i = 5 Then
      Exit For
    Else
      Debug.Print (i)
    End If
  Next
End Sub
