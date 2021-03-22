Attribute VB_Name = "Module1"
Sub continue()
  For i = 1 To 10
    If i = 3 Then
      GoTo 1
    Else
      Debug.Print (i)
    End If
1:
  Next
End Sub
