Attribute VB_Name = "Module1"
Sub testGoto()
  x = InputBox("enter a value")
  GoTo 1
  MsgBox x
1:   c = x + 8
  MsgBox c
End Sub
