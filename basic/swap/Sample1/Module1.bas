Attribute VB_Name = "Module1"
Sub Scope()
  Dim x As Integer, y As Integer
  x = 6
  y = 8
  MsgBox "Before Swap: x = " & x & ", y = " & y
  Call Switch(x, y)
  MsgBox "After Swap: x = " & x & ", y = " & y
End Sub

Sub Switch(a, b)
  Dim temp As Integer
  temp = a
  a = b
  b = temp
End Sub
